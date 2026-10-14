"""Web 服务：看板、授权管理、OAuth 回调、导出。

启动：uvicorn app:app --host 0.0.0.0 --port 8080
"""
import csv
import io
import json
import secrets
import threading
import time
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

import db
import kuaishou
import sync
from fastapi import Depends, FastAPI, HTTPException, Query
from fastapi.middleware.gzip import GZipMiddleware
from fastapi.responses import FileResponse, RedirectResponse, Response
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from pydantic import BaseModel
from openpyxl import Workbook

BASE = Path(__file__).resolve().parent
app = FastAPI()
# 接口响应动辄几 MB（/api/photos 一个月近万行），压缩后小 5 倍左右，加载快很多
app.add_middleware(GZipMiddleware, minimum_size=500)
security = HTTPBasic(auto_error=False)

kuaishou.load_env()
db.init()

CSV_COLUMNS = [
    ("advertiser_id", "账号ID"),
    ("stat_date", "日期"),
    ("charge", "花费(元)"),
    ("show", "曝光数"),
    ("gmv", "全站GMV"),
    ("roi", "全站ROI"),
]


def check_auth(credentials: Optional[HTTPBasicCredentials] = Depends(security)):
    cfg = kuaishou.config()
    if not cfg["auth_password"]:
        return True  # 未配置密码 = 开发期不校验
    if credentials is None or not (
        secrets.compare_digest(credentials.username, cfg["auth_user"])
        and secrets.compare_digest(credentials.password, cfg["auth_password"])
    ):
        raise HTTPException(401, "未授权", headers={"WWW-Authenticate": "Basic"})
    return True


def _api(app_id=None):
    """按 app_id 取 MAPI 客户端；None/未知则用配置里的第一个应用。"""
    return kuaishou.api_for(app_id)


def _range(start: Optional[str], end: Optional[str]):
    if not end:
        end = date.today().isoformat()
    if not start:
        start = (date.today() - timedelta(days=30)).isoformat()
    return start, end


# ---- 页面 ----

@app.get("/", dependencies=[Depends(check_auth)], response_class=FileResponse)
def dashboard():
    return FileResponse(BASE / "templates" / "dashboard.html")


@app.get("/auth", dependencies=[Depends(check_auth)], response_class=FileResponse)
def auth_page():
    return FileResponse(BASE / "templates" / "auth.html")


# ---- 数据 ----

@app.get("/api/accounts", dependencies=[Depends(check_auth)])
def api_accounts():
    return db.list_accounts()


@app.get("/api/apps", dependencies=[Depends(check_auth)])
def api_apps():
    """已配置的 MAPI 应用 ID 列表，供授权页选择（不同资质的店铺要用对应应用）。"""
    return {"apps": list(kuaishou.apps().keys())}


# ---- 自定义分组 ----

def _per_account_metrics(start, end):
    """区间内按账号汇总，供分组汇总用。比率给的是分子分母（要加总后再除，不能平均比率）。
    播放率取 daily 表（sync 用 v1/report/account_report 写入），不再依赖罗盘明细。"""
    def blank():
        return {"gmv": 0.0, "show": 0, "charge": 0.0, "videos": 0,
                "played_num": 0.0, "played_3s": 0.0, "played_end": 0.0, "has_play": False}

    per = {}
    for r in db.query_daily(start, end):
        a = per.setdefault(r["advertiser_id"], blank())
        a["gmv"] += r["gmv"] or 0
        a["show"] += r["show"] or 0
        a["charge"] += r["charge"] or 0
        if r.get("played_num"):
            a["played_num"] += r["played_num"]
            a["played_3s"] += r["played_3s"] or 0
            a["played_end"] += r["played_end"] or 0
            a["has_play"] = True
    for aid, n in db.query_video_counts(start, end).items():
        per.setdefault(aid, blank())["videos"] = n
    return per


def _sum_group(per, accounts):
    m = {"gmv": 0.0, "show": 0, "charge": 0.0, "videos": 0,
         "played_num": 0.0, "played_3s": 0.0, "played_end": 0.0}
    has_play = False
    for aid in accounts:
        d = per.get(aid)
        if not d:
            continue
        for k in m:
            m[k] += d.get(k) or 0
        has_play = has_play or bool(d.get("has_play"))

    def rate(num):
        return round(num / m["played_num"] * 100, 2) if (has_play and m["played_num"]) else None

    return {
        "gmv": round(m["gmv"], 2), "show": m["show"], "charge": round(m["charge"], 2),
        "videos": m["videos"],
        "avg_gmv": round(m["gmv"] / m["videos"], 2) if m["videos"] else None,
        "r3": rate(m["played_3s"]),
        "rend": rate(m["played_end"]),
    }


@app.get("/api/groups", dependencies=[Depends(check_auth)])
def api_groups(start: Optional[str] = Query(None), end: Optional[str] = Query(None)):
    start, end = _range(start, end)
    per = _per_account_metrics(start, end)
    return {"start": start, "end": end,
            "groups": [{**g, "metrics": _sum_group(per, g["accounts"])} for g in db.list_groups()]}


class GroupBody(BaseModel):
    name: str
    advertiser_ids: list[str] = []


@app.post("/api/groups", dependencies=[Depends(check_auth)])
def create_group(body: GroupBody):
    name = (body.name or "").strip()
    if not name:
        raise HTTPException(400, "分组名不能为空")
    gid = db.create_group(name)
    if gid is None:
        raise HTTPException(409, f"已存在同名分组：{name}")
    db.set_group_accounts(gid, body.advertiser_ids)
    return {"id": gid}


@app.put("/api/groups/{group_id}", dependencies=[Depends(check_auth)])
def update_group(group_id: int, body: GroupBody):
    """改名 + 重设成员，一次到位（弹窗里的「保存」）。"""
    name = (body.name or "").strip()
    if not name:
        raise HTTPException(400, "分组名不能为空")
    if not db.rename_group(group_id, name):
        raise HTTPException(409, f"已存在同名分组：{name}")
    db.set_group_accounts(group_id, body.advertiser_ids)
    return {"ok": True}


@app.delete("/api/groups/{group_id}", dependencies=[Depends(check_auth)])
def delete_group(group_id: int):
    db.delete_group(group_id)
    return {"ok": True}


@app.get("/api/data", dependencies=[Depends(check_auth)])
def api_data(start: Optional[str] = Query(None), end: Optional[str] = Query(None)):
    start, end = _range(start, end)
    return {"rows": db.query_daily(start, end), "accounts": db.list_accounts()}


@app.get("/api/plans", dependencies=[Depends(check_auth)])
def api_plans(start: Optional[str] = Query(None), end: Optional[str] = Query(None)):
    start, end = _range(start, end)
    return {"rows": db.query_campaign(start, end), "accounts": db.list_accounts()}


@app.get("/api/photos", dependencies=[Depends(check_auth)])
def api_photos(start: Optional[str] = Query(None), end: Optional[str] = Query(None)):
    start, end = _range(start, end)
    rows = db.query_photo(start, end)
    # 播放地址按 photo_id 去重单独给一份：同一视频多天共用一条，避免每行都塞一遍长 URL
    urls = db.query_photo_url_map()
    pids = {r["photo_id"] for r in rows}
    return {"rows": rows, "accounts": db.list_accounts(),
            "urls": {p: u["url"] for p, u in urls.items() if p in pids and u.get("url")}}


@app.get("/api/video-url", dependencies=[Depends(check_auth)])
def api_video_url(advertiser_id: str = Query(...), photo_id: str = Query(...)):
    """按需取/刷新单个视频的播放地址（库里存的可能过期）。"""
    t = db.get_tokens(advertiser_id)
    if not t or not t.get("access_token"):
        raise HTTPException(400, "该账号未授权或 token 缺失，请先授权并同步")
    acc = db.get_account(advertiser_id) or {}
    api = _api(acc.get("app_id"))
    info = api.filemanage_video_list(t["access_token"], advertiser_id, photo_id=photo_id)
    if not info:  # token 可能过期，刷新后重试一次
        toks = sync.refresh_all()
        if toks.get(advertiser_id):
            info = api.filemanage_video_list(toks[advertiser_id], advertiser_id, photo_id=photo_id)
    hit = info.get(str(photo_id))
    if not hit or not hit.get("url"):
        raise HTTPException(404, "没取到播放地址（视频可能已删除或不在视频库）")
    db.upsert_photo_url(str(photo_id), hit["url"], hit["cover"])
    return {"url": hit["url"], "cover_url": hit["cover"]}


class AccountUpdate(BaseModel):
    name: Optional[str] = None
    subject: Optional[str] = None


@app.post("/api/accounts/{advertiser_id}", dependencies=[Depends(check_auth)])
def update_account(advertiser_id: str, body: AccountUpdate):
    db.set_account(advertiser_id, name=body.name, subject=body.subject)
    return {"ok": True}


# ---- 手动触发同步 ----
# uvicorn 单进程，模块级状态够用；同一时刻只允许一个同步在跑，避免并发写库。
_sync_state = {"running": False, "started_at": None, "finished_at": None, "error": None}


def _run_sync():
    try:
        sync.main()
    except Exception as e:
        _sync_state["error"] = str(e)
    finally:
        _sync_state["running"] = False
        _sync_state["finished_at"] = time.time()


@app.post("/api/sync", dependencies=[Depends(check_auth)])
def api_sync():
    """手动跑一次同步（与每日定时任务同一个流程）。耗时约一分钟，前端轮询状态。"""
    if _sync_state["running"]:
        raise HTTPException(409, "已有同步在进行中，请稍候")
    _sync_state.update(running=True, started_at=time.time(), finished_at=None, error=None)
    threading.Thread(target=_run_sync, daemon=True).start()
    return {"started": True}


@app.get("/api/sync/status", dependencies=[Depends(check_auth)])
def api_sync_status():
    return _sync_state


@app.delete("/api/accounts/{advertiser_id}", dependencies=[Depends(check_auth)])
def delete_account(advertiser_id: str):
    db.delete_account(advertiser_id)
    return {"ok": True}


# ---- 导出 ----

@app.get("/api/export.csv", dependencies=[Depends(check_auth)])
def export_csv(start: Optional[str] = Query(None), end: Optional[str] = Query(None)):
    start, end = _range(start, end)
    rows = db.query_daily(start, end)
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow([name for _, name in CSV_COLUMNS])
    for r in rows:
        w.writerow([r[key] for key, _ in CSV_COLUMNS])
    return Response(
        buf.getvalue(),
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename=kuaishou_{start}_{end}.csv"},
    )


@app.get("/api/export.xlsx", dependencies=[Depends(check_auth)])
def export_xlsx(start: Optional[str] = Query(None), end: Optional[str] = Query(None)):
    start, end = _range(start, end)
    rows = db.query_daily(start, end)
    wb = Workbook()
    ws = wb.active
    ws.append([name for _, name in CSV_COLUMNS])
    for r in rows:
        ws.append([r[key] for key, _ in CSV_COLUMNS])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return Response(
        buf.read(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=kuaishou_{start}_{end}.xlsx"},
    )


# ---- 视频明细导出 ----

PHOTO_COLUMNS = [
    ("advertiser_id", "账号ID"),
    ("account_name", "账号名称"),
    ("stat_date", "日期"),
    ("photo_id", "视频ID"),
    ("photo_name", "视频名称"),
    ("kind", "类型"),
    ("charge", "花费(元)"),
    ("show", "曝光数"),
    ("gmv", "全站GMV"),
    ("roi", "全站ROI"),
    ("gmv3", "3日累计GMV"),
    ("r3", "3s完播率(%)"),
]


def _photo_export_rows(start, end):
    names = {a["advertiser_id"]: (a["name"] or a["advertiser_id"]) for a in db.list_accounts()}
    out = []
    for r in db.query_photo(start, end):
        exp = r.get("exposure")
        out.append({
            "advertiser_id": r["advertiser_id"],
            "account_name": names.get(r["advertiser_id"], r["advertiser_id"]),
            "stat_date": r["stat_date"],
            "photo_id": r["photo_id"],
            "photo_name": r["photo_name"],
            "kind": "AI" if r["is_ai"] else "手剪",
            "charge": r["charge"],
            "show": r["show"] if r.get("show") is not None else "",
            "gmv": r["gmv"],
            "roi": r["roi"],
            "gmv3": r["gmv3"] if r.get("gmv3") is not None else "",
            "r3": round(r["r3n"] / exp, 2) if (r.get("r3n") is not None and exp) else "",
        })
    return out


@app.get("/api/export-photos.csv", dependencies=[Depends(check_auth)])
def export_photos_csv(start: Optional[str] = Query(None), end: Optional[str] = Query(None)):
    start, end = _range(start, end)
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow([name for _, name in PHOTO_COLUMNS])
    for r in _photo_export_rows(start, end):
        w.writerow([r[key] for key, _ in PHOTO_COLUMNS])
    return Response(
        buf.getvalue(),
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename=kuaishou_photos_{start}_{end}.csv"},
    )


@app.get("/api/export-photos.xlsx", dependencies=[Depends(check_auth)])
def export_photos_xlsx(start: Optional[str] = Query(None), end: Optional[str] = Query(None)):
    start, end = _range(start, end)
    wb = Workbook()
    ws = wb.active
    ws.title = "视频明细"
    ws.append([name for _, name in PHOTO_COLUMNS])
    for r in _photo_export_rows(start, end):
        ws.append([r[key] for key, _ in PHOTO_COLUMNS])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return Response(
        buf.read(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=kuaishou_photos_{start}_{end}.xlsx"},
    )


# ---- OAuth ----

@app.get("/oauth/start", dependencies=[Depends(check_auth)])
def oauth_start(app: Optional[str] = Query(None)):
    state = secrets.token_urlsafe(16)
    db.save_state(state, app)
    return RedirectResponse(_api(app).authorize_url(state))


@app.get("/api/authorize-url", dependencies=[Depends(check_auth)])
def api_authorize_url(app: Optional[str] = Query(None)):
    """生成一次性的授权链接，返回给前端复制（发给店铺持有人自己登录授权）。
    app 为要使用的 MAPI 应用 ID，省略则用配置里的第一个。"""
    state = secrets.token_urlsafe(16)
    db.save_state(state, app)
    return {"url": _api(app).authorize_url(state)}


@app.get("/oauth/callback")
def oauth_callback(auth_code: str = Query(...), state: str = Query("")):
    app_id = db.pop_state(state) if state else None
    if app_id is None:
        raise HTTPException(400, "无效的 state，请重新发起授权")
    api = _api(app_id)
    resp = api.get_token(auth_code)
    if not api.ok(resp):
        raise HTTPException(400, f"换取 token 失败: {resp.get('message')}")
    data = resp.get("data") or resp
    access_token = data.get("access_token")
    refresh_token = data.get("refresh_token")
    if not access_token:
        raise HTTPException(400, f"token响应缺字段: {json.dumps(resp, ensure_ascii=False)}")
    # 优先用 token 里的 advertiser_id；为 0 时退回 approval/list
    print(f"[oauth] app={api.app_id} token响应: {json.dumps(resp, ensure_ascii=False)}", flush=True)
    aid = data.get("advertiser_id") or 0
    if aid:
        ids = [str(aid)]
    else:
        ids = kuaishou.parse_approval_advertiser_ids(api.approval_list(access_token))
    if not ids:
        raise HTTPException(400, f"无法确定账户。token响应: {json.dumps(resp, ensure_ascii=False)}")
    for x in ids:
        if api.app_id in kuaishou.ENGINE_APPS:
            # 磁力引擎应用只补 3秒播放率/完播率，单独存，别覆盖店铺的金牛应用绑定
            db.save_engine_tokens(x, access_token,
                                  kuaishou.token_expiry(data, "access_token_expires_in"),
                                  refresh_token, kuaishou.token_expiry(data, "refresh_token_expires_in"))
            continue
        db.upsert_account(x, app_id=api.app_id)
        db.save_tokens(x, access_token, kuaishou.token_expiry(data, "access_token_expires_in"),
                       refresh_token, kuaishou.token_expiry(data, "refresh_token_expires_in"))
        db.set_account(x, token_status="ok")
    return RedirectResponse("/auth")


if __name__ == "__main__":
    import uvicorn
    kuaishou.load_env()
    db.init()
    uvicorn.run(app, host="0.0.0.0", port=8080)
