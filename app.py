"""Web 服务：看板、授权管理、OAuth 回调、导出。

启动：uvicorn app:app --host 0.0.0.0 --port 8080
"""
import csv
import io
import json
import secrets
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

import db
import kuaishou
import sync
from fastapi import Depends, FastAPI, HTTPException, Query
from fastapi.responses import FileResponse, RedirectResponse, Response
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from pydantic import BaseModel
from openpyxl import Workbook

BASE = Path(__file__).resolve().parent
app = FastAPI()
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


def _api():
    cfg = kuaishou.config()
    return kuaishou.KuaishouAPI(cfg["app_id"], cfg["secret"], cfg["redirect_uri"])


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
    return {"rows": db.query_photo(start, end), "accounts": db.list_accounts()}


@app.get("/api/video-url", dependencies=[Depends(check_auth)])
def api_video_url(advertiser_id: str = Query(...), photo_id: str = Query(...)):
    """按需取/刷新单个视频的播放地址（库里存的可能过期）。"""
    t = db.get_tokens(advertiser_id)
    if not t or not t.get("access_token"):
        raise HTTPException(400, "该账号未授权或 token 缺失，请先授权并同步")
    api = _api()
    info = api.filemanage_video_list(t["access_token"], advertiser_id, photo_id=photo_id)
    if not info:  # token 可能过期，刷新后重试一次
        toks = sync.refresh_all(api)
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
def oauth_start():
    state = secrets.token_urlsafe(16)
    db.save_state(state)
    return RedirectResponse(_api().authorize_url(state))


@app.get("/api/authorize-url", dependencies=[Depends(check_auth)])
def api_authorize_url():
    """生成一次性的授权链接，返回给前端复制（发给店铺持有人自己登录授权）。"""
    state = secrets.token_urlsafe(16)
    db.save_state(state)
    return {"url": _api().authorize_url(state)}


@app.get("/oauth/callback")
def oauth_callback(auth_code: str = Query(...), state: str = Query("")):
    if not state or not db.pop_state(state):
        raise HTTPException(400, "无效的 state，请重新发起授权")
    api = _api()
    resp = api.get_token(auth_code)
    if not api.ok(resp):
        raise HTTPException(400, f"换取 token 失败: {resp.get('message')}")
    data = resp.get("data") or resp
    access_token = data.get("access_token")
    refresh_token = data.get("refresh_token")
    if not access_token:
        raise HTTPException(400, f"token响应缺字段: {json.dumps(resp, ensure_ascii=False)}")
    # 优先用 token 里的 advertiser_id；为 0 时退回 approval/list
    print(f"[oauth] token响应: {json.dumps(resp, ensure_ascii=False)}", flush=True)
    aid = data.get("advertiser_id") or 0
    if aid:
        ids = [str(aid)]
    else:
        ids = kuaishou.parse_approval_advertiser_ids(api.approval_list(access_token))
    if not ids:
        raise HTTPException(400, f"无法确定账户。token响应: {json.dumps(resp, ensure_ascii=False)}")
    for x in ids:
        db.upsert_account(x)
        db.save_tokens(x, access_token, kuaishou.token_expiry(data, "access_token_expires_in"),
                       refresh_token, kuaishou.token_expiry(data, "refresh_token_expires_in"))
        db.set_account(x, token_status="ok")
    return RedirectResponse("/auth")


if __name__ == "__main__":
    import uvicorn
    kuaishou.load_env()
    db.init()
    uvicorn.run(app, host="0.0.0.0", port=8080)
