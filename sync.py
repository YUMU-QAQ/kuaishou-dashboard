"""每日拉数任务。cron 每日 09:01 调用，也可手动跑。

流程：刷新各账号 token → 逐个账号拉数（首次回溯 90 天，之后回补 t-1~t-7）→ 断链告警。

用法：
    python sync.py          # 例行
    python sync.py --force-full   # 强制对全部账号重新回溯 90 天
"""
import sys
import time
from datetime import date, timedelta

import db
import kuaishou

BACKFILL_DAYS = 7     # 首次回溯（实时报表最大跨度 7 天）
ROLLING_DAYS = 7     # 每日回补 t-1 ~ t-7
RETRY = 2            # 拉数失败重试次数


def recent_dates(n):
    """最近 n 天（不含今天），从最近到最远。"""
    today = date.today()
    return [(today - timedelta(days=i)).isoformat() for i in range(1, n + 1)]


def refresh_all():
    """按 (所属应用, refresh_token) 分组，每组只 refresh 一次，结果广播到组内所有账号。
    同一 refresh_token 只属于一个应用，分组时带上 app_id 才不会拿错 secret 去刷。
    返回 {advertiser_id: access_token}。"""
    groups = {}
    for acc in db.list_accounts():
        t = db.get_tokens(acc["advertiser_id"])
        if t and t.get("refresh_token"):
            key = (acc.get("app_id") or "", t["refresh_token"])
            groups.setdefault(key, []).append(acc["advertiser_id"])

    result = {}
    for (app_id, rt), aids in groups.items():
        api = kuaishou.api_for(app_id)
        try:
            resp = api.refresh(rt)
            if not api.ok(resp):
                raise RuntimeError(f"refresh 失败: {resp.get('code')} {resp.get('message')}")
            data = resp.get("data") or {}
            at = data.get("access_token")
            nrt = data.get("refresh_token")
            if not at:
                raise RuntimeError("refresh 响应缺少 access_token")
            for aid in aids:
                db.save_tokens(aid, at, kuaishou.token_expiry(data, "access_token_expires_in"),
                               nrt, kuaishou.token_expiry(data, "refresh_token_expires_in"))
                result[aid] = at
        except Exception:
            for aid in aids:
                db.set_account(aid, token_status="broken")
                result[aid] = None
    return result


def fetch_account(api, access_token, advertiser_id, start_date, end_date):
    """拉全站数据（账户级日粒度），返回 {stat_date: 指标dict}。"""
    for _ in range(RETRY + 1):
        try:
            resp = api.storewide_report(access_token, advertiser_id, start_date, end_date)
            if not api.ok(resp):
                raise RuntimeError(f"storewideReport: {resp.get('code')} {resp.get('message')}")
            return kuaishou.parse_storewide(resp)
        except Exception:
            if _ == RETRY:
                raise
            time.sleep(1.0)


def fetch_campaign(api, access_token, advertiser_id, start_date, end_date, account_host=""):
    """拉计划维度数据，返回 {stat_date: {campaign_id: {name,host,category,charge,gmv,roi}}}。"""
    name_map = {}
    try:
        name_map = api.ad_search_campaigns(access_token, advertiser_id, start_date, end_date)
    except Exception:
        pass  # 名字拿不到不致命，数据照存
    resp = api.storewide_report(access_token, advertiser_id, start_date, end_date,
                                view_type=3000, select_columns=kuaishou.STOREWIDE_CAMPAIGN_COLUMNS)
    if not api.ok(resp):
        raise RuntimeError(f"storewideReport(计划): {resp.get('code')} {resp.get('message')}")
    return kuaishou.parse_storewide_campaign(resp, name_map, account_host)


def fetch_photo(api, access_token, advertiser_id, start_date, end_date):
    """拉视频维度数据，返回 {stat_date: {photo_id: {name,is_ai,charge,gmv,roi}}}。"""
    import datetime

    def _ms(d, plus_days=0):
        dt = (datetime.datetime.strptime(d, "%Y-%m-%d") + datetime.timedelta(days=plus_days))
        dt = dt.replace(tzinfo=datetime.timezone(datetime.timedelta(hours=8)))
        return int(dt.timestamp() * 1000)

    info = {}
    try:
        # 视频库按创建时间筛，跨度上限 94 天，取 end_date 前 90 天拿全 photo_id -> 名称/播放地址
        info = api.filemanage_video_list(access_token, advertiser_id,
                                         _ms(end_date, plus_days=-90), _ms(end_date, plus_days=1))
    except Exception:
        pass  # 名字/地址拿不到不致命，用 photo_id 兜底
    for pid, x in info.items():
        db.upsert_photo_url(pid, x.get("url"), x.get("cover"))
    name_map = {pid: x["name"] for pid, x in info.items()}
    # 罗盘导出的真名优先（覆盖 MAPI 视频库小店铺数字ID），未命中回退 filemanage 名字
    name_map = {**name_map, **db.query_photo_map()}
    resp = api.storewide_report(access_token, advertiser_id, start_date, end_date,
                                view_type=3000, select_columns=kuaishou.STOREWIDE_PHOTO_COLUMNS)
    if not api.ok(resp):
        raise RuntimeError(f"storewideReport(视频): {resp.get('code')} {resp.get('message')}")
    return kuaishou.parse_storewide_photo(resp, name_map)


def sync_one(api, advertiser_id, access_token, force_full=False):
    if force_full or not db.has_daily(advertiser_id):
        days = recent_dates(BACKFILL_DAYS)
    else:
        days = recent_dates(ROLLING_DAYS)
    start_date, end_date = days[-1], days[0]

    rows = fetch_account(api, access_token, advertiser_id, start_date, end_date)
    # 3秒播放率/完播率：全站接口不提供，只有 v1/report/account_report 有。失败不致命，别拖垮整账号。
    # 只覆盖 storewide 已有数据的日期——否则会把 charge/gmv 写成 NULL。
    try:
        play = kuaishou.parse_account_report(
            api.account_report(access_token, advertiser_id, start_date, end_date))
        for stat_date, d in play.items():
            if stat_date in rows:
                rows[stat_date].update(d)
    except Exception as e:
        print(f"{advertiser_id}: 播放率接口失败，本次跳过（{e}）")
    for stat_date, d in rows.items():
        db.upsert_daily(advertiser_id, stat_date, d)

    # 计划维度：ad/search 拿计划名 + storewide 计划维度拿花费/GMV，主播认不出用账户名兜底
    account = next((a for a in db.list_accounts() if a["advertiser_id"] == advertiser_id), {})
    account_host = kuaishou.ACCOUNT_HOST.get(account.get("name") or "", "")
    camps = fetch_campaign(api, access_token, advertiser_id, start_date, end_date, account_host)
    n_camp = 0
    for stat_date, by_cid in camps.items():
        for cid, c in by_cid.items():
            db.upsert_campaign(advertiser_id, cid, c["name"], c["host"], c["category"],
                               stat_date, c["charge"], c["gmv"], c["roi"])
            n_camp += 1

    # 视频维度：filemanage 拿视频名 + storewide 视频维度拿花费/GMV，AI 判定落库
    photos = fetch_photo(api, access_token, advertiser_id, start_date, end_date)
    n_photo = 0
    for stat_date, by_pid in photos.items():
        for pid, p in by_pid.items():
            db.upsert_photo(advertiser_id, pid, p["name"], p["is_ai"], stat_date,
                            p["charge"], p["gmv"], p["roi"], p.get("campaign_id"))
            n_photo += 1

    db.set_account(advertiser_id, token_status="ok", last_sync_at=time.time())
    return f"ok:{len(rows)}行:{n_camp}计划:{n_photo}视频"


def alert(webhook, text):
    if not webhook:
        return
    import requests
    requests.post(webhook, json={"msgtype": "text", "text": {"content": text}}, timeout=10)


def main():
    db.init()
    kuaishou.load_env()
    cfg = kuaishou.config()
    if not kuaishou.apps():
        print("缺少应用配置：请在 .env 里填 APPS=app_id:secret[,...]（或 APP_ID/APP_SECRET）", file=sys.stderr)
        sys.exit(1)

    force_full = "--force-full" in sys.argv

    accounts = db.list_accounts()
    if not accounts:
        print("还没有已授权账号。先访问 /auth 页完成授权。")
        return

    tokens_by_aid = refresh_all()
    broken, failed = [], []
    for acc in accounts:
        aid = acc["advertiser_id"]
        at = tokens_by_aid.get(aid)
        if not at:
            broken.append(aid)
            print(f"{aid}: token 失败")
            continue
        # 每个账号用自己的应用（资质不同，应用不同）
        try:
            result = sync_one(kuaishou.api_for(acc.get("app_id")), aid, at, force_full)
            print(f"{aid}: {result}")
        except Exception as e:
            # 单个账号失败不能拖垮整轮同步（否则后面的账号永远轮不到）
            msg = str(e)
            if "不在token账户管理下" in msg:
                msg = "授权已失效，需该店铺重新授权（该账号已不在 token 授权范围内）"
            failed.append(f"{aid}({msg})")
            db.set_account(aid, token_status="broken")
            print(f"{aid}: 失败 {msg}")

    if broken or failed:
        names = ", ".join(broken + failed)
        alert(cfg["dingtalk_webhook"], f"快手看板：以下账号同步失败，需人工处理：{names}")
    print("完成。断链账号：", broken if broken else "无")
    if failed:
        print("同步失败账号：", "; ".join(failed))


if __name__ == "__main__":
    main()
