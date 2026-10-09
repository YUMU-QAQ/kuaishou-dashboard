"""快手 MAPI 客户端。

字段名与响应结构来自官方 Go SDK（github.com/bububa/kwai-marketing-api），
部分未获官方确认，接真接口后如字段不符，改本文件顶部常量即可，别散改业务代码。
"""
import json
import os
import re
import time
from urllib.parse import quote

import requests

GATEWAY = "https://ad.e.kuaishou.com/rest/openapi"
AUTHORIZE_URL = "https://developers.e.kuaishou.com/tools/authorize"
SCOPE = [
    "esp_ad_query",
    "esp_ad_manage",
    "esp_report_service",
    "esp_account_service",
    "public_dmp_service",
    "public_agent_service",
    "public_account_service",
]

# storewideReport 金额单位是 1/1000 元（3608600 → 3608.60 元），ROI 用两字段相除（单位抵消）。
STOREWIDE_SELECT_COLUMNS = ["t0_order_payment_amt", "cost_total", "show_cnt", "report_date_str"]
STOREWIDE_CAMPAIGN_COLUMNS = ["campaign_id", "t0_order_payment_amt", "cost_total", "report_date_str"]
STOREWIDE_PHOTO_COLUMNS = ["photo_id", "campaign_id", "t0_order_payment_amt", "cost_total", "report_date_str"]


def is_ai(title):
    """AI 视频判定（罗盘口径）：标题末尾 A，或含 j/J 紧贴话题块（后跟空白+#）之前。"""
    t = (title or "").replace("\xa0", " ").strip()
    return t.endswith("A") or re.search(r"[jJ](?=\s*#)", t) is not None

# 计划名里的主播关键词（罗盘口径：先认主播，认不出用账户名兜底）
HOSTS = ("泽鹏", "花花", "小琦琦")
ACCOUNT_HOST = {"四时颐养堂": "泽鹏", "花花好物清单": "花花"}
_DATE_RE = re.compile(r"^\d+(\.\d+)?$")


def parse_campaign(name):
    """'9.7-树德堂阿胶浆-泽鹏【目标投产4.8】' -> ('树德堂阿胶浆', '泽鹏')"""
    tokens = [t for t in re.sub(r"【[^】]*】", "", name).split("-") if t.strip()]
    if tokens and _DATE_RE.match(tokens[0]):
        tokens = tokens[1:]
    host, rest = "", []
    for t in tokens:
        hit = next((h for h in HOSTS if h in t), None)
        if hit and not host:
            host = hit
            t = t.replace(hit, "")
        if t.strip():
            rest.append(t.strip())
    category = re.sub(r"\d+$", "", "-".join(rest)).strip("-_ ") or "未分类"
    return category, host


def load_env(path=".env"):
    """手动读 .env，不引入 python-dotenv。已存在的环境变量不覆盖。"""
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())


def config():
    return {
        "app_id": os.environ.get("APP_ID", ""),
        "secret": os.environ.get("APP_SECRET", ""),
        "redirect_uri": os.environ.get("REDIRECT_URI", ""),
        "dingtalk_webhook": os.environ.get("DINGTALK_WEBHOOK", ""),
        "auth_user": os.environ.get("AUTH_USER", "admin"),
        "auth_password": os.environ.get("AUTH_PASSWORD", ""),
    }


class KuaishouAPI:
    def __init__(self, app_id, secret, redirect_uri):
        self.app_id = app_id
        self.secret = secret
        self.redirect_uri = redirect_uri

    def authorize_url(self, state):
        scope = quote(json.dumps(SCOPE, ensure_ascii=False, separators=(",", ":")))
        return (
            f"{AUTHORIZE_URL}?app_id={self.app_id}"
            f"&redirect_uri={quote(self.redirect_uri, safe='')}"
            f"&scope={scope}&state={state}&oauth_type=advertiser"
        )

    def _post(self, path, payload, access_token=None):
        headers = {"Content-Type": "application/json"}
        if access_token:
            headers["Access-Token"] = access_token
        r = requests.post(GATEWAY + path, json=payload, headers=headers, timeout=30)
        r.raise_for_status()
        return r.json()

    @staticmethod
    def ok(resp):
        # 快手各接口成功码不统一：token 用 0，approval 用 1，兼容 200
        return resp.get("code") in (0, 1, 200)

    def get_token(self, auth_code):
        return self._post("/oauth2/authorize/access_token", {
            "app_id": self.app_id, "secret": self.secret, "auth_code": auth_code,
        })

    def refresh(self, refresh_token):
        return self._post("/oauth2/authorize/refresh_token", {
            "app_id": self.app_id, "secret": self.secret, "refresh_token": refresh_token,
        })

    def storewide_report(self, access_token, advertiser_id, start_date, end_date, query_version=4,
                         view_type=1, select_columns=None):
        """获取全站数据。start_date/end_date 为 YYYY-MM-DD，含端点。
        query_version：4=商品全站推广 5=直播全站推广。view_type：1=账户维度 3000=计划/视频维度。
        自动翻页拿全（page_size 上限 2000）。"""
        import datetime

        def _ms(d, plus_days=0):
            dt = (datetime.datetime.strptime(d, "%Y-%m-%d") + datetime.timedelta(days=plus_days))
            dt = dt.replace(tzinfo=datetime.timezone(datetime.timedelta(hours=8)))
            return int(dt.timestamp() * 1000)

        body = {
            "advertiser_id": int(advertiser_id),
            "start_time": _ms(start_date),
            "end_time": _ms(end_date, plus_days=1),
            "select_columns": select_columns or STOREWIDE_SELECT_COLUMNS,
            "group_type": 1,
            "query_version": query_version,
            "view_type": view_type,
        }
        detail_all = []
        resp = None
        page = 1
        while True:
            body["page_info"] = {"current_page": page, "page_size": 2000}
            resp = self._post("/gw/esp/report/storewideReport", body, access_token)
            data = resp.get("data") or {}
            batch = data.get("detail") or []
            detail_all.extend(batch)
            total = data.get("total_count") or 0
            if not batch or len(detail_all) >= total:
                break
            page += 1
            if page > 50:
                break
        if resp is not None:
            resp.setdefault("data", {})["detail"] = detail_all
        return resp

    def ad_search_campaigns(self, access_token, advertiser_id, start_date, end_date):
        """拿 campaign_id -> campaign_name 映射（ad/search 组维度，含计划名）。"""
        import datetime

        def _ms(d, plus_days=0):
            dt = (datetime.datetime.strptime(d, "%Y-%m-%d") + datetime.timedelta(days=plus_days))
            dt = dt.replace(tzinfo=datetime.timezone(datetime.timedelta(hours=8)))
            return int(dt.timestamp() * 1000)

        r = self._post("/gw/esp/v1/report/ad/search", {
            "advertiser_id": int(advertiser_id),
            "page_info": {"page": 1, "page_size": 200},
            "param": {"start_time": _ms(start_date), "end_time": _ms(end_date, plus_days=1),
                      "search_level": 2},
        }, access_token)
        rows = (r.get("data") or {}).get("data") or []
        return {str(x.get("campaign_id")): x.get("campaign_name") or "" for x in rows if x.get("campaign_id")}

    def filemanage_video_list(self, access_token, advertiser_id, start_time=None, end_time=None,
                              photo_id=None):
        """视频库列表。返回 {photo_id: {name, url, cover}}。
        按创建时间筛时跨度 ≤94 天；传 photo_id 时可不传时间。"""
        result = {}
        page = 1
        while True:
            param = {}
            if photo_id:
                param["photo_id"] = str(photo_id)
            else:
                param = {"start_time": start_time, "end_time": end_time}
            r = self._post("/gw/esp/v1/filemanage/video/list", {
                "advertiser_id": int(advertiser_id),
                "page_info": {"page": page, "page_size": 100},
                "param": param,
            }, access_token)
            data = r.get("data") or {}
            details = data.get("details") or []
            for x in details:
                pid = x.get("photo_id")
                if pid:
                    result[str(pid)] = {
                        "name": x.get("photo_name") or "",
                        "url": x.get("url") or "",
                        "cover": x.get("cover_url") or "",
                    }
            total = data.get("total_count") or 0
            if not details or photo_id or len(result) >= total:
                break
            page += 1
            if page > 100:  # 防御性上限
                break
        return result

    def approval_list(self, access_token):
        """取授权覆盖的广告主 ID 列表。page_size 上限 200。"""
        return self._post("/oauth2/authorize/approval/list", {
            "app_id": self.app_id, "secret": self.secret,
            "access_token": access_token,
            "page_no": 1, "page_size": 200,
        }, access_token)


def parse_approval_advertiser_ids(resp):
    data = resp.get("data") or {}
    details = data.get("details") or []
    return [str(d) for d in details if d]


def parse_storewide(resp):
    """从 storewideReport 响应提取 {stat_date: {charge, show, gmv, roi}}。
    金额字段 ÷1000 转元；roi = t0_order_payment_amt / cost_total。"""
    out = {}
    for row in (resp.get("data") or {}).get("detail") or []:
        date = row.get("report_date_str")
        if not date:
            continue
        cost = float(row.get("cost_total") or 0)
        amt = float(row.get("t0_order_payment_amt") or 0)
        out[date] = {
            "charge": round(cost / 1000, 2),
            "show": int(row.get("show_cnt") or 0),
            "gmv": round(amt / 1000, 2),
            "roi": round(amt / cost, 4) if cost else 0.0,
        }
    return out


def parse_storewide_campaign(resp, name_map, account_host=""):
    """计划维度响应 -> {stat_date: {campaign_id: {name, host, category, charge, gmv, roi}}}。
    name_map 来自 ad_search_campaigns；主播认不出时用 account_host 兜底。"""
    out = {}
    for row in (resp.get("data") or {}).get("detail") or []:
        date = row.get("report_date_str")
        cid = str(row.get("campaign_id") or "")
        if not date or not cid:
            continue
        cost = float(row.get("cost_total") or 0)
        amt = float(row.get("t0_order_payment_amt") or 0)
        name = name_map.get(cid, "")
        category, host = parse_campaign(name)
        if not host and account_host:
            host = account_host
        out.setdefault(date, {})[cid] = {
            "name": name, "host": host, "category": category,
            "charge": round(cost / 1000, 2), "gmv": round(amt / 1000, 2),
            "roi": round(amt / cost, 4) if cost else 0.0,
        }
    return out


def parse_storewide_photo(resp, name_map):
    """视频维度响应 -> {stat_date: {photo_id: {name, is_ai, charge, gmv, roi}}}。
    name_map 来自 filemanage_video_list；名字没同步时用 photo_id 兜底。"""
    out = {}
    for row in (resp.get("data") or {}).get("detail") or []:
        date = row.get("report_date_str")
        pid = str(row.get("photo_id") or "")
        if not date or not pid or pid == "0":
            continue
        cost = float(row.get("cost_total") or 0)
        amt = float(row.get("t0_order_payment_amt") or 0)
        name = name_map.get(pid, pid)
        out.setdefault(date, {})[pid] = {
            "name": name, "is_ai": is_ai(name),
            "charge": round(cost / 1000, 2), "gmv": round(amt / 1000, 2),
            "roi": round(amt / cost, 4) if cost else 0.0,
        }
    return out


def token_expiry(resp, key):
    """token 响应的 expires_in 是秒数，转成绝对时间戳。"""
    secs = resp.get(key)
    if not secs:
        return None
    return time.time() + float(secs)
