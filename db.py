"""SQLite 存储：accounts / tokens / daily / oauth_states 四张表。

所有函数内部开连接、用完即关，单进程够用。
"""
import os
import sqlite3
from contextlib import closing
from pathlib import Path

DB_PATH = Path(os.environ.get("KUAISHOU_DB", Path(__file__).resolve().parent / "data.db"))

SCHEMA = """
CREATE TABLE IF NOT EXISTS accounts (
    advertiser_id TEXT PRIMARY KEY,
    name          TEXT,
    subject       TEXT,
    token_status  TEXT DEFAULT 'pending',
    last_sync_at  REAL
);
CREATE TABLE IF NOT EXISTS tokens (
    advertiser_id      TEXT PRIMARY KEY,
    access_token       TEXT,
    access_expires_at  REAL,
    refresh_token      TEXT,
    refresh_expires_at REAL,
    updated_at         REAL
);
CREATE TABLE IF NOT EXISTS daily (
    advertiser_id      TEXT,
    stat_date          TEXT,
    charge             REAL,
    show               INTEGER,
    photo_click        INTEGER,
    photo_click_ratio  REAL,
    photo_click_cost   REAL,
    impression_1k_cost REAL,
    conversion_num     INTEGER,
    conversion_num_cost REAL,
    gmv                REAL,
    roi                REAL,
    paid_order         INTEGER,
    updated_at         REAL,
    PRIMARY KEY (advertiser_id, stat_date)
);
CREATE TABLE IF NOT EXISTS daily_campaign (
    advertiser_id  TEXT,
    campaign_id    TEXT,
    campaign_name  TEXT,
    host           TEXT,
    category       TEXT,
    stat_date      TEXT,
    charge         REAL,
    gmv            REAL,
    roi            REAL,
    updated_at     REAL,
    PRIMARY KEY (advertiser_id, campaign_id, stat_date)
);
CREATE TABLE IF NOT EXISTS daily_photo (
    advertiser_id  TEXT,
    photo_id       TEXT,
    photo_name     TEXT,
    is_ai          INTEGER,
    stat_date      TEXT,
    charge         REAL,
    gmv            REAL,
    roi            REAL,
    updated_at     REAL,
    PRIMARY KEY (advertiser_id, photo_id, stat_date)
);
CREATE TABLE IF NOT EXISTS photo_map (
    photo_id    TEXT PRIMARY KEY,
    photo_name  TEXT,
    updated_at  REAL
);
CREATE TABLE IF NOT EXISTS photo_daily (
    advertiser_id  TEXT,
    photo_id       TEXT,
    campaign_id    TEXT,
    stat_date      TEXT,
    exposure       REAL,   -- 罗盘曝光数（完播率加权的分母）
    gmv3           REAL,   -- 3日累计GMV(元)
    r3n            REAL,   -- 3s完播率(0~100) × 曝光数
    is_ai          INTEGER,-- 按视频名判定的 AI 标记
    updated_at     REAL,
    PRIMARY KEY (advertiser_id, photo_id, campaign_id, stat_date)
);
CREATE TABLE IF NOT EXISTS photo_url (
    photo_id    TEXT PRIMARY KEY,
    url         TEXT,
    cover_url   TEXT,
    updated_at  REAL
);
CREATE TABLE IF NOT EXISTS oauth_states (
    state       TEXT PRIMARY KEY,
    created_at  REAL
);
"""


def connect():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c


def init():
    with closing(connect()) as c:
        # photo_daily 是 xlsx 派生缓存，旧结构（缺 campaign_id）直接重建
        try:
            c.execute("SELECT campaign_id FROM photo_daily LIMIT 1")
        except sqlite3.OperationalError:
            c.execute("DROP TABLE IF EXISTS photo_daily")
        c.executescript(SCHEMA)
        # 迁移：给旧库补 roi 列（旧库是磁力引擎口径，无 roi）
        try:
            c.execute("ALTER TABLE daily ADD COLUMN roi REAL")
        except sqlite3.OperationalError:
            pass
        c.commit()


def _query(sql, params=()):
    with closing(connect()) as c:
        return [dict(r) for r in c.execute(sql, params).fetchall()]


def _execute(sql, params=()):
    with closing(connect()) as c:
        c.execute(sql, params)
        c.commit()


# ---- accounts ----

def upsert_account(advertiser_id, name=None, subject=None):
    with closing(connect()) as c:
        c.execute(
            "INSERT INTO accounts (advertiser_id, name, subject) VALUES (?, ?, ?) "
            "ON CONFLICT(advertiser_id) DO UPDATE SET "
            "name=COALESCE(excluded.name, name), subject=COALESCE(excluded.subject, subject)",
            (advertiser_id, name, subject),
        )
        c.commit()


def list_accounts():
    return _query("SELECT * FROM accounts ORDER BY rowid")


def set_account(advertiser_id, name=None, subject=None, token_status=None, last_sync_at=None):
    sets, params = [], []
    if name is not None:
        sets.append("name=?"); params.append(name)
    if subject is not None:
        sets.append("subject=?"); params.append(subject)
    if token_status is not None:
        sets.append("token_status=?"); params.append(token_status)
    if last_sync_at is not None:
        sets.append("last_sync_at=?"); params.append(last_sync_at)
    if not sets:
        return
    params.append(advertiser_id)
    with closing(connect()) as c:
        c.execute(f"UPDATE accounts SET {', '.join(sets)} WHERE advertiser_id=?", params)
        c.commit()


def delete_account(advertiser_id):
    with closing(connect()) as c:
        c.execute("DELETE FROM accounts WHERE advertiser_id=?", (advertiser_id,))
        c.execute("DELETE FROM tokens WHERE advertiser_id=?", (advertiser_id,))
        c.execute("DELETE FROM daily WHERE advertiser_id=?", (advertiser_id,))
        c.execute("DELETE FROM daily_campaign WHERE advertiser_id=?", (advertiser_id,))
        c.execute("DELETE FROM daily_photo WHERE advertiser_id=?", (advertiser_id,))
        c.commit()


# ---- tokens ----

def save_tokens(advertiser_id, access_token, access_expires_at, refresh_token, refresh_expires_at):
    import time
    with closing(connect()) as c:
        c.execute(
            "INSERT INTO tokens (advertiser_id, access_token, access_expires_at, refresh_token, refresh_expires_at, updated_at) "
            "VALUES (?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(advertiser_id) DO UPDATE SET access_token=excluded.access_token, "
            "access_expires_at=excluded.access_expires_at, refresh_token=excluded.refresh_token, "
            "refresh_expires_at=excluded.refresh_expires_at, updated_at=excluded.updated_at",
            (advertiser_id, access_token, access_expires_at, refresh_token, refresh_expires_at, time.time()),
        )
        c.commit()


def get_tokens(advertiser_id):
    with closing(connect()) as c:
        row = c.execute("SELECT * FROM tokens WHERE advertiser_id=?", (advertiser_id,)).fetchone()
        return dict(row) if row else None


# ---- daily ----

def upsert_daily(advertiser_id, stat_date, d):
    import time
    with closing(connect()) as c:
        c.execute(
            "INSERT INTO daily (advertiser_id, stat_date, charge, show, photo_click, photo_click_ratio, "
            "photo_click_cost, impression_1k_cost, conversion_num, conversion_num_cost, gmv, roi, paid_order, updated_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(advertiser_id, stat_date) DO UPDATE SET charge=excluded.charge, show=excluded.show, "
            "photo_click=excluded.photo_click, photo_click_ratio=excluded.photo_click_ratio, "
            "photo_click_cost=excluded.photo_click_cost, impression_1k_cost=excluded.impression_1k_cost, "
            "conversion_num=excluded.conversion_num, conversion_num_cost=excluded.conversion_num_cost, "
            "gmv=excluded.gmv, roi=excluded.roi, paid_order=excluded.paid_order, updated_at=excluded.updated_at",
            (advertiser_id, stat_date, d.get("charge"), d.get("show"), d.get("photo_click"),
             d.get("photo_click_ratio"), d.get("photo_click_cost"), d.get("impression_1k_cost"),
             d.get("conversion_num"), d.get("conversion_num_cost"), d.get("gmv"), d.get("roi"),
             d.get("paid_order"), time.time()),
        )
        c.commit()


def has_daily(advertiser_id):
    with closing(connect()) as c:
        row = c.execute("SELECT 1 FROM daily WHERE advertiser_id=? LIMIT 1", (advertiser_id,)).fetchone()
        return row is not None


def query_daily(start_date, end_date, advertiser_ids=None):
    sql = "SELECT * FROM daily WHERE stat_date BETWEEN ? AND ?"
    params = [start_date, end_date]
    if advertiser_ids:
        sql += f" AND advertiser_id IN ({','.join('?' * len(advertiser_ids))})"
        params += list(advertiser_ids)
    sql += " ORDER BY stat_date, advertiser_id"
    return _query(sql, params)


# ---- daily_campaign ----

def upsert_campaign(advertiser_id, campaign_id, campaign_name, host, category,
                    stat_date, charge, gmv, roi):
    import time
    with closing(connect()) as c:
        c.execute(
            "INSERT INTO daily_campaign (advertiser_id, campaign_id, campaign_name, host, category, "
            "stat_date, charge, gmv, roi, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(advertiser_id, campaign_id, stat_date) DO UPDATE SET "
            "campaign_name=excluded.campaign_name, host=excluded.host, category=excluded.category, "
            "charge=excluded.charge, gmv=excluded.gmv, roi=excluded.roi, updated_at=excluded.updated_at",
            (advertiser_id, campaign_id, campaign_name, host, category, stat_date,
             charge, gmv, roi, time.time()),
        )
        c.commit()


def query_campaign(start_date, end_date, advertiser_ids=None):
    sql = "SELECT * FROM daily_campaign WHERE stat_date BETWEEN ? AND ?"
    params = [start_date, end_date]
    if advertiser_ids:
        sql += f" AND advertiser_id IN ({','.join('?' * len(advertiser_ids))})"
        params += list(advertiser_ids)
    sql += " ORDER BY stat_date, gmv DESC"
    rows = _query(sql, params)
    # 补计划级 AI视频3日累计GMV（来自罗盘导入的 photo_daily，MAPI 不提供）
    ai_map = query_campaign_ai_gmv3()
    for r in rows:
        r["ai_gmv3"] = ai_map.get((r["advertiser_id"], r["campaign_id"], r["stat_date"]), None)
    return rows


# ---- daily_photo ----

def upsert_photo(advertiser_id, photo_id, photo_name, is_ai, stat_date, charge, gmv, roi):
    import time
    with closing(connect()) as c:
        c.execute(
            "INSERT INTO daily_photo (advertiser_id, photo_id, photo_name, is_ai, stat_date, "
            "charge, gmv, roi, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(advertiser_id, photo_id, stat_date) DO UPDATE SET "
            "photo_name=excluded.photo_name, is_ai=excluded.is_ai, charge=excluded.charge, "
            "gmv=excluded.gmv, roi=excluded.roi, updated_at=excluded.updated_at",
            (advertiser_id, photo_id, photo_name, 1 if is_ai else 0, stat_date,
             charge, gmv, roi, time.time()),
        )
        c.commit()


def query_photo(start_date, end_date, advertiser_ids=None):
    """读视频维度。名字以 photo_map 为准（覆盖历史行的数字ID），AI 标记按最终名字重算。"""
    sql = "SELECT * FROM daily_photo WHERE stat_date BETWEEN ? AND ?"
    params = [start_date, end_date]
    if advertiser_ids:
        sql += f" AND advertiser_id IN ({','.join('?' * len(advertiser_ids))})"
        params += list(advertiser_ids)
    sql += " ORDER BY stat_date, gmv DESC"
    rows = _query(sql, params)

    import kuaishou  # 延迟导入，避免与 kuaishou 的潜在循环
    name_map = query_photo_map()
    daily = query_photo_daily_map()
    urls = query_photo_url_map()
    for r in rows:
        nm = name_map.get(r["photo_id"])
        if nm and not str(nm).isdigit():
            r["photo_name"] = nm
        r["is_ai"] = 1 if kuaishou.is_ai(r["photo_name"]) else 0
        # 罗盘补充指标：3日累计GMV、3s完播率分子/分母（MAPI 不提供，缺则 None）
        d = daily.get((r["advertiser_id"], r["photo_id"], r["stat_date"]))
        r["gmv3"] = d["gmv3"] if d else None
        r["exposure"] = d["exposure"] if d else None
        r["r3n"] = d["r3n"] if d else None
        # 播放地址（可能过期，前端可点「刷新链接」重新取）
        u = urls.get(r["photo_id"]) or {}
        r["video_url"] = u.get("url")
        r["cover_url"] = u.get("cover_url")
    return rows


# ---- photo_map（罗盘导出的视频ID->名称映射，覆盖 MAPI 视频库缺失的名字）----

def upsert_photo_map(photo_id, photo_name):
    import time
    with closing(connect()) as c:
        c.execute(
            "INSERT INTO photo_map (photo_id, photo_name, updated_at) VALUES (?, ?, ?) "
            "ON CONFLICT(photo_id) DO UPDATE SET photo_name=excluded.photo_name, updated_at=excluded.updated_at",
            (photo_id, photo_name, time.time()),
        )
        c.commit()


def query_photo_map():
    return {r["photo_id"]: r["photo_name"] for r in _query("SELECT photo_id, photo_name FROM photo_map")}


def upsert_photo_daily(advertiser_id, photo_id, campaign_id, stat_date, exposure, gmv3, r3n, is_ai):
    import time
    with closing(connect()) as c:
        c.execute(
            "INSERT INTO photo_daily (advertiser_id, photo_id, campaign_id, stat_date, exposure, gmv3, r3n, is_ai, updated_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(advertiser_id, photo_id, campaign_id, stat_date) DO UPDATE SET "
            "exposure=excluded.exposure, gmv3=excluded.gmv3, r3n=excluded.r3n, "
            "is_ai=excluded.is_ai, updated_at=excluded.updated_at",
            (advertiser_id, photo_id, campaign_id, stat_date, exposure, gmv3, r3n,
             1 if is_ai else 0, time.time()),
        )
        c.commit()


def query_photo_daily_map():
    """{(advertiser_id, photo_id, stat_date): {exposure, gmv3, r3n}}，多计划同视频按天求和。"""
    out = {}
    sql = ("SELECT advertiser_id, photo_id, stat_date, SUM(exposure) e, SUM(gmv3) g, SUM(r3n) r "
           "FROM photo_daily GROUP BY advertiser_id, photo_id, stat_date")
    for r in _query(sql):
        out[(r["advertiser_id"], r["photo_id"], r["stat_date"])] = {
            "exposure": r["e"], "gmv3": r["g"], "r3n": r["r"],
        }
    return out


def query_campaign_ai_gmv3():
    """{(advertiser_id, campaign_id, stat_date): AI视频3日累计GMV}"""
    sql = ("SELECT advertiser_id, campaign_id, stat_date, SUM(gmv3) g FROM photo_daily "
           "WHERE is_ai=1 GROUP BY advertiser_id, campaign_id, stat_date")
    return {(r["advertiser_id"], r["campaign_id"], r["stat_date"]): r["g"] for r in _query(sql)}


# ---- photo_url（视频播放地址，同步时从 MAPI 视频库取）----

def upsert_photo_url(photo_id, url, cover_url):
    import time
    if not url:
        return
    with closing(connect()) as c:
        c.execute(
            "INSERT INTO photo_url (photo_id, url, cover_url, updated_at) VALUES (?, ?, ?, ?) "
            "ON CONFLICT(photo_id) DO UPDATE SET url=excluded.url, cover_url=excluded.cover_url, "
            "updated_at=excluded.updated_at",
            (photo_id, url, cover_url, time.time()),
        )
        c.commit()


def query_photo_url(photo_id):
    rows = _query("SELECT url, cover_url FROM photo_url WHERE photo_id=?", (photo_id,))
    return rows[0] if rows else None


def query_photo_url_map():
    return {r["photo_id"]: r for r in _query("SELECT photo_id, url, cover_url FROM photo_url")}


# ---- oauth_states ----

def save_state(state):
    import time
    _execute("INSERT OR REPLACE INTO oauth_states (state, created_at) VALUES (?, ?)", (state, time.time()))


def pop_state(state):
    with closing(connect()) as c:
        row = c.execute("SELECT 1 FROM oauth_states WHERE state=?", (state,)).fetchone()
        if row:
            c.execute("DELETE FROM oauth_states WHERE state=?", (state,))
            c.commit()
            return True
        return False
