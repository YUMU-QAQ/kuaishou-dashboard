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
    app_id        TEXT,
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
-- 磁力引擎应用的授权（与金牛授权分开存）：同一店铺会同时挂两个应用，
-- 引擎那份只用来调 v1/report/creative_report 拿 3秒播放率/完播率，不能覆盖 accounts.app_id。
CREATE TABLE IF NOT EXISTS engine_tokens (
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
    -- 播放率三个分子分母，来自 v1/report/account_report（磁力引擎口径，MAPI 全站接口不提供）。
    -- 跨天/跨账号要加总分子分母再除，不能平均比率。
    played_num         REAL,   -- 播放基数（行为数）
    played_3s          REAL,   -- 3秒播放数
    played_end         REAL,   -- 完播数
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
    campaign_id    TEXT,
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
    rend           REAL,   -- 作品完播率(0~100) × 曝光数
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
    app_id      TEXT,
    created_at  REAL
);
CREATE TABLE IF NOT EXISTS acct_groups (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT UNIQUE,
    created_at  REAL
);
CREATE TABLE IF NOT EXISTS group_accounts (
    group_id       INTEGER,
    advertiser_id  TEXT,
    PRIMARY KEY (group_id, advertiser_id)
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
        # 迁移：多应用支持，给旧库补 app_id 列
        for sql in ("ALTER TABLE accounts ADD COLUMN app_id TEXT",
                    "ALTER TABLE oauth_states ADD COLUMN app_id TEXT",
                    "ALTER TABLE daily_photo ADD COLUMN campaign_id TEXT",
                    "ALTER TABLE daily ADD COLUMN played_num REAL",
                    "ALTER TABLE daily ADD COLUMN played_3s REAL",
                    "ALTER TABLE daily ADD COLUMN played_end REAL",
                    "ALTER TABLE photo_daily ADD COLUMN rend REAL"):
            try:
                c.execute(sql)
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

def upsert_account(advertiser_id, name=None, subject=None, app_id=None):
    with closing(connect()) as c:
        c.execute(
            "INSERT INTO accounts (advertiser_id, name, subject, app_id) VALUES (?, ?, ?, ?) "
            "ON CONFLICT(advertiser_id) DO UPDATE SET "
            "name=COALESCE(excluded.name, name), subject=COALESCE(excluded.subject, subject), "
            "app_id=COALESCE(excluded.app_id, app_id)",
            (advertiser_id, name, subject, app_id),
        )
        c.commit()


def list_accounts():
    return _query("SELECT * FROM accounts ORDER BY rowid")


def get_account(advertiser_id):
    rows = _query("SELECT * FROM accounts WHERE advertiser_id=?", (advertiser_id,))
    return rows[0] if rows else None


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


def save_engine_tokens(advertiser_id, access_token, access_expires_at, refresh_token, refresh_expires_at):
    """磁力引擎应用的授权，独立存，不碰 accounts.app_id / tokens。"""
    import time
    with closing(connect()) as c:
        c.execute(
            "INSERT INTO engine_tokens (advertiser_id, access_token, access_expires_at, refresh_token, refresh_expires_at, updated_at) "
            "VALUES (?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(advertiser_id) DO UPDATE SET access_token=excluded.access_token, "
            "access_expires_at=excluded.access_expires_at, refresh_token=excluded.refresh_token, "
            "refresh_expires_at=excluded.refresh_expires_at, updated_at=excluded.updated_at",
            (advertiser_id, access_token, access_expires_at, refresh_token, refresh_expires_at, time.time()),
        )
        c.commit()


def get_engine_tokens(advertiser_id):
    with closing(connect()) as c:
        row = c.execute("SELECT * FROM engine_tokens WHERE advertiser_id=?", (advertiser_id,)).fetchone()
        return dict(row) if row else None


# ---- daily ----

def upsert_daily(advertiser_id, stat_date, d):
    import time
    with closing(connect()) as c:
        c.execute(
            "INSERT INTO daily (advertiser_id, stat_date, charge, show, photo_click, photo_click_ratio, "
            "photo_click_cost, impression_1k_cost, conversion_num, conversion_num_cost, gmv, roi, paid_order, "
            "played_num, played_3s, played_end, updated_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(advertiser_id, stat_date) DO UPDATE SET charge=excluded.charge, show=excluded.show, "
            "photo_click=excluded.photo_click, photo_click_ratio=excluded.photo_click_ratio, "
            "photo_click_cost=excluded.photo_click_cost, impression_1k_cost=excluded.impression_1k_cost, "
            "conversion_num=excluded.conversion_num, conversion_num_cost=excluded.conversion_num_cost, "
            "gmv=excluded.gmv, roi=excluded.roi, paid_order=excluded.paid_order, "
            "played_num=COALESCE(excluded.played_num, played_num), "
            "played_3s=COALESCE(excluded.played_3s, played_3s), "
            "played_end=COALESCE(excluded.played_end, played_end), updated_at=excluded.updated_at",
            (advertiser_id, stat_date, d.get("charge"), d.get("show"), d.get("photo_click"),
             d.get("photo_click_ratio"), d.get("photo_click_cost"), d.get("impression_1k_cost"),
             d.get("conversion_num"), d.get("conversion_num_cost"), d.get("gmv"), d.get("roi"),
             d.get("paid_order"),
             d.get("played_num"), d.get("played_3s"), d.get("played_end"), time.time()),
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
    return _query(sql, params)


# ---- daily_photo ----

def upsert_photo(advertiser_id, photo_id, photo_name, is_ai, stat_date, charge, gmv, roi, campaign_id=None):
    import time
    with closing(connect()) as c:
        c.execute(
            "INSERT INTO daily_photo (advertiser_id, photo_id, photo_name, campaign_id, is_ai, stat_date, "
            "charge, gmv, roi, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(advertiser_id, photo_id, stat_date) DO UPDATE SET "
            "photo_name=excluded.photo_name, campaign_id=COALESCE(excluded.campaign_id, campaign_id), "
            "is_ai=excluded.is_ai, charge=excluded.charge, "
            "gmv=excluded.gmv, roi=excluded.roi, updated_at=excluded.updated_at",
            (advertiser_id, photo_id, photo_name, campaign_id, 1 if is_ai else 0, stat_date,
             charge, gmv, roi, time.time()),
        )
        c.commit()


def query_photo(start_date, end_date, advertiser_ids=None):
    """读视频维度。名字以 photo_map 为准（覆盖历史行的数字ID），AI 标记按最终名字重算。
    播放地址不在这里返回——太长且同视频多天重复，改由 /api/photos 单独给一份去重的。"""
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
        r.pop("updated_at", None)  # 前端不用，去掉省体积
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


def upsert_photo_daily(advertiser_id, photo_id, campaign_id, stat_date, exposure, gmv3, r3n, is_ai, rend=None):
    import time
    with closing(connect()) as c:
        c.execute(
            "INSERT INTO photo_daily (advertiser_id, photo_id, campaign_id, stat_date, exposure, gmv3, r3n, rend, is_ai, updated_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
            "ON CONFLICT(advertiser_id, photo_id, campaign_id, stat_date) DO UPDATE SET "
            "exposure=excluded.exposure, gmv3=excluded.gmv3, r3n=excluded.r3n, "
            "rend=COALESCE(excluded.rend, rend), "
            "is_ai=excluded.is_ai, updated_at=excluded.updated_at",
            (advertiser_id, photo_id, campaign_id, stat_date, exposure, gmv3, r3n, rend,
             1 if is_ai else 0, time.time()),
        )
        c.commit()


def query_photo_daily_map():
    """{(advertiser_id, photo_id, stat_date): {exposure, gmv3, r3n, rend}}，多计划同视频按天求和。"""
    out = {}
    sql = ("SELECT advertiser_id, photo_id, stat_date, SUM(exposure) e, SUM(gmv3) g, "
           "SUM(r3n) r, SUM(rend) rend FROM photo_daily GROUP BY advertiser_id, photo_id, stat_date")
    for r in _query(sql):
        out[(r["advertiser_id"], r["photo_id"], r["stat_date"])] = {
            "exposure": r["e"], "gmv3": r["g"], "r3n": r["r"], "rend": r["rend"],
        }
    return out


def query_photo_daily_totals(start_date, end_date, advertiser_ids=None):
    """{advertiser_id: {exposure, r3n, rend}}——3秒播放率/完播率的分子分母，区间内按账号汇总。
    有播放率数据的账号才出现在结果里（没导罗盘明细的账号没有）。"""
    sql = ("SELECT advertiser_id, SUM(exposure) exposure, SUM(r3n) r3n, SUM(rend) rend "
           "FROM photo_daily WHERE stat_date BETWEEN ? AND ?")
    params = [start_date, end_date]
    if advertiser_ids:
        sql += f" AND advertiser_id IN ({','.join('?' * len(advertiser_ids))})"
        params += list(advertiser_ids)
    sql += " GROUP BY advertiser_id"
    return {r["advertiser_id"]: {"exposure": r["exposure"], "r3n": r["r3n"], "rend": r["rend"]}
            for r in _query(sql, params)}


def query_video_counts(start_date, end_date, advertiser_ids=None):
    """{advertiser_id: 视频数}——区间内有投放记录的去重视频数。"""
    sql = ("SELECT advertiser_id, COUNT(DISTINCT photo_id) n FROM daily_photo "
           "WHERE stat_date BETWEEN ? AND ?")
    params = [start_date, end_date]
    if advertiser_ids:
        sql += f" AND advertiser_id IN ({','.join('?' * len(advertiser_ids))})"
        params += list(advertiser_ids)
    sql += " GROUP BY advertiser_id"
    return {r["advertiser_id"]: r["n"] for r in _query(sql, params)}


# ---- 自定义分组（把若干账号归到一个组里看汇总） ----

def list_groups():
    """[{id, name, accounts: [advertiser_id, ...]}]"""
    out = {r["id"]: {"id": r["id"], "name": r["name"], "accounts": []} for r in _query("SELECT * FROM acct_groups ORDER BY id")}
    for r in _query("SELECT group_id, advertiser_id FROM group_accounts"):
        if r["group_id"] in out:
            out[r["group_id"]]["accounts"].append(r["advertiser_id"])
    return list(out.values())


def create_group(name):
    import time
    with closing(connect()) as c:
        try:
            cur = c.execute("INSERT INTO acct_groups (name, created_at) VALUES (?, ?)", (name, time.time()))
            c.commit()
            return cur.lastrowid
        except sqlite3.IntegrityError:
            return None  # 同名分组已存在


def rename_group(group_id, name):
    """改分组名。返回 False 表示与已有分组重名。"""
    with closing(connect()) as c:
        try:
            c.execute("UPDATE acct_groups SET name=? WHERE id=?", (name, group_id))
            c.commit()
            return True
        except sqlite3.IntegrityError:
            return False


def delete_group(group_id):
    with closing(connect()) as c:
        c.execute("DELETE FROM acct_groups WHERE id=?", (group_id,))
        c.execute("DELETE FROM group_accounts WHERE group_id=?", (group_id,))
        c.commit()


def set_group_accounts(group_id, advertiser_ids):
    with closing(connect()) as c:
        c.execute("DELETE FROM group_accounts WHERE group_id=?", (group_id,))
        c.executemany("INSERT OR IGNORE INTO group_accounts (group_id, advertiser_id) VALUES (?, ?)",
                      [(group_id, a) for a in advertiser_ids])
        c.commit()


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

def save_state(state, app_id=None):
    import time
    _execute("INSERT OR REPLACE INTO oauth_states (state, app_id, created_at) VALUES (?, ?, ?)",
             (state, app_id, time.time()))


def pop_state(state):
    """取出并删除 state。返回该次授权应使用的 app_id（无则空串）；state 无效返回 None。"""
    with closing(connect()) as c:
        row = c.execute("SELECT app_id FROM oauth_states WHERE state=?", (state,)).fetchone()
        if not row:
            return None
        c.execute("DELETE FROM oauth_states WHERE state=?", (state,))
        c.commit()
        return row["app_id"] or ""
