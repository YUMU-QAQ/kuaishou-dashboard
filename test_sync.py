"""两个最容易写错的逻辑的自我检查：日期窗口计算 + 报表字段映射解析。

不连真实接口，纯函数断言。跑：python test_sync.py
"""
from datetime import date

import kuaishou
import sync


def test_recent_dates():
    from unittest import mock
    with mock.patch.object(sync, "date") as mock_date:
        mock_date.today.return_value = date(2026, 1, 1)  # 跨年边界
        mock_date.side_effect = date

        d7 = sync.recent_dates(7)
        assert len(d7) == 7, "回补应为 7 天"
        assert d7[0] == "2025-12-31", f"最近一天应为 t-1，得 {d7[0]}"
        assert d7[-1] == "2025-12-25", f"最远一天应为 t-7，得 {d7[-1]}"
        assert "2026-01-01" not in d7, "窗口不应包含今天"


def test_parse_storewide():
    # storewideReport 响应，金额 ÷1000 转元，roi = amt/cost
    resp = {
        "code": 0,
        "data": {"detail": [
            {"report_date_str": "2026-10-07", "cost_total": 938401.0,
             "show_cnt": 47422, "t0_order_payment_amt": 3608600},
        ]},
    }
    days = kuaishou.parse_storewide(resp)
    d = days["2026-10-07"]
    assert d["charge"] == 938.40, f"消耗应 ÷1000，得 {d['charge']}"
    assert d["gmv"] == 3608.60, f"GMV 应 ÷1000，得 {d['gmv']}"
    assert d["roi"] == 3.8455, f"ROI = 3608600/938401，得 {d['roi']}"
    assert d["show"] == 47422

    # detail 缺失（空报表）不抛异常
    assert kuaishou.parse_storewide({"code": 0, "data": {}}) == {}


def test_parse_campaign():
    c, h = kuaishou.parse_campaign("9.7-树德堂阿胶浆-泽鹏【目标投产4.8】")
    assert c == "树德堂阿胶浆", f"品类应为 树德堂阿胶浆，得 {c}"
    assert h == "泽鹏", f"主播应为 泽鹏，得 {h}"
    c, h = kuaishou.parse_campaign("7.20-泽鹏-洗发水")  # 主播品类颠倒写法
    assert h == "泽鹏", f"主播应为 泽鹏，得 {h}"
    assert c == "洗发水", f"品类应为 洗发水，得 {c}"


def test_apps_and_api_for():
    import os
    from unittest import mock
    env = {"APPS": "appA:secA,appB:secB", "REDIRECT_URI": "http://x/cb"}
    env["APP_ID"] = env["APP_SECRET"] = ""
    with mock.patch.dict(os.environ, env):
        assert kuaishou.apps() == {"appA": "secA", "appB": "secB"}
        assert kuaishou.api_for("appB").secret == "secB"
        assert kuaishou.api_for("appB").app_id == "appB"
        # 未指定 / 未知 app 都回退到配置里的第一个，不能报错
        assert kuaishou.api_for(None).app_id == "appA"
        assert kuaishou.api_for("不存在").app_id == "appA"
    # 兼容单应用老配置
    with mock.patch.dict(os.environ, {"APPS": "", "APP_ID": "only", "APP_SECRET": "s"}):
        assert kuaishou.apps() == {"only": "s"}


def test_oauth_state_carries_app():
    """一个 state 对应一个应用：回调要按发起时的应用去换 token。"""
    import os, tempfile
    from unittest import mock
    import db
    with tempfile.TemporaryDirectory() as d:
        with mock.patch.object(db, "DB_PATH", os.path.join(d, "t.db")):
            db.init()
            db.save_state("s1", "appB")
            db.save_state("s2")  # 未指定应用
            assert db.pop_state("s1") == "appB"
            assert db.pop_state("s2") == ""
            assert db.pop_state("s1") is None, "state 应一次性，用过即失效"
            assert db.pop_state("没这个") is None
            # 账号记应用；后续不带 app_id 的更新不能把已记的应用清掉
            db.upsert_account("1", app_id="appB")
            assert db.get_account("1")["app_id"] == "appB"
            db.upsert_account("1")
            assert db.get_account("1")["app_id"] == "appB"


def test_groups_and_totals():
    """分组表、完播率字段、区间汇总（比率要分子分母分别加总）。"""
    import os, tempfile
    from unittest import mock
    import db
    with tempfile.TemporaryDirectory() as d:
        with mock.patch.object(db, "DB_PATH", os.path.join(d, "t.db")):
            db.init()
            gid = db.create_group("测试组")
            assert gid, "新建分组应返回 id"
            assert db.create_group("测试组") is None, "同名分组应被拒"
            db.set_group_accounts(gid, ["a1", "a2"])
            assert db.list_groups()[0]["accounts"] == ["a1", "a2"]
            assert db.rename_group(gid, "改名后"), "改名应成功"
            assert db.list_groups()[0]["name"] == "改名后"
            gid2 = db.create_group("另一个组")
            assert db.rename_group(gid2, "改名后") is False, "改成已有名字应被拒"
            db.delete_group(gid2)

            db.upsert_daily("a1", "2026-10-01", {"gmv": 100.0, "show": 10, "charge": 5.0})
            db.upsert_photo("a1", "p1", "某视频 A", True, "2026-10-01", 5.0, 100.0, 20.0, campaign_id="c1")
            db.upsert_photo_daily("a1", "p1", "c1", "2026-10-01", 1000.0, 50.0, 33.5, True, 25.0)

            tot = db.query_photo_daily_totals("2026-10-01", "2026-10-01")
            assert tot["a1"]["rend"] == 25.0, "完播率分子应落库"
            assert tot["a1"]["r3n"] == 33.5
            assert db.query_video_counts("2026-10-01", "2026-10-01")["a1"] == 1
            row = db.query_photo("2026-10-01", "2026-10-01")[0]
            assert row["campaign_id"] == "c1", "计划维度 AI GMV 需要 campaign_id"

            db.delete_group(gid)
            assert db.list_groups() == [], "删除分组后不应残留"


def test_account_report_parse_and_store():
    """v1/report/account_report 的播放率分子分母：解析 + 落库 + 读回。
    口径按行为数(played_num)算，88759186/10-08 实测 41.43% / 5.54%。"""
    resp = {"code": 0, "data": {"details": [
        {"stat_date": "2026-10-08", "played_num": 21830, "played_three_seconds": 9045,
         "played_end": 1210},
        {"stat_date": "", "played_num": 1, "played_three_seconds": 1, "played_end": 1},  # 空日期要跳过
    ]}}
    parsed = kuaishou.parse_account_report(resp)
    assert list(parsed) == ["2026-10-08"], f"空日期应跳过，得 {list(parsed)}"
    assert parsed["2026-10-08"]["played_3s"] == 9045
    # 缺失字段补 0，别抛 KeyError
    assert kuaishou.parse_account_report(
        {"data": {"details": [{"stat_date": "2026-10-08"}]}})["2026-10-08"]["played_3s"] == 0

    import db
    import os
    import tempfile
    from unittest import mock
    with tempfile.TemporaryDirectory() as d:
        with mock.patch.object(db, "DB_PATH", os.path.join(d, "t.db")):
            db.init()
            db.upsert_daily("a1", "2026-10-08", {"gmv": 1.0, "show": 1, "charge": 1.0,
                                                 **parsed["2026-10-08"]})
            row = db.query_daily("2026-10-08", "2026-10-08")[0]
            assert round(row["played_3s"] / row["played_num"] * 100, 2) == 41.43
            assert round(row["played_end"] / row["played_num"] * 100, 2) == 5.54
            # 二次写入不带播放率时不能把已有值抹成 NULL（COALESCE）
            db.upsert_daily("a1", "2026-10-08", {"gmv": 2.0, "show": 2, "charge": 2.0})
            row = db.query_daily("2026-10-08", "2026-10-08")[0]
            assert row["played_3s"] == 9045 and row["gmv"] == 2.0, "重写不应清空播放率"


if __name__ == "__main__":
    test_recent_dates()
    test_parse_storewide()
    test_parse_campaign()
    test_apps_and_api_for()
    test_oauth_state_carries_app()
    test_groups_and_totals()
    test_account_report_parse_and_store()
    print("test_sync.py 全部通过")
