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


if __name__ == "__main__":
    test_recent_dates()
    test_parse_storewide()
    test_parse_campaign()
    print("test_sync.py 全部通过")
