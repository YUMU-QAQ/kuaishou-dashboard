"""从罗盘导出的「自定义报表明细」xlsx 导入两类数据：

1. 视频ID -> 视频名称 映射（photo_map）——MAPI 视频库对小店上传视频不存名字，用罗盘真名覆盖。
2. 视频 × 日期 的补充指标（photo_daily）：3日累计GMV、曝光数、3s完播率分子
   ——MAPI 的 storewideReport 不返回这些字段。

用法：python import_photo_map.py <明细.xlsx>   （可重复导入，按主键覆盖）
"""
import sys
from collections import defaultdict

import openpyxl

import db
import kuaishou


def fnum(v):
    s = str(v or "").replace("%", "").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def main():
    src = sys.argv[1]
    wb = openpyxl.load_workbook(src, read_only=False, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[0]
    idx = {c: i for i, c in enumerate(hdr)}

    named = skip = 0
    # (账户ID, 视频ID, 计划ID, 时间) -> 聚合
    agg = defaultdict(lambda: {"exposure": 0.0, "gmv3": 0.0, "r3n": 0.0, "rend": 0.0, "ai": False})
    i_rend = idx.get("作品完播率")  # 老导出可能没有这列
    for r in rows[1:]:
        vid = str(r[idx["视频ID"]])
        name = str(r[idx["视频名称"]])
        if vid == "0":
            continue
        ai = False
        # 名字映射
        if name and name != vid and not name.isdigit():
            db.upsert_photo_map(vid, name)
            named += 1
            ai = kuaishou.is_ai(name)
        else:
            skip += 1
        # 补充指标
        aid = str(r[idx["账户ID"]])
        cid = str(r[idx["广告计划ID"]])
        date = str(r[idx["时间"]])
        exp = fnum(r[idx["曝光数"]])
        g = agg[(aid, vid, cid, date)]
        g["exposure"] += exp
        g["gmv3"] += fnum(r[idx["3日累计GMV"]])
        g["r3n"] += fnum(r[idx["作品3s完播率"]]) * exp   # 完播率(0~100)×曝光，加权分子
        if i_rend is not None:
            g["rend"] += fnum(r[i_rend]) * exp
        g["ai"] = g["ai"] or ai

    for (aid, vid, cid, date), g in agg.items():
        db.upsert_photo_daily(aid, vid, cid, date, round(g["exposure"], 2),
                              round(g["gmv3"], 2), round(g["r3n"], 4), g["ai"],
                              round(g["rend"], 4) if i_rend is not None else None)

    print(f"名字映射 {named} 条（跳过 {skip} 条数字ID/空名）；补充指标 {len(agg)} 条 "
          f"（{len({k[3] for k in agg})} 个日期）")


if __name__ == "__main__":
    db.init()
    main()
