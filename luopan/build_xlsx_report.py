#!/usr/bin/env python3
"""把罗盘导出的明细 xlsx 汇成 Excel 版日报（账号维度 / 每日趋势 / 视频明细）。

口径与 HTML 版一致：账号总GMV 用全站GMV；完播率按曝光加权；比率不做行求和。
金额、比率一律写成公式，交 recalc.py 计算，不落死值。
"""
import os
import sys

import luopan_from_excel as L
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

MONEY = '#,##0.00'
COUNT = '#,##0'
RATE = '0.00%'
RATIO = '0.00'

HEAD_FILL = PatternFill("solid", fgColor="1E40AF")
HEAD_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BODY_FONT = Font(name="Arial", size=10)
NOTE_FONT = Font(name="Arial", size=9, color="475569")


def style_header(ws, ncol, row=1, freeze=True):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    if freeze:
        ws.freeze_panes = ws.cell(row=row + 1, column=1)


def put(ws, r, c, value, fmt=None):
    cell = ws.cell(row=r, column=c, value=value)
    cell.font = BODY_FONT
    if fmt:
        cell.number_format = fmt
    return cell


def build_account_sheet(ws, accounts, overall, day=""):
    heads = ["账号", "总GMV(全站)(元)", "AI视频3日累计GMV(元)", "AI视频数", "发布视频数",
             "单条平均GMV(元)", "总曝光量", "总投放消耗(元)", "投产比ROI",
             "3s完播率", "5s完播率", "完播率"]
    for i, h in enumerate(heads, 1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, len(heads))

    # 隐藏辅助列：比率按曝光加权的分子，供「合计」行正确加权
    for i, h in enumerate(["辅助_3s*曝光", "辅助_5s*曝光", "辅助_完播*曝光"], 13):
        ws.cell(row=1, column=i, value=h)
    ws.column_dimensions["M"].hidden = True
    ws.column_dimensions["N"].hidden = True
    ws.column_dimensions["O"].hidden = True

    for idx, a in enumerate(accounts):
        r = 2 + idx
        put(ws, r, 1, a["name"])
        put(ws, r, 2, a["gmv"], MONEY)
        put(ws, r, 3, a["ai_gmv3"], MONEY)
        put(ws, r, 4, a["ai_videos"], COUNT)
        put(ws, r, 5, a["videos"], COUNT)
        put(ws, r, 6, f"=IFERROR(B{r}/E{r},0)", MONEY)     # 单条平均GMV = 总GMV / 发布视频数
        put(ws, r, 7, a["exp"], COUNT)
        put(ws, r, 8, a["cost"], MONEY)
        put(ws, r, 9, f"=IFERROR(B{r}/H{r},0)", RATIO)      # ROI = 总GMV / 总消耗
        put(ws, r, 10, a["r3"] / 100, RATE)
        put(ws, r, 11, a["r5"] / 100, RATE)
        put(ws, r, 12, a["rend"] / 100, RATE)
        put(ws, r, 13, a["r3"] / 100 * a["exp"])
        put(ws, r, 14, a["r5"] / 100 * a["exp"])
        put(ws, r, 15, a["rend"] / 100 * a["exp"])

    last = 1 + len(accounts)
    tr = last + 1
    put(ws, tr, 1, "合计").font = Font(name="Arial", size=10, bold=True)
    for col in (2, 3, 7, 8):
        put(ws, tr, col, f"=SUM({get_column_letter(col)}2:{get_column_letter(col)}{last})",
            MONEY if col in (2, 3, 8) else COUNT)
    # 视频数不能 SUM：同一视频可能同时投给两个账号，相加会重复计（ID=0 两边都有）。
    # 用全区间去重值（硬编码，已注释说明），否则「单条平均GMV」会偏。
    for col, key, label in ((4, "ai_videos", "AI 视频"), (5, "videos", "视频")):
        cell = put(ws, tr, col, overall[key], COUNT)
        cell.comment = Comment(
            f"全区间按视频ID去重的{label}数（{label}不含「无视频」哨兵值 ID=0）。\n"
            f"不能写成 =SUM({get_column_letter(col)}2:{get_column_letter(col)}{last})："
            f"各账号去重数之和会重复计跨账号出现的同一 ID。", "报表")
    put(ws, tr, 6, f"=IFERROR(B{tr}/E{tr},0)", MONEY)
    put(ws, tr, 9, f"=IFERROR(B{tr}/H{tr},0)", RATIO)
    for col, hcol in ((10, "M"), (11, "N"), (12, "O")):
        put(ws, tr, col, f"=IFERROR(SUM({hcol}2:{hcol}{last})/G{tr},0)", RATE)
    for col in range(1, 13):
        ws.cell(row=tr, column=col).font = Font(name="Arial", size=10, bold=True)

    ws.cell(row=tr + 2, column=1,
            value=(f"报告日 {day}。" if day else "") +
                  "完播率为曝光数加权（明细只有比率、无分子分母，无法精确聚合）；"
                  "「合计」行的比率按各账号曝光加权，不是简单平均。"
                  "AI视频口径：标题末尾 A，或含 j/J（紧贴话题块之前）。").font = NOTE_FONT
    for col, w in zip("ABCDEFGHIJKL", (16, 16, 20, 10, 11, 15, 12, 15, 10, 10, 10, 10)):
        ws.column_dimensions[col].width = w
    return tr


def build_daily_sheet(ws, daily, acc_names):
    heads = ["日期", "账号", "全站GMV(元)", "3日累计GMV(元)", "消耗(元)",
             "视频数", "曝光", "ROI"]
    for i, h in enumerate(heads, 1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, len(heads))
    r = 2
    for d in daily:
        for n in acc_names:
            m = d["per"].get(n)
            if not m:
                continue
            put(ws, r, 1, d["date"])
            put(ws, r, 2, n)
            put(ws, r, 3, m["gmv"], MONEY)
            put(ws, r, 4, m["gmv3"], MONEY)
            put(ws, r, 5, m["cost"], MONEY)
            put(ws, r, 6, m["videos"], COUNT)
            put(ws, r, 7, m["exp"], COUNT)
            put(ws, r, 8, f"=IFERROR(C{r}/E{r},0)", RATIO)
            r += 1
    for col, w in zip("ABCDEFGH", (12, 16, 16, 18, 12, 10, 12, 9)):
        ws.column_dimensions[col].width = w

    # 趋势图：Excel 原生柱状图要有「日期×账号」矩阵做数据源，平铺表喂不了，
    # 所以在平铺表下方搭一块辅助数据区（A:C），图表放表格右侧顶部（J2）便于直接看。
    # 注：平铺表是「天×账号」行，不能用 len(daily) 算空行——那只是天数。
    # r 才是循环结束后的下一个空行。
    h0 = r + 1
    ws.cell(row=h0, column=1, value="图表数据（辅助）").font = Font(
        name="Arial", size=10, bold=True, color="1E40AF")
    put(ws, h0 + 1, 1, "日期")
    for j, n in enumerate(acc_names):
        put(ws, h0 + 1, 2 + j, n)
    for i, d in enumerate(daily):
        put(ws, h0 + 2 + i, 1, d["date"])
        for j, n in enumerate(acc_names):
            m = d["per"].get(n)
            put(ws, h0 + 2 + i, 2 + j, m["gmv"] if m else None, MONEY)

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = f"每日全站GMV（元）｜{daily[0]['date']} ~ {daily[-1]['date']}"
    chart.y_axis.title = "GMV(元)"
    chart.x_axis.title = "日期"
    chart.height, chart.width = 8, 18
    data = Reference(ws, min_col=2, max_col=1 + len(acc_names),
                     min_row=h0 + 1, max_row=h0 + 1 + len(daily))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=h0 + 2, max_row=h0 + 1 + len(daily)))
    chart.x_axis.delete = False              # 不关掉，否则 Excel 里轴会隐掉
    chart.y_axis.delete = False
    ws.add_chart(chart, "J2")


def build_video_sheet(ws, videos):
    heads = ["账号", "视频名称", "计划", "是否AI", "曝光量", "投放消耗(元)",
             "3s完播率", "5s完播率", "完播率", "投产比ROI", "3日累计GMV(元)", "全站GMV(元)"]
    for i, h in enumerate(heads, 1):
        ws.cell(row=1, column=i, value=h)
    style_header(ws, len(heads))
    r = 2
    for acct, vs in videos.items():
        for v in vs:
            put(ws, r, 1, acct)
            put(ws, r, 2, v["name"])
            put(ws, r, 3, v["campaign"])
            put(ws, r, 4, "AI" if v["ai"] else "")
            put(ws, r, 5, v["exp"], COUNT)
            put(ws, r, 6, v["cost"], MONEY)
            put(ws, r, 7, v["r3"] / 100, RATE)
            put(ws, r, 8, v["r5"] / 100, RATE)
            put(ws, r, 9, v["rend"] / 100, RATE)
            put(ws, r, 10, f"=IFERROR(L{r}/F{r},0)", RATIO)
            put(ws, r, 11, v["gmv3"], MONEY)
            put(ws, r, 12, v["gmv"], MONEY)
            r += 1
    ws.auto_filter.ref = f"A1:L{r - 1}"
    for col, w in zip("ABCDEFGHIJKL", (14, 46, 30, 8, 10, 12, 10, 10, 10, 10, 14, 14)):
        ws.column_dimensions[col].width = w


def _merged_text(ws, r, text, font, ncol=8, cjk_per_line=70):
    """把长句写进 A:ncol 的合并单元格并设行高。合并单元格不会自动撑高，只能手算。"""
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
    c = ws.cell(row=r, column=1, value=text)
    c.font = font
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = max(15, (len(text) // cjk_per_line + 1) * 15)


def build_insights_sheet(ws, ins, day=""):
    """运营建议。结论在前，再做四象限汇总和重点行动名单。"""
    r = 1
    ws.cell(row=r, column=1, value=f"运营建议（报告日 {day}，基于当日数据自动生成）").font = Font(
        name="Arial", size=12, bold=True, color="1E40AF")
    r += 2
    ws.cell(row=r, column=1, value="结论").font = Font(name="Arial", size=10, bold=True)
    r += 1
    for t in ins["conclusions"]:
        _merged_text(ws, r, t, BODY_FONT)
        r += 1
    r += 1

    ws.cell(row=r, column=1, value="预算分配建议").font = Font(name="Arial", size=10, bold=True)
    r += 1
    heads = ["账号", "建议动作", "视频数", "花费合计(元)", "动作含义"]
    for i, h in enumerate(heads, 1):
        ws.cell(row=r, column=i, value=h)
    style_header(ws, len(heads), row=r, freeze=False)
    r += 1
    act_meta = {"放大": "ROI 达标且已是主力花费，可放量",
                "加量测试": "ROI 达标但花费小，小步加预算看能否放量",
                "缩减优化": "高花费但 ROI 不达标，优先缩量或换素材",
                "观察": "低花费且 ROI 不达标，暂不动"}
    for acc, q in ins["quadrant"].items():
        for k in ("放大", "加量测试", "缩减优化", "观察"):
            lst = q["groups"][k]
            if not lst:
                continue
            put(ws, r, 1, acc)
            put(ws, r, 2, k)
            put(ws, r, 3, len(lst), COUNT)
            put(ws, r, 4, round(sum(v["cost"] for v in lst), 2), MONEY)
            put(ws, r, 5, act_meta[k])
            r += 1
    note = (f"达标线＝该账号整体ROI（" +
            " / ".join(f"{a} {q['base_roi']:.2f}" for a, q in ins["quadrant"].items()) +
            f"）；高花费线＝{ins['high_cost']} 元（取有效视频花费中位数）；"
            f"仅统计花费≥{ins['active_cost']} 元的视频。")
    _merged_text(ws, r, note, NOTE_FONT)
    r += 2

    _merged_text(ws, r,
                 f"重点行动名单（花费≥{ins['active_cost']} 元且 ROI<1.5，按花费降序，"
                 f"共 {ins['low_roi']['count']} 条 / {ins['low_roi']['cost']} 元，"
                 f"下表列前 {len(ins['low_roi']['items'])} 条）",
                 Font(name="Arial", size=10, bold=True))
    r += 1
    heads2 = ["账号", "视频名称", "计划", "花费(元)", "曝光", "全站GMV(元)", "投产比ROI", "完播率"]
    for i, h in enumerate(heads2, 1):
        ws.cell(row=r, column=i, value=h)
    style_header(ws, len(heads2), row=r, freeze=False)
    r += 1
    for v in ins["low_roi"]["items"]:
        put(ws, r, 1, v["acc"])
        put(ws, r, 2, v["name"])
        put(ws, r, 3, v["campaign"])
        put(ws, r, 4, v["cost"], MONEY)
        put(ws, r, 5, v["exp"], COUNT)
        put(ws, r, 6, v["gmv"], MONEY)
        put(ws, r, 7, v["roi"], RATIO)
        put(ws, r, 8, v["rend"] / 100, RATE)
        r += 1
    r += 1
    _merged_text(ws, r, (
        f"异常排查：{ins['waste_zero_gmv']['count']} 个视频花费≥{ins['active_cost']} 元却 0 GMV，"
        f"合计 {ins['waste_zero_gmv']['cost']} 元。本区间有效投放视频 {ins['active_count']} 个，"
        f"占全部花费 {ins['active_share']}%。完播率相关性：3s完播 {ins['corr']['r3_roi']}、"
        f"作品完播 {ins['corr']['rend_roi']}（n={ins['corr']['n']}），均无正相关，不作为砍素材依据。"),
        NOTE_FONT)
    for col, w in zip("ABCDEFGH", (16, 46, 30, 12, 10, 14, 10, 10)):
        ws.column_dimensions[col].width = w


def build_editor_sheet(ws, ed, day=""):
    """给剪辑的建议：当日跑得好的视频名单，供复刻打法。"""
    r = 1
    ws.cell(row=r, column=1, value=f"给剪辑的建议（报告日 {day}）").font = Font(
        name="Arial", size=12, bold=True, color="1E40AF")
    r += 2
    _merged_text(ws, r, "当日有 GMV 的视频按全站GMV降序取前 "
                        f"{len(ed['items'])} 条。不按 ROI 排——花费极小的视频 ROI 会虚高，"
                        "判读以 GMV 为主、ROI 为辅。", BODY_FONT)
    r += 1
    for n in ed["notes"]:
        _merged_text(ws, r, n, BODY_FONT)
        r += 1
    r += 1

    heads = ["视频名称", "计划", "AI", "曝光", "花费(元)", "全站GMV(元)", "投产比ROI", "完播率"]
    for i, h in enumerate(heads, 1):
        ws.cell(row=r, column=i, value=h)
    style_header(ws, len(heads), row=r)
    r += 1
    for v in ed["items"]:
        put(ws, r, 1, v["name"])
        put(ws, r, 2, v["campaign"])
        put(ws, r, 3, "AI" if v.get("ai") else "")
        put(ws, r, 4, v["exp"], COUNT)
        put(ws, r, 5, v["cost"], MONEY)
        put(ws, r, 6, v["gmv"], MONEY)
        put(ws, r, 7, v["roi"], RATIO)
        put(ws, r, 8, v["rend"] / 100, RATE)
        r += 1
    for col, w in zip("ABCDEFGH", (46, 30, 8, 10, 12, 14, 10, 10)):
        ws.column_dimensions[col].width = w


def build_notes_sheet(ws, src, payload):
    lines = [
        ("数据说明", ""),
        ("来源文件", os.path.basename(src)),
        ("报告日", payload["report_day"] + "（主体数据；取数据里最新的一天，但排除没跑完的今天）"),
        ("对比日", str(payload.get("prev_day") or "无")),
        ("排除今天", "是（当天数据只累计到导出时刻，会拉低花费、抬高ROI，不参与汇总）"
                    if payload.get("excluded_today") else "否（数据里没有今天）"),
        ("数据区间", payload["range"] + "（每日趋势列全部日期，供参照）"),
        ("生成时间", payload["generated"]),
        ("", ""),
        ("口径", ""),
        ("账号总GMV", "取「全站GMV」列汇总（另并列「AI视频3日累计GMV」）"),
        ("发布视频数", "按视频ID去重计数；同一视频出现在多个计划/日期只算一个"),
        ("不含视频ID=0", "视频ID=0 是罗盘「无视频」哨兵值，不计入视频数、也不进视频明细；"
                        "它带的钱仍计入账号汇总，故视频表合计小于账号合计"),
        ("合计视频数", "全区间去重值，同样不含视频ID=0"),
        ("单条平均GMV", "总GMV ÷ 发布视频数"),
        ("投产比ROI", "全站GMV ÷ 投放消耗；与平台自带的「全站ROI」口径不同，不可混用"),
        ("完播率", "按曝光数加权（明细只有比率、无分子分母，无法精确聚合）"),
        ("AI视频口径", "标题末尾为 A，或标题含 j/J（须紧贴话题块之前）。j/J 用位置限定，"
                      "避免误召回素材文件名里的 dj（如 W8.12-dj2.mp4）；A 限定在整条末尾，"
                      "避免误召回 AIGC、10A级 这类正文里的字母。"),
        ("比率求和", "ROI、完播率等比率指标不可对明细行求和或求平均，须先加总分子分母再除"),
        ("", ""),
        ("已知脏点", ""),
        ("视频名为数字ID", "部分视频名未同步，直接显示为数字ID，非错误"),
        ("有消耗无GMV", "花花好物清单个别日期有消耗但GMV为0，导致该行ROI为0"),
        ("特殊字符", "个别视频名含英文引号或换行"),
    ]
    for i, (k, v) in enumerate(lines, 1):
        a = ws.cell(row=i, column=1, value=k)
        b = ws.cell(row=i, column=2, value=v)
        a.font = Font(name="Arial", size=10, bold=(v == "" and k != ""))
        b.font = NOTE_FONT
        b.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 70


def main():
    src = sys.argv[1]
    out = sys.argv[2]
    _, body = L.read_rows(src)
    # 日报口径：主体 = 数据里最新的一天，排除没跑完的今天；趋势仍列全部日期做参照。
    all_dates = sorted({r["时间"] for r in body})
    day, prev, excluded_today = L.pick_report_day(all_dates)
    rows_day = [r for r in body if r["时间"] == day]
    print(f"报告日 {day}" + (f" | 前一日 {prev}" if prev else "")
          + (" | 已排除今天（数据未跑完）" if excluded_today else ""))

    by_acc, by_day_acc = {}, {}
    for r in rows_day:
        by_acc.setdefault(r["账户名称"], []).append(r)
    for r in body:
        by_day_acc.setdefault((r["时间"], r["账户名称"]), []).append(r)

    accounts = []
    for name, rows in sorted(by_acc.items(), key=lambda kv: -len(kv[1])):
        m = L.acc_metrics(rows)
        m["name"] = name
        accounts.append(m)
    acc_names = sorted({r["账户名称"] for r in body})   # 趋势列全部账号
    daily = [{"date": d,
              "per": {n: L.acc_metrics(by_day_acc[(d, n)])
                      for n in acc_names if (d, n) in by_day_acc}}
             for d in all_dates]

    videos = {}
    for name, rows in by_acc.items():
        grp = {}
        for r in rows:
            if r["视频ID"] == "0":  # 无视频哨兵值，不是真视频
                continue
            grp.setdefault(r["视频ID"], []).append(r)
        vlist = []
        for vid, vrows in grp.items():
            m = L.acc_metrics(vrows)
            m["acc"] = name
            m["name"] = vrows[0]["视频名称"]
            m["ai"] = L.is_ai(m["name"])
            m["campaign"] = vrows[0]["广告计划名称"]
            vlist.append(m)
        vlist.sort(key=lambda x: -x["gmv"])
        videos[name] = vlist

    payload = {"report_day": day, "prev_day": prev, "excluded_today": excluded_today,
               "range": f"{all_dates[0]} ~ {all_dates[-1]}", "days": len(all_dates),
               "generated": L.datetime.now(tz=L.CST).strftime("%Y-%m-%d %H:%M")}

    # 自检：账号合计 = 明细合计（报告日）
    overall = L.acc_metrics(rows_day)
    assert abs(sum(a["gmv"] for a in accounts) - overall["gmv"]) < 0.05
    assert abs(sum(a["cost"] for a in accounts) - overall["cost"]) < 0.05

    wb = Workbook()
    build_account_sheet(wb.active, accounts, overall, day)
    wb.active.title = "账号汇总"
    ins = L.build_insights(videos, accounts, overall)
    build_insights_sheet(wb.create_sheet("运营建议"), ins, day)
    build_editor_sheet(wb.create_sheet("给剪辑的建议"), ins["editor"], day)
    build_daily_sheet(wb.create_sheet("每日趋势"), daily, acc_names)
    build_video_sheet(wb.create_sheet("视频明细"), videos)
    build_notes_sheet(wb.create_sheet("数据说明"), src, payload)
    wb.save(out)
    print(f"{day} | {len(accounts)} 账号 / {sum(a['videos'] for a in accounts)} 视频 / "
          f"趋势 {len(daily)} 天 -> {out}")


if __name__ == "__main__":
    main()
