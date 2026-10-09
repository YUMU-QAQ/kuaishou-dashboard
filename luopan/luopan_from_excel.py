#!/usr/bin/env python3
"""读罗盘导出的自定义报表明细 xlsx，出账号/视频两级聚合，生成单文件 HTML 报告。

数据源列：时间 账户ID 账户名称 视频ID 视频名称 代理商ID 代理商名称 广告计划ID
广告计划名称 花费(元) 曝光数 作品3s完播率 作品5s完播率 作品完播率 3日累计GMV 全站GMV 全站ROI 账户类型
用 XML 直解析：openpyxl read_only 因文件缺 <dimension> 只读到 1 格。
"""
import json
import os
import re
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from xml.etree import ElementTree as ET

CST = timezone(timedelta(hours=8))  # 罗盘日界按北京时间，别用宿主机时区


def pick_report_day(all_dates):
    """报告主体日 = 数据里最新的一天，但排除「今天」。

    排除今天是因为当天数据没跑完（导出时花费只累计到那一刻），拿它当报告主体
    会出一份 ROI、环比都失真的报告。排除后仍取「数据里最新的一天」而不是写死
    「今天减一」，所以导出延迟时不会出空报告。
    """
    today = datetime.now(tz=CST).strftime("%Y-%m-%d")
    days = [d for d in all_dates if d != today]
    if not days:            # 数据里只有今天，只能用它，并标记为不完整
        return all_dates[-1], None, True
    return days[-1], (days[-2] if len(days) > 1 else None), bool(all_dates[-1] == today)

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def read_rows(path):
    z = zipfile.ZipFile(path)
    root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    sheet = root.find("m:sheetData", NS)

    def cell(c):
        if c.get("t") == "inlineStr":
            is_ = c.find("m:is", NS)
            if is_ is None:
                return ""
            return "".join(t.text or "" for t in is_.iter(
                "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
        v = c.find("m:v", NS)
        return v.text if v is not None else ""

    data = []
    for r in sheet:
        d = {}
        for c in r.findall("m:c", NS):
            col = re.match(r"[A-Z]+", c.get("r")).group(0)
            d[col] = cell(c)
        data.append(d)
    hdr = data[0]
    cols = {v: k for k, v in hdr.items()}
    body = []
    for r in data[1:]:
        body.append({name: r.get(col, "") for name, col in cols.items()})
    return cols, body


def fnum(s):
    s = (s or "").replace("%", "").replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def is_ai(title):
    """判定 AI 视频：标题末尾打 A，或标题含 j/J（紧贴话题块之前）。

    j/J 用「后跟空白+#」限定，否则会误召回素材文件名里的 dj（如 W8.12-dj2.mp4）。
    标题里混了 \xa0 不换行空格，先归一化再判末尾。
    """
    t = (title or "").replace("\xa0", " ").strip()
    return t.endswith("A") or re.search(r"[jJ](?=\s*#)", t) is not None


def acc_metrics(rows):
    """一组明细行的合计。完播率按曝光数加权。"""
    cost = sum(fnum(r["花费(元)"]) for r in rows)
    exp = sum(fnum(r["曝光数"]) for r in rows)
    gmv = sum(fnum(r["全站GMV"]) for r in rows)
    gmv3 = sum(fnum(r["3日累计GMV"]) for r in rows)
    vids = {r["视频ID"] for r in rows if r["视频ID"] != "0"}  # 0 = 无视频哨兵值
    ai_rows = [r for r in rows if is_ai(r["视频名称"])]
    wr = {}
    for k in ("作品3s完播率", "作品5s完播率", "作品完播率"):
        num = sum(fnum(r[k]) * fnum(r["曝光数"]) for r in rows)
        wr[k] = round(num / exp, 2) if exp else 0.0
    return {
        "cost": round(cost, 2), "exp": exp, "gmv": round(gmv, 2),
        "gmv3": round(gmv3, 2), "videos": len(vids),
        "avg_gmv": round(gmv / len(vids), 2) if vids else 0.0,
        "roi": round(gmv / cost, 2) if cost else 0.0,
        # AI 视频口径：标题末尾 A 或含 j/J（见 is_ai）
        "ai_videos": len({r["视频ID"] for r in ai_rows if r["视频ID"] != "0"}),
        "ai_cost": round(sum(fnum(r["花费(元)"]) for r in ai_rows), 2),
        "ai_gmv": round(sum(fnum(r["全站GMV"]) for r in ai_rows), 2),
        "ai_gmv3": round(sum(fnum(r["3日累计GMV"]) for r in ai_rows), 2),
        "r3": wr["作品3s完播率"], "r5": wr["作品5s完播率"], "rend": wr["作品完播率"],
    }


def _pearson(xs, ys):
    """皮尔逊相关系数，手算避免引入 numpy 依赖。"""
    n = len(xs)
    if n < 3:
        return 0.0
    mx, my = sum(xs) / n, sum(ys) / n
    num = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    dx = sum((x - mx) ** 2 for x in xs) ** 0.5
    dy = sum((y - my) ** 2 for y in ys) ** 0.5
    return num / (dx * dy) if dx and dy else 0.0


# 决策阈值。都来自实测分布：花费中位数极低（多数视频几乎没花钱），
# 97 个花费≥5 的视频就吃掉 98.7% 的花费，所以决策分析只在「有效花费」子集上做。
ACTIVE_COST = 5.0    # 花费≥5 元才纳入决策分析
LOW_ROI = 1.5        # 低于此视为低效投放（用于浪费排查）
MAX_LIST = 15        # 各名单最多列这么多条
EDITOR_TOP = 10      # 给剪辑同学看的表现最好视频条数


def build_insights(videos, accounts, overall):
    """把明细转成运营可直接行动的结论。阈值和数据一起输出，便于复核。"""
    V = [v for vs in videos.values() for v in vs]
    tot_cost = sum(v["cost"] for v in V)
    tot_gmv = sum(v["gmv"] for v in V)
    act = [v for v in V if v["cost"] >= ACTIVE_COST]
    act_share = sum(v["cost"] for v in act) / tot_cost * 100 if tot_cost else 0.0

    # 高花费线 = 有效视频花费的中位数，随数据走，不写死
    costs = sorted(v["cost"] for v in act)
    high_cost = round(costs[len(costs) // 2], 1) if costs else 0.0

    # 帕累托：多少视频贡献 80% GMV
    by_gmv = sorted(V, key=lambda x: -x["gmv"])
    cum, pareto_n = 0.0, len(by_gmv)
    for i, v in enumerate(by_gmv, 1):
        cum += v["gmv"]
        if tot_gmv and cum / tot_gmv >= 0.8:
            pareto_n = i
            break

    # 预算四象限（按账号各自的整体 ROI 作为达标线）
    quad = {}
    for a in accounts:
        base = a["roi"]
        q = {"放大": [], "加量测试": [], "缩减优化": [], "观察": []}
        for v in act:
            if v["acc"] != a["name"]:
                continue
            eff = v["roi"] >= base
            hv = v["cost"] >= high_cost
            key = "放大" if (eff and hv) else "加量测试" if eff else "缩减优化" if hv else "观察"
            q[key].append(v)
        for k in q:
            q[k].sort(key=lambda x: -x["cost"])
        quad[a["name"]] = {"base_roi": base, "groups": q}

    # 浪费：有效花费却没回报
    waste = sorted([v for v in act if v["gmv"] == 0], key=lambda x: -x["cost"])
    low_roi = sorted([v for v in act if v["roi"] < LOW_ROI], key=lambda x: -x["cost"])

    # 完播率能不能预测 ROI？实测不能，必须如实说，别拿它当砍素材的依据
    corr = {"r3_roi": round(_pearson([v["r3"] for v in act], [v["roi"] for v in act]), 3),
            "rend_roi": round(_pearson([v["rend"] for v in act], [v["roi"] for v in act]), 3),
            "n": len(act)}

    return {
        "active_cost": ACTIVE_COST, "active_count": len(act),
        "active_share": round(act_share, 1), "high_cost": high_cost,
        "total_cost": round(tot_cost, 2), "total_gmv": round(tot_gmv, 2),
        "total_videos": len(V), "pareto_n": pareto_n,
        "pareto_share": round(pareto_n / len(V) * 100, 1) if V else 0,
        "quadrant": quad,
        "waste_zero_gmv": {"count": len(waste), "cost": round(sum(v["cost"] for v in waste), 2),
                          "items": waste[:MAX_LIST]},
        "low_roi": {"count": len(low_roi), "cost": round(sum(v["cost"] for v in low_roi), 2),
                    "share": round(sum(v["cost"] for v in low_roi) / tot_cost * 100, 1) if tot_cost else 0,
                    "items": low_roi[:MAX_LIST]},
        "corr": corr,
        "editor": _editor_advice(V),
        "conclusions": _conclusions(V, act, tot_cost, pareto_n, waste, low_roi, corr),
    }


def _conclusions(V, act, tot_cost, pareto_n, waste, low_roi, corr):
    """生成结论句。只写数据支持得住的话。"""
    out = []
    if V:
        out.append(
            f"花费高度集中：{len(act)} 个视频（占 {len(act)/len(V)*100:.1f}%）吃掉 "
            f"{sum(v['cost'] for v in act)/tot_cost*100:.1f}% 的花费；GMV 更集中，"
            f"前 {pareto_n} 个视频（{pareto_n/len(V)*100:.1f}%）贡献 80%。"
            f"日常只需盯这几十个，不必逐条看全量。")
    if waste:
        out.append(
            f"零转化浪费：{len(waste)} 个视频花费≥{ACTIVE_COST:.0f} 元却 0 GMV，"
            f"合计 {sum(v['cost'] for v in waste):.2f} 元。优先关停，核对是否计划/出价问题。")
    if low_roi:
        extra = [v for v in low_roi if v["gmv"] > 0]
        tail = (f"其中 {len(extra)} 个有成交但仍低于 {LOW_ROI}。"
                if extra else "这批与上面「零转化」完全重合，不必重复处理。")
        out.append(
            f"低效投放：{len(low_roi)} 个视频 ROI<{LOW_ROI}，合计花费 "
            f"{sum(v['cost'] for v in low_roi):.2f} 元（占 {sum(v['cost'] for v in low_roi)/tot_cost*100:.1f}%）。"
            f"按名单逐条缩量或换素材，比全面降预算有效。" + tail)
    out.append(
        f"完播率与 ROI 无正相关（3s完播率 corr={corr['r3_roi']}，作品完播率 corr={corr['rend_roi']}，"
        f"n={corr['n']}）。不要用「完播率低」直接砍素材，完播率只当内容健康度参考，投产决策看 ROI。")
    return out


def _editor_advice(V):
    """给剪辑同学：拉出当日跑出 GMV 的头部视频，供复刻打法。

    只按全站GMV 排，不按 ROI——花费极小的视频 ROI 会虚高到几十上百，
    当排名依据会误导；ROI 只在名单里并列展示，判读以 GMV 为主。
    """
    top = sorted([v for v in V if v["gmv"] > 0], key=lambda x: -x["gmv"])[:EDITOR_TOP]
    camps = Counter(v["campaign"] for v in top)
    ai_n = sum(1 for v in top if v.get("ai"))
    notes = []
    if top:
        name, cnt = camps.most_common(1)[0]
        notes.append(
            f"当日有 GMV 的视频 {sum(1 for v in V if v['gmv'] > 0)} 条，"
            f"头部 {len(top)} 条里有 {cnt} 条来自「{name}」，AI 视频 {ai_n} 条。")
        notes.append(
            "复刻时看标题打法（价格对比 / 销量背书 / 痛点词），不要看完播率高低——"
            "完播率与投产在本数据里无正相关。")
    return {"items": top, "camps": camps.most_common(3), "ai_count": ai_n, "notes": notes}


def main():
    src = sys.argv[1]
    out = sys.argv[2]
    cols, body = read_rows(src)
    print("rows:", len(body), "cols:", list(cols.keys())[:5], "...")

    # 日报口径：报告主体 = 数据里最新的一天，排除没跑完的今天（见 pick_report_day）。
    # 趋势仍保留全部日期做参照。
    all_dates = sorted({r["时间"] for r in body})
    day, prev, excluded_today = pick_report_day(all_dates)
    rows_day = [r for r in body if r["时间"] == day]
    print(f"报告日 {day}" + (f" | 前一日 {prev}" if prev else " | 无前一日")
          + (" | 已排除今天（数据未跑完）" if excluded_today else ""))

    by_acc = defaultdict(list)
    by_day_acc = defaultdict(list)
    for r in body:
        by_day_acc[(r["时间"], r["账户名称"])].append(r)
    for r in rows_day:
        by_acc[r["账户名称"]].append(r)

    accounts = []
    for name, rows in sorted(by_acc.items(), key=lambda kv: -len(kv[1])):
        m = acc_metrics(rows)
        m["name"] = name
        accounts.append(m)

    # 每日趋势：日期 x 账号（全部日期，不只报告日）
    acc_names = sorted({r["账户名称"] for r in body})
    daily = []
    for d in all_dates:
        per = {}
        for n in acc_names:
            rows = by_day_acc.get((d, n))
            if rows:
                per[n] = acc_metrics(rows)
        daily.append({"date": d, "per": per})

    # 视频维度明细（报告日内，按账号 x 视频ID 聚合）。photo_id=0 不是真视频，不进视频表；
    # 它带的钱仍计入账号维度，所以视频表合计会比账号合计少这一块（见下方断言）。
    videos = {}
    for name, rows in by_acc.items():
        grp = defaultdict(list)
        for r in rows:
            if r["视频ID"] == "0":
                continue
            grp[r["视频ID"]].append(r)
        vlist = []
        for vid, vrows in grp.items():
            m = acc_metrics(vrows)
            m["id"] = vid
            m["acc"] = name
            m["name"] = vrows[0]["视频名称"]
            m["ai"] = is_ai(m["name"])
            m["campaign"] = vrows[0]["广告计划名称"]
            vlist.append(m)
        vlist.sort(key=lambda x: -x["gmv"])
        videos[name] = vlist

    overall = acc_metrics(rows_day)
    prev_overall = acc_metrics([r for r in body if r["时间"] == prev]) if prev else None

    # 自检：账号合计 = 明细合计；视频表 + 无视频行 = 账号合计
    assert abs(sum(a["gmv"] for a in accounts) - overall["gmv"]) < 0.05, "账号GMV合计对不上"
    assert abs(sum(a["cost"] for a in accounts) - overall["cost"]) < 0.05, "账号消耗对不上"
    assert abs(sum(a["exp"] for a in accounts) - overall["exp"]) < 1, "账号曝光对不上"
    # 容差随视频数放宽：每行各自四舍五入到分，多行求和会累积亚分误差。
    # 真丢行是整条量级，这个阈值照样拦得住。
    for a in accounts:
        tol = 0.05 + 0.01 * a["videos"]
        id0 = [r for r in by_acc[a["name"]] if r["视频ID"] == "0"]
        vs = sum(v["gmv"] for v in videos[a["name"]]) + sum(fnum(r["全站GMV"]) for r in id0)
        assert abs(vs - a["gmv"]) < tol, f"视频GMV求和 != 账号 {a['name']}"
        vc = sum(v["cost"] for v in videos[a["name"]]) + sum(fnum(r["花费(元)"]) for r in id0)
        assert abs(vc - a["cost"]) < tol, f"视频消耗求和 != 账号 {a['name']}"

    payload = {
        "report_day": day,
        "prev_day": prev,
        "excluded_today": excluded_today,
        "range": f"{all_dates[0]} ~ {all_dates[-1]}",
        "days": len(all_dates),
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "accounts": accounts,
        "daily": daily,
        "videos": videos,
        "overall": overall,
        "prev": prev_overall,
        "insights": build_insights(videos, accounts, overall),
    }
    with open(out, "w", encoding="utf-8") as f:
        f.write(render(payload))
    print(f"账号合计校验通过 | {len(accounts)} 账号 / 共 {sum(a['videos'] for a in accounts)} 视频 "
          f"| 总GMV {overall['gmv']} / 消耗 {overall['cost']} / ROI {overall['roi']}")
    print("->", out)


HTML = """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>罗盘投放日报</title>
<style>
:root{color-scheme:light;
--primary:#1E40AF;--secondary:#3B82F6;--accent:#D97706;
--bg:#F8FAFC;--fg:#1E3A8A;--card:#fff;--muted:#E9EEF6;--muted-fg:#475569;
--border:#DBEAFE;--ring:#1E40AF;
--radius:12px;--shadow:0 1px 2px rgba(30,58,138,.06),0 1px 3px rgba(30,58,138,.04)}
*{box-sizing:border-box}
body{margin:0;padding:24px;background:var(--bg);color:var(--fg);
font:14px/1.5 -apple-system,"Segoe UI",Roboto,"Helvetica Neue","Microsoft YaHei",sans-serif;
font-variant-numeric:tabular-nums;-webkit-font-smoothing:antialiased}
h1{font-size:20px;font-weight:700;margin:0 0 4px;letter-spacing:-.01em}
.sub{color:var(--muted-fg);font-size:12px;margin-bottom:20px}
h2{font-size:14px;font-weight:600;margin:26px 0 10px;color:var(--fg);letter-spacing:.01em}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:8px}
.kpi{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);
padding:14px 16px;box-shadow:var(--shadow);position:relative;overflow:hidden}
.kpi:before{content:"";position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--secondary)}
.kpi.accent:before{background:var(--accent)}
.kpi .k{font-size:11px;font-weight:600;letter-spacing:.04em;text-transform:uppercase;color:var(--muted-fg)}
.kpi .v{font-size:24px;font-weight:700;margin-top:6px;color:var(--primary);line-height:1.15;
font-variant-numeric:tabular-nums}
table{border-collapse:collapse;width:100%;background:var(--card);font-variant-numeric:tabular-nums}
th,td{padding:9px 12px;text-align:right;border-bottom:1px solid var(--muted);white-space:nowrap}
th{background:var(--muted);font-weight:600;color:var(--muted-fg);font-size:11px;letter-spacing:.03em;
text-transform:uppercase;text-align:right}
th:first-child,td:first-child{text-align:left}
tbody tr:last-child td{border-bottom:0}
tbody tr{transition:background 150ms ease}
tbody tr:hover td{background:#F1F5FF}
tr.acct{cursor:pointer}
tr.acct .caret{color:var(--secondary);display:inline-block;transition:transform 200ms ease;
margin-right:6px;font-size:10px}
tr.acct.open .caret{transform:rotate(90deg)}
.detail td{padding:0;background:#FCFDFF}
.detail .wrap{padding:12px 14px 16px}
.vname{max-width:380px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:inline-block;
vertical-align:bottom;text-align:left;color:var(--fg)}
.hi{color:#B91C1C;font-weight:600}
.lo{color:#15803D;font-weight:600}
.go-down:after{content:"▼ ";color:#B91C1C}
.go-up:after{content:"▲ ";color:#15803D}
th.sortable{cursor:pointer;user-select:none;transition:color 150ms ease}
th.sortable:hover{color:var(--primary)}
th .ar{color:#93A6C4;font-size:9px}
th.asc .ar:before{content:"▲";color:var(--primary)}
th.desc .ar:before{content:"▼";color:var(--primary)}
.toolbar{display:flex;gap:10px;align-items:center;margin:0 0 10px}
.toolbar input{padding:7px 12px;border:1px solid var(--border);border-radius:9px;font:inherit;
font-size:13px;width:220px;background:var(--card);color:var(--fg);
transition:border-color 150ms ease,box-shadow 150ms ease}
.toolbar input:focus{outline:none;border-color:var(--primary);box-shadow:0 0 0 3px rgba(30,64,175,.15)}
.legend{display:flex;gap:16px;font-size:12px;color:var(--muted-fg);margin:6px 2px}
.legend i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:6px;vertical-align:middle}
.chart{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);
padding:14px 10px 6px;box-shadow:var(--shadow)}
/* 表格横向滚动包裹层：不放这里，11 列的账号表会把整页撑宽，下面的每日趋势被挤到屏幕外 */
.tablebox{overflow-x:auto;border:1px solid var(--border);border-radius:var(--radius);
background:var(--card);box-shadow:var(--shadow)}
.tablebox>table{border:0;border-radius:0;width:auto;min-width:100%}
.ins-cards{display:grid;gap:10px;margin:0 0 6px}
.ins-card{background:var(--card);border:1px solid var(--border);border-left:3px solid var(--accent);
border-radius:var(--radius);padding:12px 14px;box-shadow:var(--shadow);font-size:13px;color:var(--fg)}
.ins-card.warn{border-left-color:#D97706;background:#FFFBEB}
.ins-sub{font-size:12px;color:var(--muted-fg);font-weight:400;margin-left:6px}
.badge{display:inline-block;font-size:11px;font-weight:600;padding:2px 9px;border-radius:20px;white-space:nowrap}
.kpi .dt{font-size:11px;margin-top:4px;color:var(--muted-fg)}
.kpi .dt.up{color:#15803D}
.kpi .dt.down{color:#B91C1C}
.b-add{background:#DCFCE7;color:#166534}
.b-test{background:#DBEAFE;color:#1E3A8A}
.b-cut{background:#FEE2E2;color:#991B1B}
.b-watch{background:#E9EEF6;color:#475569}
:focus-visible{outline:2px solid var(--ring);outline-offset:2px}
@media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}}
</style></head><body>
<h1 id="h1">罗盘投放日报</h1>
<div class="sub" id="sub"></div>
<div class="kpis" id="kpis"></div>
<h2>运营建议 <span class="ins-sub">基于报告日数据自动生成，阈值见下方说明</span></h2>
<div class="ins-cards" id="insConcl"></div>
<div class="ins-cards" id="insWarn"></div>
<div class="tablebox" style="margin-top:10px"><table id="quad"><thead><tr>
<th>账号</th><th>建议动作</th><th>视频数</th><th>花费合计(元)</th><th>动作含义</th>
</tr></thead><tbody></tbody></table></div>
<div class="sub" id="quadNote" style="margin:8px 2px 0"></div>
<h2>重点行动名单 <span class="ins-sub" id="actNote"></span></h2>
<div class="tablebox"><table id="actT"><thead><tr>
<th>账号</th><th>视频名称</th><th>计划</th><th>花费(元)</th><th>曝光</th>
<th>全站GMV(元)</th><th>投产比ROI</th><th>完播率</th></tr></thead><tbody></tbody></table></div>
<h2>给剪辑的建议 <span class="ins-sub" id="edRule"></span></h2>
<div class="ins-cards" id="edNotes"></div>
<div class="tablebox" style="margin-top:10px"><table id="edT"><thead><tr>
<th>视频名称</th><th>计划</th><th>AI</th><th>曝光</th><th>花费(元)</th>
<th>全站GMV(元)</th><th>投产比ROI</th><th>完播率</th></tr></thead><tbody></tbody></table></div>
<h2 id="acctH">账号维度汇总</h2>
<div class="toolbar"><input id="acctFilter" placeholder="筛选账号名称…"></div>
<div class="tablebox"><table id="acct"><thead><tr>
<th class="sortable" data-k="name">账号<span class="ar"></span></th>
<th class="sortable" data-k="gmv">总GMV(全站)<span class="ar"></span></th>
<th class="sortable" data-k="gmv3">AI视频3日累计GMV<span class="ar"></span></th>
<th class="sortable" data-k="ai_videos">AI视频数<span class="ar"></span></th>
<th class="sortable" data-k="videos">发布视频数<span class="ar"></span></th>
<th class="sortable" data-k="avg_gmv">单条平均GMV<span class="ar"></span></th>
<th class="sortable" data-k="exp">总曝光量<span class="ar"></span></th>
<th class="sortable" data-k="cost">总投放消耗<span class="ar"></span></th>
<th class="sortable" data-k="roi">投产比ROI<span class="ar"></span></th>
<th class="sortable" data-k="r3">3s完播率<span class="ar"></span></th>
<th class="sortable" data-k="r5">5s完播率<span class="ar"></span></th>
<th class="sortable" data-k="rend">完播率<span class="ar"></span></th></tr></thead><tbody></tbody></table></div>
<h2>每日趋势 <span class="ins-sub">全部日期，供参照</span></h2>
<div class="chart"><div class="legend" id="legend"></div><div id="chart"></div></div>
<div class="tablebox"><table id="daily"><thead><tr>
<th>日期</th><th>账号</th><th>全站GMV</th><th>3日累计GMV</th><th>消耗</th>
<th>视频数</th><th>曝光</th><th>ROI</th></tr></thead><tbody></tbody></table></div>
<h2>视频维度明细 <span style="font-weight:400;color:#999;font-size:12px">（点击账号行展开）</span></h2>
<div id="vtables"></div>
<script>const DATA=__DATA__;
const f=n=>n.toLocaleString("zh-CN",{minimumFractionDigits:2,maximumFractionDigits:2});
const i=n=>n.toLocaleString("zh-CN");
const p=n=>n.toFixed(2)+"%";
function fmtGmv(n){return n>=10000?(n/10000).toFixed(2)+"万":f(n);}
document.getElementById("h1").textContent="罗盘投放日报 · "+DATA.report_day+"（账号 + 视频两级）";
document.getElementById("sub").textContent=
  (DATA.prev_day?("对比前一日 "+DATA.prev_day+" · "):"")+
  "数据区间 "+DATA.range+" · 生成 "+DATA.generated+
  (DATA.excluded_today?" · 已排除今天（数据未跑完，只作趋势参照）":"")+
  " · 数据源：磁力罗盘自定义报表明细 · AI视频口径：标题末尾 A，或含 j/J（紧贴话题块之前）";
document.getElementById("acctH").textContent="账号维度汇总（"+DATA.report_day+"）";
const o=DATA.overall,pv=DATA.prev;
const kpis=[["总GMV(全站)",fmtGmv(o.gmv)],["AI视频3日累计GMV",fmtGmv(o.ai_gmv3)],
  ["AI视频数",i(o.ai_videos)],["总投放消耗","¥"+fmtGmv(o.cost)],["综合ROI",o.roi,1],
  ["视频数",i(o.videos)],["总曝光量",i(o.exp)],["3s完播率",p(o.r3)]];
document.getElementById("kpis").innerHTML=kpis.map(k=>
  '<div class="kpi'+(k[2]?" accent":"")+'"><div class="k">'+k[0]+'</div><div class="v">'+k[1]+'</div></div>').join("");
// 环比前一日
(function(){
  if(!pv)return;
  const pairs=[["总GMV(全站)",o.gmv,pv.gmv,0],["AI视频3日累计GMV",o.ai_gmv3,pv.ai_gmv3,0],
    ["AI视频数",o.ai_videos,pv.ai_videos,0],["总投放消耗",o.cost,pv.cost,0],
    ["综合ROI",o.roi,pv.roi,1],["视频数",o.videos,pv.videos,0],
    ["总曝光量",o.exp,pv.exp,0],["3s完播率",o.r3,pv.r3,1]];
  const cards=document.querySelectorAll("#kpis .kpi");
  pairs.forEach((p,i)=>{
    const diff=p[1]-p[2],card=cards[i];if(!card)return;
    let txt,cls="";
    if(Math.abs(diff)<1e-9){txt="持平 较前一日";}
    else{
      const dir=diff>0?"▲":"▼";cls=diff>0?"up":"down";
      txt=dir+(p[3]?" "+Math.abs(diff).toFixed(2):" "+Math.abs(p[2]?diff/p[2]*100:0).toFixed(1)+"%")+" 较前一日";
    }
    const el=document.createElement("div");el.className="dt "+cls;el.textContent=txt;
    card.appendChild(el);
  });
})();
// 运营建议
const ins=DATA.insights;
document.getElementById("insConcl").innerHTML=ins.conclusions.map(t=>
  '<div class="ins-card">'+t+'</div>').join("");
const ACT={"放大":["b-add","ROI 达标且已是主力花费，可放量"],
  "加量测试":["b-test","ROI 达标但花费小，小步加预算看能否放量"],
  "缩减优化":["b-cut","高花费但 ROI 不达标，优先缩量或换素材"],
  "观察":["b-watch","低花费且 ROI 不达标，暂不动"]};
const qb=document.querySelector("#quad tbody");
Object.keys(ins.quadrant).forEach(acc=>{
  const g=ins.quadrant[acc].groups;
  Object.keys(ACT).forEach(k=>{
    const list=g[k]; if(!list.length)return;
    const cost=list.reduce((s,v)=>s+v.cost,0);
    const tr=document.createElement("tr");
    tr.innerHTML='<td>'+acc+'</td><td><span class="badge '+ACT[k][0]+'">'+k+'</span></td>'+
      '<td>'+i(list.length)+'</td><td>'+f(cost)+'</td>'+
      '<td style="text-align:left;color:#475569">'+ACT[k][1]+'</td>';
    qb.appendChild(tr);
  });
});
document.getElementById("quadNote").textContent=
  "达标线＝该账号整体 ROI（"+Object.keys(ins.quadrant).map(a=>
    a+" "+ins.quadrant[a].base_roi.toFixed(2)).join(" / ")+
  "）；高花费线＝"+f(ins.high_cost)+" 元（取有效视频花费中位数）；"+
  "仅统计花费≥"+ins.active_cost+" 元的视频，低花费视频不参与象限。";
document.getElementById("actNote").textContent=
  "花费≥"+ins.active_cost+" 元且 ROI<1.5，按花费降序，列前 "+ins.low_roi.items.length+" 条"+
  "（共 "+ins.low_roi.count+" 条 / "+f(ins.low_roi.cost)+" 元）";
const ab2=document.querySelector("#actT tbody");
ins.low_roi.items.forEach(v=>{
  const tr=document.createElement("tr");
  tr.innerHTML='<td>'+v.acc+'</td>'+
    '<td><span class="vname" title="'+v.name.replace(/"/g,"&quot;")+'">'+v.name+'</span></td>'+
    '<td style="text-align:left;color:#475569">'+v.campaign+'</td>'+
    '<td>'+f(v.cost)+'</td><td>'+i(v.exp)+'</td><td>'+f(v.gmv)+'</td>'+
    '<td class="hi">'+v.roi.toFixed(2)+'</td><td>'+p(v.rend)+'</td>';
  ab2.appendChild(tr);
});
// 给剪辑的建议：拉出跑得好的视频
const ed=ins.editor;
document.getElementById("edRule").textContent=
  "当日有 GMV 的视频按全站GMV降序取前 "+ed.items.length+" 条（不按 ROI 排——花费极小的视频 ROI 会虚高）";
document.getElementById("edNotes").innerHTML=ed.notes.map(t=>
  '<div class="ins-card">'+t+'</div>').join("");
const edb=document.querySelector("#edT tbody");
ed.items.forEach(v=>{
  const tr=document.createElement("tr");
  tr.innerHTML=
    '<td><span class="vname" title="'+v.name.replace(/"/g,"&quot;")+'">'+v.name+'</span></td>'+
    '<td style="text-align:left;color:#475569">'+v.campaign+'</td>'+
    '<td>'+(v.ai?'<span class="badge b-add">AI</span>':'<span style="color:#CBD5E1">—</span>')+'</td>'+
    '<td>'+i(v.exp)+'</td><td>'+f(v.cost)+'</td><td class="hi">'+f(v.gmv)+'</td>'+
    '<td>'+v.roi.toFixed(2)+'</td><td>'+p(v.rend)+'</td>';
  edb.appendChild(tr);
});
document.getElementById("insWarn").innerHTML=
  '<div class="ins-card warn">异常排查：<b>'+i(ins.waste_zero_gmv.count)+'</b> 个视频花费≥'+
  ins.active_cost+' 元却 0 GMV，合计 <b>'+f(ins.waste_zero_gmv.cost)+' 元</b>。'+
  '本区间有效投放视频 '+i(ins.active_count)+' 个（花费≥'+ins.active_cost+' 元），占全部花费 '+
  ins.active_share+'%；高花费线 '+f(ins.high_cost)+' 元（取有效视频花费中位数）。'+
  '完播率相关性：3s完播 '+ins.corr.r3_roi+'、作品完播 '+ins.corr.rend_roi+'（n='+ins.corr.n+
  '），均无正相关，不作为砍素材依据。</div>';
// 账号表：可点击表头排序 + 名称筛选
let sortK="gmv",sortDir=-1,filter="";
function renderAccts(){
  const ab=document.querySelector("#acct tbody");ab.innerHTML="";
  let list=DATA.accounts.filter(a=>a.name.indexOf(filter)>=0);
  list=list.slice().sort((x,y)=>{
    const a=x[sortK],b=y[sortK];
    return typeof a==="string"?sortDir*a.localeCompare(b,"zh"):sortDir*(a-b);
  });
  list.forEach(a=>{
    const tr=document.createElement("tr");tr.className="acct";
    tr.innerHTML='<td><span class="caret">▶</span>'+a.name+'</td>'+
      '<td class="hi">'+f(a.gmv)+'</td><td>'+f(a.ai_gmv3)+'</td><td>'+i(a.ai_videos)+'</td>'+
      '<td>'+i(a.videos)+'</td>'+
      '<td>'+f(a.avg_gmv)+'</td><td>'+i(a.exp)+'</td><td>¥'+f(a.cost)+'</td>'+
      '<td class="'+(a.roi>=1?"hi":"lo")+'">'+a.roi.toFixed(2)+'</td>'+
      '<td>'+p(a.r3)+'</td><td>'+p(a.r5)+'</td><td>'+p(a.rend)+'</td>';
    ab.appendChild(tr);
    const dtr=document.createElement("tr");dtr.className="detail";dtr.hidden=true;
    const td=document.createElement("td");td.colSpan=12;
    // 关键：先在 detached div 里建好，再整体搬进 td。
    // 直接给 <td> 设 innerHTML，解析器在表格单元格上下文会错乱拆嵌套 table——
    // 实测多出 1 行且丢 data-camp，导致筛选抛错、整页脚本中断（每日趋势也一起没了）。
    const wrap=document.createElement("div");wrap.className="wrap";
    wrap.innerHTML=videoTable(a.name);
    td.appendChild(wrap);
    dtr.appendChild(td);ab.appendChild(dtr);
    const vrows=wrap.querySelectorAll("tbody tr"),vcount=wrap.querySelector(".vcount");
    const recount=q=>{let n=0;vrows.forEach(r=>{const ok=(r.dataset.camp||"").indexOf(q)>=0;
      r.hidden=!ok;if(ok)n++;});if(vcount)vcount.textContent=n+" 条视频，按全站GMV降序";};
    wrap.querySelector(".vfilter").oninput=e=>recount(e.target.value.trim());
    recount("");
    tr.onclick=()=>{tr.classList.toggle("open");dtr.hidden=!dtr.hidden;};
  });
  document.querySelectorAll("#acct th.sortable").forEach(th=>{
    th.classList.remove("asc","desc");
    if(th.dataset.k===sortK)th.classList.add(sortDir>0?"asc":"desc");
  });
}
document.querySelectorAll("#acct th.sortable").forEach(th=>{
  th.onclick=()=>{const k=th.dataset.k;
    if(sortK===k)sortDir=-sortDir;else{sortK=k;sortDir=k==="name"?1:-1;}
    renderAccts();};
});
document.getElementById("acctFilter").oninput=e=>{filter=e.target.value.trim();renderAccts();};
renderAccts();
// 每日趋势图（内联 SVG，分组柱：每天各账号全站GMV）
function buildChart(){
  const dates=DATA.daily.map(d=>d.date.slice(5));
  const names=DATA.accounts.map(a=>a.name);
  const colors=["#1E40AF","#D97706","#3B82F6"];
  const series=names.map(n=>DATA.daily.map(d=>d.per[n]?d.per[n].gmv:0));
  const max=Math.max(1,...series.flat());
  const W=780,H=300,pl=64,pr=12,pt=14,pb=34;
  const iw=W-pl-pr,ih=H-pt-pb,yMax=Math.ceil(max/1000)*1000||1;
  const gw=iw/dates.length,bw=Math.min(26,(gw-10)/names.length);
  let s='<svg viewBox="0 0 '+W+' '+H+'" width="100%" style="max-width:820px">';
  for(let t=0;t<=4;t++){
    const v=yMax*t/4,y=pt+ih-ih*t/4;
    s+='<line x1="'+pl+'" y1="'+y+'" x2="'+(W-pr)+'" y2="'+y+'" stroke="#DBEAFE"/>'+
       '<text x="'+(pl-6)+'" y="'+(y+4)+'" font-size="10" fill="#64748B" text-anchor="end">'+
       (v>=10000?(v/10000)+"万":v)+'</text>';
  }
  dates.forEach((d,di)=>{
    const gx=pl+gw*di;
    names.forEach((n,ni)=>{
      const val=series[ni][di],h=ih*val/yMax;
      const x=gx+gw/2-(names.length*bw)/2+ni*bw,y=pt+ih-h;
      s+='<rect x="'+x+'" y="'+y+'" width="'+(bw-3)+'" height="'+h+'" rx="2" fill="'+
         colors[ni%colors.length]+'"><title>'+d+' '+n+': ¥'+val.toFixed(2)+'</title></rect>';
    });
    s+='<text x="'+(gx+gw/2)+'" y="'+(H-12)+'" font-size="10" fill="#64748B" text-anchor="middle">'+d+'</text>';
  });
  document.getElementById("chart").innerHTML=s+'</svg>';
  document.getElementById("legend").innerHTML=names.map((n,ni)=>
    '<span><i style="background:'+colors[ni%colors.length]+'"></i>'+n+'</span>').join("");
}
buildChart();
function videoTable(name){
  const vs=DATA.videos[name];
  let h='<div class="toolbar"><input class="vfilter" placeholder="筛选计划名称…"></div>'+
    '<div class="tablebox"><table><thead><tr><th>视频名称</th><th>计划</th><th>AI</th><th>曝光量</th><th>投放消耗</th>'+
    '<th>3s完播率</th><th>5s完播率</th><th>完播率</th><th>投产比ROI</th>'+
    '<th>3日累计GMV</th><th>全站GMV</th></tr></thead><tbody>';
  vs.forEach(v=>{
    h+='<tr data-camp="'+(v.campaign||"").replace(/"/g,"&quot;")+'">'+
      '<td><span class="vname" title="'+v.name.replace(/"/g,"&quot;")+'">'+v.name+'</span></td>'+
      '<td style="text-align:left;color:#888">'+v.campaign+'</td>'+
      '<td>'+(v.ai?'<span class="badge b-add">AI</span>':'<span style="color:#CBD5E1">—</span>')+'</td>'+
      '<td>'+i(v.exp)+'</td><td>¥'+f(v.cost)+'</td><td>'+p(v.r3)+'</td><td>'+p(v.r5)+'</td>'+
      '<td>'+p(v.rend)+'</td><td class="'+(v.roi>=1?"hi":"lo")+'">'+v.roi.toFixed(2)+'</td>'+
      '<td>'+f(v.gmv3)+'</td><td class="hi">'+f(v.gmv)+'</td></tr>';
  });
  return h+'</tbody></table></div><div class="sub vcount" style="margin:8px 0 0"></div>';
}
// 每日趋势
const db=document.querySelector("#daily tbody");
DATA.daily.forEach(d=>{
  const names=Object.keys(d.per);
  names.forEach((n,li)=>{
    const m=d.per[n];
    const tr=document.createElement("tr");
    tr.innerHTML=(li===0?'<td rowspan="'+names.length+'">'+d.date+'</td>':'')+
      '<td>'+n+'</td><td>'+f(m.gmv)+'</td><td>'+f(m.gmv3)+'</td><td>¥'+f(m.cost)+'</td>'+
      '<td>'+i(m.videos)+'</td><td>'+i(m.exp)+'</td><td>'+m.roi.toFixed(2)+'</td>';
    db.appendChild(tr);
  });
});
</script></body></html>"""


def render(payload):
    return HTML.replace("__DATA__", json.dumps(payload, ensure_ascii=False))


if __name__ == "__main__":
    main()
