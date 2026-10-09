#!/usr/bin/env python3
"""按「品类 × 主播 × 素材标记」汇总视频投放数据。

用法:
    python3 summarize_by_marker.py --markers A,j,J
    python3 summarize_by_marker.py --markers j --in out/xxx/report.csv --out byJ.csv

标记判定用**包含**（photo_name 里出现该字母就算），不做位置判断。
原因：同一个账号体系下标记位置不统一——A 这类挂在整条标题最末尾
（`… #主持人泽鹏 A`），j/J 这类紧贴正文末尾、话题块之前
（`…告别体虚乏力j #花花严选周日上新 …`）。按位置判会整批漏掉。

包含判定会误召回（AIGC 的 A、10A 级、文件名里的 dj 等）。脚本会在命中里
挑出「标记不贴末尾」的标题打印出来，出数前自己看一眼。

注意一条标题可能同时含多个标记，且同一 photo_id 可能出现在两个品类，
所以「各行相加」会大于去重后的合计。对比总量时以去重合计为准。
"""
import argparse
import csv
import glob
import os
import re
import sys
from collections import defaultdict

# 账户名兜底映射：计划名里没写主播时用
ACCOUNT_HOST = {"四时颐养堂": "泽鹏", "花花好物清单": "花花"}
HOSTS = ("泽鹏", "花花", "小琦琦")
DATE_RE = re.compile(r"^\d+(\.\d+)?$")

# 参与汇总的指标：(字段名, 输出列名)
FIELDS = [("cost_total_yuan", "花费"), ("t3_gmv_yuan", "3日累计GMV"),
          ("custom_order_payment_amt_direct_yuan", "直接成交GMV")]


def parse_campaign(name):
    """'9.7-树德堂阿胶浆-泽鹏【目标投产4.8】' -> ('树德堂阿胶浆', '泽鹏')"""
    tokens = [t for t in re.sub(r"【[^】]*】", "", name).split("-") if t.strip()]
    if tokens and DATE_RE.match(tokens[0]):
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


def latest_report(out_root):
    hits = sorted(glob.glob(os.path.join(out_root, "*", "report.csv")))
    if not hits:
        sys.exit(f"没找到 report.csv，先跑 fetch_report.py。找的是 {out_root}/*/report.csv")
    return hits[-1]


def looks_like_marker(title, marker):
    """标记正常贴在正文末尾（话题块之前）或整条末尾。不贴的多半是误召回。"""
    body = re.split(r"\s+#", title.strip(), maxsplit=1)[0]
    return body.strip().endswith(marker) or title.strip().endswith(marker)


def aggregate(rows):
    """把明细行聚成 {品类, 主播, 视频数, 各指标}。"""
    groups = defaultdict(lambda: {"rows": 0, "photos": set(),
                                  **{k: 0.0 for _, k in FIELDS}})
    for r in rows:
        category, host = parse_campaign(r["campaign_name"])
        host = host or ACCOUNT_HOST.get(r["account_name"], "未知")
        g = groups[(category, host)]
        g["rows"] += 1
        g["photos"].add(r["photo_id"])
        for field, key in FIELDS:
            g[key] += float(r[field] or 0)
    return groups


def to_rows(marker, groups):
    out = []
    for (category, host), v in groups.items():
        cost = v["花费"]
        row = {"标记": marker, "品类": category, "主播": host,
               "视频数": len(v["photos"]), "数据行": v["rows"]}
        row.update({k: round(v[k], 2) for _, k in FIELDS})
        row["ROI_3日GMV除花费"] = round(v["3日累计GMV"] / cost, 3) if cost else 0
        out.append(row)
    return sorted(out, key=lambda x: -x["3日累计GMV"])


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    p = argparse.ArgumentParser()
    p.add_argument("--markers", default="A,j,J", help="逗号分隔，标题里包含即算")
    p.add_argument("--in", dest="src", help="report.csv 路径，默认取最新一次拉取")
    p.add_argument("--out", help="输出 CSV 路径")
    args = p.parse_args()

    markers = [m.strip() for m in args.markers.split(",") if m.strip()]
    src = args.src or latest_report(os.path.join(here, "out"))
    rows = list(csv.DictReader(open(src, encoding="utf-8-sig")))

    result, picked_union, suspicious = [], set(), defaultdict(set)
    for marker in markers:
        picked = [r for r in rows if marker in r["photo_name"]]
        picked_union.update(id(r) for r in picked)
        for r in picked:
            if not looks_like_marker(r["photo_name"], marker):
                suspicious[marker].add(r["photo_name"])
        result.extend(to_rows(marker, aggregate(picked)))
        # 自检：按标记分组的合计必须等于该标记筛选后的明细合计
        want = {k: sum(float(r[f] or 0) for r in picked) for f, k in FIELDS}
        got = {k: sum(o[k] for o in result if o["标记"] == marker) for _, k in FIELDS}
        for k in want:
            assert abs(want[k] - got[k]) < 0.05, \
                f"标记 {marker} 的 {k} 对不上: 明细 {want[k]:.2f} vs 分组 {got[k]:.2f}"

    union = [r for r in rows if id(r) in picked_union]
    total = {k: round(sum(float(r[f] or 0) for r in union), 2) for f, k in FIELDS}
    total["视频数"] = len({r["photo_id"] for r in union})
    total["数据行"] = len(union)

    # 文件名带上源数据的拉取时间戳，否则换个时间范围重跑会默默覆盖上一份
    pick = os.path.basename(os.path.dirname(src))
    dst = args.out or os.path.join(here, f"summary_marker_{''.join(markers)}_{pick}.csv")
    fields = ["标记", "品类", "主播", "视频数", "数据行",
              *(k for _, k in FIELDS), "ROI_3日GMV除花费"]
    with open(dst, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(result)

    print(f"标记 {markers} | 明细 {len(rows)} 行 | 输出 -> {dst}\n")
    print(f"{'标记':<5}{'品类':<12}{'主播':<6}{'视频数':>6}{'花费':>11}"
          f"{'3日累计GMV':>13}{'直接成交GMV':>13}{'ROI':>7}")
    for o in result:
        print(f"{o['标记']:<5}{o['品类']:<12}{o['主播']:<6}{o['视频数']:>6}{o['花费']:>11.2f}"
              f"{o['3日累计GMV']:>13.2f}{o['直接成交GMV']:>13.2f}{o['ROI_3日GMV除花费']:>7.2f}")
    print(f"\n去重合计（含多标记的标题只算一次）: 视频 {total['视频数']} 个 / 数据行 {total['数据行']}"
          f" / 花费 {total['花费']:.2f} / 3日累计GMV {total['3日累计GMV']:.2f}"
          f" / 直接成交GMV {total['直接成交GMV']:.2f}")
    dup = sum(o["数据行"] for o in result) - total["数据行"]
    if dup:
        print(f"标记行相加会多算 {dup} 行（同一标题含多个标记），对比时以去重合计为准。")
    for marker, titles in suspicious.items():
        if titles:
            print(f"\n标记 {marker} 命中 {len(titles)} 个「不贴末尾」的标题，疑似误召回，"
                  f"出数前确认一下：")
            for t in list(titles)[:5]:
                print(f"    {t[:80]}")


if __name__ == "__main__":
    main()
