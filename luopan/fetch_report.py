#!/usr/bin/env python3
"""快手磁力罗盘 - 自定义报表拉数。

用法:
    python3 fetch_report.py --check                          # 只验证 cookie 是否有效
    python3 fetch_report.py --begin 2026-09-27 --end 2026-09-29
    python3 fetch_report.py --begin 2026-09-01 --end 2026-09-30 --dimensions DAY,photo_id

cookie 放在同目录 cookie.txt，内容为浏览器 Cookie 头原文（一行，多行也能容忍）。
拉下来的原始 JSON 和拍平的 CSV 都写到 out/ 下，按时间戳分目录。
"""
import argparse
import csv
import json
import os
import ssl
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime, timedelta, timezone

BASE = "https://luopan.e.kuaishou.com/rest/web/v3/custom/analyse/report"
REFERER_TMPL = "https://luopan.e.kuaishou.com/analysis/custom-report-view?reportId={rid}"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36 Edg/154.0.0.0")
HERE = os.path.dirname(os.path.abspath(__file__))
COOKIE_FILE = os.path.join(HERE, "cookie.txt")

CST = timezone(timedelta(hours=8))  # 罗盘的日界按北京时间算，接口也是这个口径

# 服务端偶发限流/抖动的返回码，重试即可
RETRY_RESULTS = {8001, 8002, 8003}
RETRY_TIMES = 4


def load_cookie():
    """Cookie 只能用字符串传；curl -b 文件在本站会被判成未登录。

    用户可能把多行/多份 Cookie 一起粘进来，按 key 去重保留最后一个，避免重复键。
    """
    if not os.path.exists(COOKIE_FILE):
        sys.exit(f"缺少 cookie 文件: {COOKIE_FILE}\n"
                 "从浏览器 DevTools 把 Cookie 请求头整行粘进去即可。")
    raw = open(COOKIE_FILE, encoding="utf-8").read().replace("\r", " ").replace("\n", " ").strip()
    if not raw:
        sys.exit(f"cookie 文件是空的: {COOKIE_FILE}")
    jar = {}
    for part in raw.split(";"):
        part = part.strip()
        if "=" in part:
            k, v = part.split("=", 1)
            jar[k.strip()] = v.strip()
    return "; ".join(f"{k}={v}" for k, v in jar.items())


def call(path, payload, cookie, timeout=60):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(BASE + path, data=body, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("Accept", "application/json")
    req.add_header("Origin", "https://luopan.e.kuaishou.com")
    req.add_header("Referer", REFERER_TMPL.format(rid=payload.get("reportId", "")))
    req.add_header("User-Agent", UA)
    req.add_header("Cookie", cookie)
    ctx = ssl.create_default_context()

    last = ""
    for attempt in range(RETRY_TIMES):
        try:
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
                data = json.loads(r.read().decode("utf-8"))
        except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as e:
            last = f"{type(e).__name__}: {e}"
            data = None
        if data is not None:
            result = data.get("result")
            # 109 = 未登录，重试也没用
            if result == 109:
                sys.exit("cookie 已失效（result=109 未登录）。重新从浏览器复制 Cookie 覆盖 cookie.txt。\n"
                         f"登录地址: {data.get('loginUrl', '')}")
            if result not in RETRY_RESULTS and result == 1:
                return data["data"]
            if result not in RETRY_RESULTS:
                sys.exit(f"接口返回异常 result={result} message={data.get('message')}\n"
                         f"{json.dumps(data, ensure_ascii=False)[:400]}")
            last = f"result={result} {data.get('message')}"
        wait = 2 ** attempt
        print(f"  重试 {attempt + 1}/{RETRY_TIMES}（{last}），{wait}s 后再试", file=sys.stderr)
        time.sleep(wait)
    sys.exit(f"接口连续 {RETRY_TIMES} 次失败，最后一次: {last}\n"
             "多半是服务端限流，隔几分钟再跑，或把时间范围拆小。")


def ms(date_str, end=False):
    """日期字符串 -> 北京时间当天 00:00 或 23:59:59 的毫秒时间戳。"""
    d = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=CST)
    t = int(d.timestamp() * 1000)
    return t + 86_399_000 if end else t


def clamp_to_now(end_ms):
    """接口不允许结束时间超过当前时刻，今天当天要按现在截断。"""
    return min(end_ms, int(datetime.now(tz=CST).timestamp() * 1000))


def get_definition(rid, cookie):
    return call("/detail", {"reportId": str(rid)}, cookie)


def row_key(row, metrics):
    """业务的自然主键 = 一行里所有非指标字段。跨页重复的行靠它去掉。"""
    return tuple(sorted((k, v) for k, v in row.items() if k not in metrics))


def fetch_rows(rid, cookie, begin_ms, end_ms, group_type, dimensions, metrics,
               biz_line, data_theme, page_size=20000):
    """取数。

    **默认一次取全（pageSize 远大于行数）**，因为服务端分页不稳定：实测 pageSize
    1000/2000/5000 分别丢 439/177/113 行，去重救不回来，只有单页才和 summary 精确一致。
    行数超出 pageSize 时才真翻页，这时靠去重 + reconcile 严格对账兜底（丢行会被拦下）。
    ponytail: 接口没暴露稳定排序参数，所以翻页路径只能靠对账发现丢行，不能自动修。
    """
    seen, rows, page = set(), [], 1
    summary, columns, total, dup = None, None, None, 0
    while True:
        payload = {
            "reportId": int(rid),
            "timeCondition": {"beginTime": begin_ms, "endTime": end_ms, "groupType": group_type},
            "accountIds": [],
            "dimensions": dimensions,
            "metrics": metrics,
            "bizLine": biz_line,
            "dataTheme": data_theme,
            "filterConditions": [],
            "pageNum": page,
            "pageSize": page_size,
        }
        data = call("/detailinfo", payload, cookie)
        batch = data.get("detailList") or []
        for r in batch:
            k = row_key(r, set(metrics))
            if k in seen:
                dup += 1
                continue
            seen.add(k)
            rows.append(r)
        total = data.get("totalCount", total)
        summary = data.get("summary", summary)
        columns = data.get("columns", columns)
        print(f"  第 {page} 页: 累计 {len(rows)}/{total}（重复 {dup}）", file=sys.stderr)
        # 只在返回空页时收手。不能拿 totalCount 判断结束：万一服务端对 pageSize 有硬上限
        # 并静默截断，totalCount 可能一起被截断，按它判断会「一致地少行」而对账抓不到。
        # 多探一页的代价是一次请求，换掉一整类静默错误。
        if not batch:
            break
        if page > 200:  # 正常不会到；到这就是服务端没在翻页，别死循环
            print("  已经翻了 200 页还没取完，中止。检查 pageSize 或缩小时间范围。", file=sys.stderr)
            break
        page += 1
    if dup:
        print(f"  去掉 {dup} 行跨页重复", file=sys.stderr)
    if total is not None and len(rows) < total:
        print(f"  警告：接口说共 {total} 行，只取到 {len(rows)} 行，差 {total - len(rows)} 行。"
              f"调大 --page-size 重试。", file=sys.stderr)
    return rows, summary, columns, total


def is_additive(metric):
    """比率类指标不能求和，对账时跳过。"""
    return "ratio" not in metric and not metric.endswith("roi")


def reconcile(rows, summary, metrics, live=False):
    """和接口自带 summary 对账。分页丢行是静默的，只有这一步能发现。

    summary 是整个时间范围算好的权威值，明细逐行求和应该等于它。
    容差用绝对 0.05——**不要按百分比**：1.6 万的花费按 0.1% 会给出 34 元的容差，
    足以把丢行判成「一致」。
    窗口含当天时数据还在变，对不上只警告不报错。
    """
    if not summary:
        return
    bad, checked = [], 0
    for m in metrics:
        if m not in summary or not is_additive(m):
            continue
        checked += 1
        want = round(float(summary[m]), 2)
        got = round(sum(float(r.get(m) or 0) for r in rows), 2)
        if abs(want - got) > 0.05:
            bad.append(f"  {m}: 明细合计 {got} vs 接口 summary {want}（差 {round(got - want, 2)}）")
    print(f"  对账：{checked} 个可加指标与接口 summary {'一致' if not bad else '对不上'}", file=sys.stderr)
    if bad:
        print("分页结果与接口 summary 对不上，明细不可信：\n" + "\n".join(bad), file=sys.stderr)
        if live:
            print("注意：时间范围含当天，数据在实时变化，偏差可能来自此。", file=sys.stderr)
            return
        print("调大 --page-size 重跑（默认一次取全，不该出现这种偏差）。", file=sys.stderr)
        sys.exit(2)


def flatten(rows):
    """行是 dict，key 顺序在同一次响应里稳定；按首次出现顺序并集成表头。"""
    cols = []
    for r in rows:
        for k in r:
            if k not in cols:
                cols.append(k)
    return cols, [[r.get(c, "") for c in cols] for r in rows]


def main():
    now = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    p = argparse.ArgumentParser()
    p.add_argument("--report-id", default="699518973025")
    p.add_argument("--begin", help="YYYY-MM-DD，含当天")
    p.add_argument("--end", help="YYYY-MM-DD，含当天；默认今天")
    p.add_argument("--group-type", default="DAY")
    p.add_argument("--dimensions", help="逗号分隔；默认用报表定义里的全部维度")
    p.add_argument("--metrics", help="逗号分隔；默认用报表定义里的全部指标")
    p.add_argument("--page-size", type=int, default=20000, help="默认一次取全")
    p.add_argument("--out", default=os.path.join(HERE, "out"))
    p.add_argument("--check", action="store_true", help="只验证 cookie 有效性")
    args = p.parse_args()

    cookie = load_cookie()
    defn = get_definition(args.report_id, cookie)
    if args.check:
        print(f"cookie 有效。报表「{defn['reportName']}」"
              f" 维度 {len(defn['dimensions'])} 个 / 指标 {len(defn['metrics'])} 个")
        return

    if not args.begin:
        sys.exit("需要 --begin 起始日期，如 --begin 2026-09-27")
    dims = ([d.strip() for d in args.dimensions.split(",")] if args.dimensions
            else [d["dataIndex"] for d in defn["dimensions"] if d["dataIndex"] != "DAY"])
    mets = ([m.strip() for m in args.metrics.split(",")] if args.metrics
            else [m["dataIndex"] for m in defn["metrics"]])

    begin_ms = ms(args.begin)
    today = datetime.now(tz=CST).strftime("%Y-%m-%d")
    end_arg = args.end or today
    end_ms = clamp_to_now(ms(end_arg, end=True))
    live = end_arg >= today
    print(f"拉取 {args.begin} ~ {end_arg} | 维度 {dims} | 指标 {len(mets)} 个",
          file=sys.stderr)
    rows, summary, columns, total = fetch_rows(
        args.report_id, cookie, begin_ms, end_ms, args.group_type, dims, mets,
        defn["bizLine"], defn["dataTheme"], args.page_size)
    reconcile(rows, summary, mets, live=live)

    run_dir = os.path.join(args.out, now.replace(":", ""))
    os.makedirs(run_dir, exist_ok=True)
    raw = {"reportName": defn["reportName"], "generatedAt": now, "begin": args.begin,
           "end": args.end, "groupType": args.group_type, "dimensions": dims, "metrics": mets,
           "totalCount": total, "summary": summary, "columns": columns, "rows": rows}
    with open(os.path.join(run_dir, "raw.json"), "w", encoding="utf-8") as f:
        json.dump(raw, f, ensure_ascii=False, indent=1)

    cols, table = flatten(rows)
    csv_path = os.path.join(run_dir, "report.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(cols)
        w.writerows(table)
    print(f"共 {len(rows)} 行 / {len(cols)} 列 -> {csv_path}")


if __name__ == "__main__":
    main()
