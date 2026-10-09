"""dashboard.html 静态一致性检查：锚点、getElementById 目标、图表坐标函数。

纯文本检查，不需要运行服务器。跑：python test_dashboard_html.py
"""
import re
from pathlib import Path

HTML = Path(__file__).resolve().parent / "templates" / "dashboard.html"


def main():
    html = HTML.read_text(encoding="utf-8")
    ids = set(re.findall(r'\bid="([^"]+)"', html))
    problems = []

    # 1) 顶部跳转锚点必须指向存在的 id（否则点了不跳）
    for anchor in re.findall(r'href="#([^"]+)"', html):
        if anchor not in ids:
            problems.append(f"锚点 #{anchor} 没有对应的 id")

    # 2) JS 里 getElementById / querySelector('#x') 的目标必须存在
    refs = set(re.findall(r'getElementById\("([^"]+)"\)', html))
    refs |= set(re.findall(r'querySelector\("#([A-Za-z0-9_-]+)', html))
    for rid in refs:
        if rid not in ids:
            problems.append(f'JS 引用了不存在的 id: #{rid}')

    # 3) 图表坐标函数必须成对定义（曾经把 const Y 删掉过）
    for name in ("X", "Y"):
        if not re.search(rf'\bconst {name} = ', html):
            problems.append(f"图表缺少坐标函数 const {name}")

    # 4) 单日视图：X 必须居中，不能依赖 step=0
    if "dates.length > 1 ? pl + step * di : pl + iw / 2" not in html:
        problems.append("单日视图的 X 未居中")

    # 5) KPI 环比必须处理「前一日无数据」，不能显示成 ▲ 0%
    if "前一日无数据" not in html:
        problems.append("KPI 缺少「前一日无数据」分支")

    if problems:
        raise SystemExit("dashboard.html 检查失败:\n  - " + "\n  - ".join(problems))
    print(f"dashboard.html OK（{len(ids)} 个 id，{len(refs)} 个 JS id 引用）")


if __name__ == "__main__":
    main()
