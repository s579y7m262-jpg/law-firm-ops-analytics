# -*- coding: utf-8 -*-
"""
渲染器：把 build/ 下的 payload 注入 src/templates/ 对应模板，
产出 dashboards/ 下的单文件看板（零外部依赖，双击即开、可直接发人）。

用法
    python src/render.py                # 渲染全部
    python src/render.py report         # 只渲染某一个（名字见 PAGES 的 key）

设计说明
    模板里只留两个占位符：__CSS__（共享样式）与 __PAYLOAD__（数据 JSON）。
    样式集中在 base.css.html，五张看板视觉一致，改一处全部生效。
"""
import io
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TPL_DIR = os.path.join(ROOT, "src", "templates")
BUILD = os.path.join(ROOT, "build")
OUT_DIR = os.path.join(ROOT, "dashboards")

# key: (payload 文件名, 模板, 输出文件名)
PAGES = {
    "report":     ("payload.json",            "report.html",     "01_经营指标看板.html"),
    "monitor":    None,   # 由 src/monitor/monitor_etl.py 自行渲染（带热更新 data.js）
    "attribution":("payload_attribution.json","attribution.html","02_业绩下滑归因.html"),
    "censoring":  ("payload_censoring.json",  "censoring.html",  "03_成交率预测.html"),
    "scenarios":  ("payload_scenarios.json",  "scenarios.html",  "04_业绩预测三情景.html"),
    "payment":    ("payload_payment.json",    "payment.html",    "05_支付能力与客单价.html"),
}


def render_one(key):
    spec = PAGES.get(key)
    if not spec:
        print(f"[skip] {key}: 无独立渲染配置")
        return None
    payload_name, tpl_name, out_name = spec
    tpl_path = os.path.join(TPL_DIR, tpl_name)
    pay_path = os.path.join(BUILD, payload_name)
    if not os.path.exists(pay_path):
        print(f"[skip] {key}: 缺 {payload_name}（先跑对应的数据构建脚本）")
        return None

    html = io.open(tpl_path, encoding="utf-8").read()
    data = io.open(pay_path, encoding="utf-8").read()
    if "__PAYLOAD__" not in html:
        raise SystemExit(f"{tpl_name} 缺少 __PAYLOAD__ 占位符")

    if "__CSS__" in html:
        css = io.open(os.path.join(TPL_DIR, "base.css.html"), encoding="utf-8").read()
        html = html.replace("__CSS__", css)
    html = html.replace("__PAYLOAD__", data)

    os.makedirs(OUT_DIR, exist_ok=True)
    out_path = os.path.join(OUT_DIR, out_name)
    io.open(out_path, "w", encoding="utf-8").write(html)
    print(f"[ok] {out_name}  {os.path.getsize(out_path)/1024:.0f} KB")
    return out_path


def main():
    targets = sys.argv[1:] or [k for k, v in PAGES.items() if v]
    for k in targets:
        render_one(k)


if __name__ == "__main__":
    main()
