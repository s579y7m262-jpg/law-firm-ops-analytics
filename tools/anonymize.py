# -*- coding: utf-8 -*-
"""
脱敏处理：把区域名、小组前缀、机构名替换为代号，并在看板页面右下角打上「已脱敏」标记。

用法
    python tools/anonymize.py            # 处理 dashboards/ 与 src/ 下的代码与模板
    python tools/anonymize.py --check    # 只检查是否还有未脱敏的标识词

为什么单独成一个脚本
    作品集仓库里的成品看板来自真实业务数据。让脱敏成为一次可复现的、有映射表的批处理，
    而不是手工改文件，才能保证「代码里写的区域代号」和「页面上显示的区域代号」严格一致。
"""
import argparse
import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 区域代号：字母与规模无关，仅为去标识化
REGION_MAP = {
    "石家庄": "D区",
    "北京": "A区",
    "成都": "B区",
    "广州": "C区",
    "杭州": "E区",
    "深圳": "F区",
    "上海": "G区",
    "燕郊": "H区",
}

FIRM_MAP = {
    "翰汇所": "某律所",
    "翰汇": "某律所",
}

BADGE = """
<style>
#anon-badge{position:fixed;right:12px;bottom:12px;z-index:99999;font:500 11px/1.6 -apple-system,
"Segoe UI","Microsoft YaHei",sans-serif;color:#5B6572;background:rgba(255,255,255,.92);
border:1px solid #E6E8EC;border-radius:999px;padding:3px 10px;box-shadow:0 1px 6px rgba(0,0,0,.06)}
@media print{#anon-badge{display:none}}
</style>
<div id="anon-badge">示例数据 · 已脱敏</div>
"""

TARGET_DIRS = ["dashboards", "src", "docs"]


def apply_map(text):
    for k in sorted(REGION_MAP, key=len, reverse=True):
        text = text.replace(k, REGION_MAP[k])
    for k in sorted(FIRM_MAP, key=len, reverse=True):
        text = text.replace(k, FIRM_MAP[k])
    # 「A区区域」这类叠词收一下
    text = text.replace("区区域", "区域").replace("区区", "区")
    return text


def inject_badge(html):
    if "anon-badge" in html:
        return html
    if "</body>" in html:
        return html.replace("</body>", BADGE + "</body>", 1)
    return html + BADGE


def iter_files():
    for d in TARGET_DIRS:
        base = os.path.join(ROOT, d)
        for dp, _, fns in os.walk(base):
            for fn in fns:
                if fn.endswith((".html", ".js", ".py", ".md", ".css")):
                    yield os.path.join(dp, fn)


def check():
    pats = list(REGION_MAP) + list(FIRM_MAP)
    rx = re.compile("|".join(pats))
    bad = 0
    for p in iter_files():
        s = io.open(p, encoding="utf-8").read()
        hits = rx.findall(s)
        if hits:
            bad += 1
            print(f"[残留] {os.path.relpath(p, ROOT)}: {len(hits)} 处 {sorted(set(hits))}")
    print("检查通过，无未脱敏标识词" if not bad else f"共 {bad} 个文件仍含标识词")
    return 1 if bad else 0


def run():
    changed = 0
    for p in iter_files():
        s = io.open(p, encoding="utf-8").read()
        t = apply_map(s)
        if p.endswith(".html") and os.sep + "dashboards" + os.sep in p:
            t = inject_badge(t)
        if t != s:
            io.open(p, "w", encoding="utf-8").write(t)
            changed += 1
    print(f"已处理 {changed} 个文件")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args()
    sys.exit(check() if a.check else (run() or 0))
