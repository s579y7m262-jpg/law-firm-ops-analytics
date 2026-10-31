# -*- coding: utf-8 -*-
"""
客户支付能力 × 客单价趋势
口径：
  - 时间：订单表「成交时间」2026.1–10
  - 去重：订单编号非空去重（一份合同=一单）
  - 付款能力四档：全款 / 首付≥90% / 首付50–90% / 首付<50%（分期客户按 首期认款÷合同总金额）
  - 客单价：合同额均价（主）、实付均价（并列）
输出：etl/out/payload_payment.json
"""
import pandas as pd, numpy as np, re, json, os, warnings
warnings.filterwarnings("ignore")

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA = os.path.join(ROOT, "data", "raw")
BUILD = os.path.join(ROOT, "build")
os.makedirs(BUILD, exist_ok=True)


def num(s):
    if isinstance(s, (int, float, np.integer, np.floating)):
        return float(s) if pd.notna(s) else np.nan
    t = re.sub(r"[^\d.\-]", "", str(s))
    try:
        return float(t) if t not in ("", "-", ".") else np.nan
    except Exception:
        return np.nan


# ── 1. 读取 ──
d1 = pd.read_excel(os.path.join(DATA, "订单表_2026.1-2026.9.xlsx"), sheet_name="订单导出明细")
d2 = pd.read_excel(os.path.join(DATA, "订单表_2026.10.xlsx"), sheet_name="订单导出明细")
raw_n = len(d1) + len(d2)
a = pd.concat([d1, d2], ignore_index=True)

# 订单编号非空去重（空值不可当同一 key）
dup = a["订单编号"].notna() & a.duplicated(subset=["订单编号"], keep="first")
dedup_n = int(dup.sum())
a = a[~dup].copy()

# ── 2. 时间窗口 ──
a["成交"] = pd.to_datetime(a["成交时间"], errors="coerce")
out_win = int((a["成交"].isna() | ~a["成交"].dt.year.eq(2026) | ~a["成交"].dt.month.between(1, 10)).sum())
a = a[a["成交"].dt.year.eq(2026) & a["成交"].dt.month.between(1, 10)].copy()
a["mon"] = a["成交"].dt.month.astype(int)

# ── 3. 金额解析 ──
a["合同"] = a["合同总金额"].map(num)
a["首期"] = a["首期认款金额"].map(num)
a["中期"] = a["中期款认款金额"].map(num)
a["尾款"] = a["尾款认款金额"].map(num)
a["增户"] = a["增户认款金额"].map(num)
a["支付"] = a["支付金额"].map(num)
a["周期"] = a["款项周期"].astype(str).str.strip().replace({"nan": np.nan, "None": np.nan})

# 首期占比（仅合同>0 且首期有值）
a["首付比"] = np.where((a["合同"] > 0) & a["首期"].notna(), a["首期"] / a["合同"], np.nan)

# ── 4. 付款能力四档 ──
def tier(r):
    if r["周期"] == "全款":
        return "全款"
    if r["周期"] == "分期":
        x = r["首付比"]
        if pd.isna(x):
            return "未标注"
        if x >= 0.9:
            return "首付≥90%"
        if x >= 0.5:
            return "首付50-90%"
        return "首付<50%"
    return "未标注"

a["档"] = a.apply(tier, axis=1)
TIERS = ["全款", "首付≥90%", "首付50-90%", "首付<50%", "未标注"]

# ── 5. 区域（用于明细，主视图不用） ──
def region(s):
    s = str(s or "")
    for r in ["A区", "B区", "C区", "D区"]:
        if s.startswith(r) or r in s:
            return r
    return "其他"
a["区域"] = a["归属部门"].map(region)

# 有效样本：合同额>0（算客单价用）
a["有效"] = a["合同"] > 0

# ── 6. 聚合 ──
def agg(df):
    n = len(df)
    nev = int(df["有效"].sum())
    c = df.loc[df["有效"], "合同"]
    p = df.loc[df["有效"], "支付"].fillna(0)
    return {
        "n": n, "n_valid": nev,
        "contract": float(c.sum()),
        "paid": float(df["支付"].fillna(0).sum()),
        "arpu_contract": float(c.sum() / nev) if nev else None,
        "arpu_paid": float(p.sum() / nev) if nev else None,
        "median_contract": float(c.median()) if nev else None,
        "down_ratio": float(df.loc[df["首付比"].notna(), "首付比"].median()) if df["首付比"].notna().any() else None,
        "down_ratio_mean": float(df.loc[df["首付比"].notna(), "首付比"].mean()) if df["首付比"].notna().any() else None,
        "install_share": float((df["周期"] == "分期").sum() / n * 100) if n else None,
    }

kpi = agg(a)

tier_rows = []
for t in TIERS:
    s = a[a["档"] == t]
    if len(s) == 0:
        continue
    g = agg(s)
    g["tier"] = t
    g["n_share"] = g["n"] / kpi["n"] * 100
    g["contract_share"] = g["contract"] / kpi["contract"] * 100 if kpi["contract"] else None
    tier_rows.append(g)

# 月度整体
month_rows = []
for m in range(1, 11):
    s = a[a["mon"] == m]
    g = agg(s)
    g["mon"] = m
    g["complete"] = m < 10
    month_rows.append(g)

# 月度 × 档位
mon_tier = []
for m in range(1, 11):
    for t in TIERS:
        s = a[(a["mon"] == m) & (a["档"] == t)]
        if len(s) == 0:
            continue
        g = agg(s)
        g["mon"] = m
        g["tier"] = t
        mon_tier.append(g)

# 合同额分档 × 分期率
BINS = [0, 3e4, 5e4, 1e5, 1.5e5, 2e5, 3e5, 1e12]
LABS = ["<3万", "3-5万", "5-10万", "10-15万", "15-20万", "20-30万", "30万+"]
a["合同档"] = pd.cut(a["合同"], bins=BINS, labels=LABS)
band_rows = []
for lb in LABS:
    s = a[a["合同档"] == lb]
    if len(s) == 0:
        continue
    n = len(s)
    inst = int((s["周期"] == "分期").sum())
    band_rows.append({
        "band": lb, "n": n, "install": inst,
        "install_rate": inst / n * 100,
        "contract": float(s["合同"].sum()),
        "arpu": float(s["合同"].mean()),
        "down_ratio": float(s.loc[s["首付比"].notna(), "首付比"].median()) if s["首付比"].notna().any() else None,
    })

# 档位 × 月份 热力：用客单价（合同）
heat = {}
for r in mon_tier:
    heat.setdefault(r["tier"], {})[r["mon"]] = r["arpu_contract"]

# ── 7. 数据质量 ──
dq = [
    {"level": "warn", "item": "10 月数据未完整",
     "detail": f"成交时间截至 {a['成交'].max().strftime('%Y-%m-%d')}，10 月仅 {int((a['mon']==10).sum())} 单",
     "impact": "10 月客单价波动不代表整月水平"},
    {"level": "warn", "item": "「全款」标记存在未付清",
     "detail": f"712 单标记全款，其中首期占合同额 <40% 的有 {int(((a['周期']=='全款')&(a['首付比']<0.4)).sum())} 单",
     "impact": "该标记按款项周期原值使用，未做二次修正；可能含增户或分期未回填"},
    {"level": "info", "item": "付款能力分档切点",
     "detail": "全款按「款项周期」；分期按 首期认款÷合同总金额：≥90% / 50–90% / <50%",
     "impact": f"分期 231 单中 <50% 占 {int(((a['周期']=='分期')&(a['首付比']<0.5)).sum())} 单，是主要分期群体"},
    {"level": "warn", "item": "合同额缺失 3 单",
     "detail": f"945 单中合同额>0 的 {kpi['n_valid']} 单",
     "impact": "客单价分母用有效样本，单数仍计入总量"},
    {"level": "info", "item": "客户星级不可用",
     "detail": "一星 1,286 / 五星 42 / 四星 15 / 三星 12 / 二星 9，一星占 94.5%",
     "impact": "视为默认值，未作为支付能力代理指标"},
    {"level": "info", "item": "回款流水表不可用",
     "detail": "「回款导出明细」工作表全部仅 2 行",
     "impact": "无法计算回款周期与分期履约率"},
]

payload = {
    "meta": {
        "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        "window": "2026.1–10",
        "last_date": a["成交"].max().strftime("%Y-%m-%d"),
        "tiers": TIERS,
        "bands": LABS,
        "raw_rows": raw_n, "dedup_dropped": dedup_n, "out_window": out_win,
    },
    "kpi": kpi,
    "tiers": tier_rows,
    "month": month_rows,
    "mon_tier": mon_tier,
    "band": band_rows,
    "heat": heat,
    "dq": dq,
}

with open(os.path.join(BUILD, "payload_payment.json"), "w", encoding="utf-8") as f:
    json.dump(payload, f, ensure_ascii=False, default=float)

# 明细导出（Excel 与核查用）
cols = ["订单编号", "案件编号", "案件名称", "合同甲方", "区域", "mon", "周期", "档",
        "合同", "首期", "中期", "尾款", "增户", "支付", "首付比", "成交"]
a[cols].to_csv(os.path.join(BUILD, "pay_detail.csv"), index=False, encoding="utf-8-sig")

print(f"原始 {raw_n} 行 → 去重删 {dedup_n} → 窗口外 {out_win} → 有效 {len(a)} 单")
print(f"合同额合计 {kpi['contract']:,.0f} | 实付 {kpi['paid']:,.0f}")
print(f"合同客单价 {kpi['arpu_contract']:,.0f} | 实付客单价 {kpi['arpu_paid']:,.0f}")
print(f"分期占比 {kpi['install_share']:.1f}% | 首付比中位 {kpi['down_ratio']:.2f}")
print("\n档位分布:")
for t in tier_rows:
    print(f"  {t['tier']:<10} n={t['n']:4d}  单数占比 {t['n_share']:5.1f}%  合同 {t['contract']:>12,.0f}  金额占比 {t['contract_share']:5.1f}%  客单价 {t['arpu_contract']:>9,.0f}")
print("\n合同档 × 分期率:")
for b in band_rows:
    print(f"  {b['band']:<7} n={b['n']:4d} 分期 {b['install']:3d} ({b['install_rate']:4.1f}%) 客单价 {b['arpu']:>9,.0f}")
