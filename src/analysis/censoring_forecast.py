# -*- coding: utf-8 -*-
"""
10 月约见成交率预测模型
------------------------------------------------------------
问题：10 月原始成交率 30.2%，但 10 月约见最晚发生在 10/29，
      每场约见的可观察窗口长度不同（早期约见已看 27 天，月末约见只看了 0 天），
      直接算比率 = 把不同成熟度的样本混在一起，必然低估。

方法（三步）：
  1. 用 2–8 月（均已观察 ≥58 天，视为成熟）估计「成交延迟分布 F(t)」：
     F(t) = 在最终会成交的约见中，约见后 t 天内成交的占比。
  2. 对每个目标月做删失校正（逆概率加权，Horvitz–Thompson 型）：
     已成交的贡献 1；未成交的贡献 P(最终成交 | 尚未成交, 已观察 d 天)
     = p(1-F(d)) / (1 - p·F(d))，解一个不动点方程得到 p。
  3. 回测：把 6/7/8/9 月人为截断到「末期 + 10 天」，用同一套方法估计，
     与这些月的真实值比对，检验模型是否无偏。

输出：etl/out/payload_censoring.json
"""
import pandas as pd, numpy as np, json, os, warnings
warnings.filterwarnings("ignore")

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA = os.path.join(ROOT, "data", "raw")
BUILD = os.path.join(ROOT, "build")
os.makedirs(BUILD, exist_ok=True)

# ── 读数据 ──
d1 = pd.read_excel(os.path.join(DATA, "订单表_2026.1-2026.9.xlsx"), sheet_name="订单导出明细")
d2 = pd.read_excel(os.path.join(DATA, "订单表_2026.10.xlsx"), sheet_name="订单导出明细")
a = pd.concat([d1, d2], ignore_index=True)
a["约见"] = pd.to_datetime(a["约见时间"], errors="coerce")
a["成交"] = pd.to_datetime(a["成交时间"], errors="coerce")
a["支付"] = pd.to_datetime(a["支付时间（首笔）"], errors="coerce")
a = a[a["约见"].notna()].copy()

CUT = pd.Timestamp("2026-10-29 23:59:59")          # 数据观察截止
a["m"] = a["约见"].dt.month
a["客场"] = ~a["约见类型"].astype(str).str.contains("主场")
# 成交延迟：按日期差，当日成交=0；43 条真倒挂（补录）归 0
a["lag"] = (a["成交"].dt.normalize() - a["约见"].dt.normalize()).dt.days
a.loc[a["lag"].notna() & (a["lag"] < 0), "lag"] = 0
a["已观察"] = (CUT - a["约见"]).dt.total_seconds() / 86400.0

MATURE = [2, 3, 4, 5, 6, 7, 8]      # 已观察 ≥58 天，视为成熟
HORIZON = 30                        # 30 天外的成交极少（25/947），做上限截断


def build_F(df, horizon=HORIZON):
    """成交延迟 CDF：F[t] = 最终成交者中，t 天内成交的占比。t=0..horizon"""
    lag = df.loc[df["成交"].notna(), "lag"].dropna()
    lag = lag.clip(upper=horizon)
    n = len(lag)
    F = np.zeros(horizon + 1)
    for t in range(horizon + 1):
        F[t] = (lag <= t).sum() / n
    # 单调化 + 补到 1
    F = np.maximum.accumulate(F)
    return F, n


def ht_estimate(df, F, horizon=HORIZON, max_iter=200, lag_days=0.0):
    """删失校正后的最终成交率估计（不动点迭代）
    lag_days：录入滞后天数，相当于每场约见多看 lag_days 天"""
    d = df["已观察"].values - lag_days
    dealt = df["成交"].notna().values & (df["lag"].values <= horizon)
    # F(d)：按已观察天数取对应完成度，超出 horizon 取 1
    Fd = np.array([F[min(int(np.floor(x)), horizon)] if x >= 0 else 0.0 for x in d])
    Fd = np.clip(Fd, 1e-6, 1.0)

    p = dealt.mean() if dealt.mean() > 0 else 0.05
    for _ in range(max_iter):
        p = max(min(p, 0.999), 1e-4)
        # 未成交者的条件成交概率
        cond = p * (1 - Fd) / (1 - p * Fd)
        new_p = np.mean(dealt * 1.0 + (~dealt) * cond)
        if abs(new_p - p) < 1e-9:
            p = new_p
            break
        p = new_p
    return float(p)


def boot_ci(df, F, B=2000, horizon=HORIZON, seed=42):
    rng = np.random.default_rng(seed)
    n = len(df)
    vals = []
    for _ in range(B):
        idx = rng.integers(0, n, n)
        vals.append(ht_estimate(df.iloc[idx], F, horizon))
    vals = np.array(vals)
    return float(np.percentile(vals, 2.5)), float(np.percentile(vals, 97.5))


def raw_rate(df):
    return df["成交"].notna().mean()


def window_rate(df, w):
    s = df[df["已观察"] >= w]
    if len(s) == 0:
        return None, 0, 0
    hit = ((s["成交"].notna()) & (s["lag"] <= w)).sum()
    return hit / len(s), int(hit), len(s)


result = {"meta": {"cut": str(CUT), "horizon": HORIZON, "mature_months": MATURE}}

# ── 1. 延迟分布 F(t) ──
F_all, n_all = build_F(a[a["m"].isin(MATURE)])
F_k, n_k = build_F(a[a["m"].isin(MATURE) & a["客场"]])
F_h, n_h = build_F(a[a["m"].isin(MATURE) & ~a["客场"]])
result["F"] = {"all": [round(x, 4) for x in F_all],
               "away": [round(x, 4) for x in F_k],
               "home": [round(x, 4) for x in F_h],
               "n_all": n_all, "n_away": n_k, "n_home": n_h}

# ── 2. 逐月：原始率 / 窗长率 / 校正率 ──
month_rows = []
for m in range(1, 11):
    for seg, F, flag in [("客场", F_k, True), ("全部", F_all, None)]:
        df = a[(a["m"] == m)] if flag is None else a[(a["m"] == m) & (a["客场"] == flag)]
        if len(df) == 0:
            continue
        r7, h7, n7 = window_rate(df, 7)
        r14, h14, n14 = window_rate(df, 14)
        est = ht_estimate(df, F)
        row = {"月": m, "口径": seg, "约见": len(df), "原始成交": int(df["成交"].notna().sum()),
               "原始率%": round(raw_rate(df) * 100, 1),
               "7天窗可观测n": n7, "7天窗率%": None if r7 is None else round(r7 * 100, 1),
               "14天窗可观测n": n14, "14天窗率%": None if r14 is None else round(r14 * 100, 1),
               "校正后最终率%": round(est * 100, 1),
               "中位观察天": round(df["已观察"].median(), 0),
               "成熟": bool(m in MATURE or (m == 9 and df["已观察"].min() >= 28))}
        if not row["成熟"]:
            lo, hi = boot_ci(df, F)
            row["95%CI低"] = round(lo * 100, 1)
            row["95%CI高"] = round(hi * 100, 1)
        month_rows.append(row)
result["month"] = month_rows

# ── 3. 回测：把成熟月截断到「末期+10 天」，看估计能否还原真值 ──
back = []
for m in [6, 7, 8, 9]:
    df = a[(a["m"] == m) & a["客场"]].copy()
    if len(df) == 0:
        continue
    truth = raw_rate(df) * 100
    last_meet = df["约见"].max()
    fake_cut = last_meet + pd.Timedelta(days=10)
    sim = df.copy()
    sim["已观察"] = (fake_cut - sim["约见"]).dt.total_seconds() / 86400.0
    # 截断：假想观测日之后才成交的，视为「未成交」
    sim.loc[(sim["成交"].notna()) & (sim["成交"] > fake_cut), "成交"] = pd.NaT
    sim_raw = raw_rate(sim) * 100
    est = ht_estimate(sim, F_k) * 100
    back.append({"月": m, "真实最终率%": round(truth, 1), "模拟截断原始率%": round(sim_raw, 1),
                 "模型估计%": round(est, 1), "误差pp": round(est - truth, 1)})
result["backtest"] = back

# ── 4. 10 月：分周与分来源 ──
oct_k = a[(a["m"] == 10) & a["客场"]].copy()
oct_k["周序"] = ((oct_k["约见"].dt.day - 1) // 7 + 1)
wk = oct_k.groupby("周序").apply(lambda x: pd.Series({
    "约见": len(x), "成交": int(x["成交"].notna().sum()),
    "原始率%": round(x["成交"].notna().mean() * 100, 1),
    "中位观察天": round(x["已观察"].median(), 0)}), include_groups=False).reset_index()
result["oct_week"] = wk.to_dict("records")

oct_k["来源组"] = oct_k["来源"].astype(str).str.strip()
src = oct_k.groupby("来源组").apply(lambda x: pd.Series({
    "约见": len(x), "成交": int(x["成交"].notna().sum()),
    "原始率%": round(x["成交"].notna().mean() * 100, 1)}), include_groups=False)
src = src[src["约见"] >= 3].sort_values("约见", ascending=False).reset_index()
result["oct_src"] = src.to_dict("records")

# ── 5. 9 月 vs 10 月：可比子样本（都取每月 1–15 日的约见，14 天窗）──
cmp_rows = []
for m in [7, 8, 9, 10]:
    df = a[(a["m"] == m) & a["客场"] & (a["约见"].dt.day <= 15)]
    r14, h14, n14 = window_rate(df, 14)
    if n14 >= 8:
        cmp_rows.append({"月": m, "上半月约见": n14, "14天内成交": int(h14),
                         "14天窗率%": round(r14 * 100, 1)})
result["half_month"] = cmp_rows

# ── 6. 全部（含主场）10 月参考 ──
oct_all = a[a["m"] == 10]
result["oct_all"] = {"约见": len(oct_all), "成交": int(oct_all["成交"].notna().sum()),
                     "原始率%": round(raw_rate(oct_all) * 100, 1),
                     "校正后%": round(ht_estimate(oct_all, F_all) * 100, 1)}

# ── 6b. 非参数条件成交率 h(d)（成熟月实证，不依赖 F 形状）──
h_tab = {}
for d in [0, 1, 2, 3, 4, 5, 7, 10, 14, 18, 21, 28]:
    ss = a[a["m"].isin(MATURE) & a["客场"]].copy()
    ss = ss[(ss["lag"].isna()) | (ss["lag"] > d)]
    if len(ss) >= 10:
        h_tab[d] = (round(float(ss["成交"].notna().mean()), 4), int(len(ss)),
                    int(ss["成交"].notna().sum()))
result["h_tab"] = {str(k): v for k, v in h_tab.items()}

_ks = np.array(sorted(h_tab))
def _hd(d):
    k = _ks[_ks <= d]
    return h_tab[k[-1]][0] if len(k) else h_tab[_ks[0]][0]

_ok = a[(a["m"] == 10) & a["客场"]].copy()
_ok["d"] = np.floor(_ok["已观察"]).astype(int)
_ok["h"] = _ok["d"].map(_hd)
_ok["deal"] = _ok["成交"].notna()
np_est = float(np.where(_ok["deal"], 1.0, _ok["h"]).mean())
result["np_est"] = round(np_est * 100, 1)
result["np_week"] = [{"周": int(w),
                      "约见": int(len(x)),
                      "成交": int(x["deal"].sum()),
                      "d中位": float(x["d"].median()),
                      "h均值": round(float(x["h"].mean()), 3),
                      "估计%": round(float(np.where(x["deal"], 1.0, x["h"]).mean() * 100), 1)}
                     for w, x in _ok.groupby((( _ok["约见"].dt.day - 1) // 7 + 1))]

# ── 6c. 支付口径（定口参考：订单状态+首笔支付）──
PAY_STATES = {"打款到账审核", "增户到账审批", "成交", "律师分配", "中期款", "尾期款"}
a["paid"] = a["订单状态"].astype(str).isin(PAY_STATES) & a["支付"].notna()
pay_rows = []
for m in [8, 9, 10]:
    for seg, flt in [("客场", a["客场"]), ("全部", pd.Series(True, index=a.index))]:
        x = a[(a["m"] == m) & flt]
        if len(x) == 0:
            continue
        pay_rows.append({"月": m, "口径": seg, "约见": int(len(x)),
                         "成交时间口径": int(x["成交"].notna().sum()),
                         "成交时间率%": round(x["成交"].notna().mean() * 100, 1),
                         "支付口径": int(x["paid"].sum()),
                         "支付口径率%": round(x["paid"].mean() * 100, 1)})
result["pay_def"] = pay_rows

# ── 6d. 10 月订单状态分布（判断结果是否已产生）──
st = a[a["m"] == 10]["订单状态"].astype(str).value_counts().head(8)
result["oct_state"] = [{"状态": k, "笔数": int(v)} for k, v in st.items()]
prog = a[a["m"] == 10]["客户进展"].astype(str).value_counts().head(6)
result["oct_progress"] = [{"进展": k, "笔数": int(v)} for k, v in prog.items()]

# ── 7. 10 月：落在各级阈值以下的概率 ──
rng = np.random.default_rng(7)
oct_b = a[(a["m"] == 10) & a["客场"]]
n = len(oct_b)
vals = np.array([ht_estimate(oct_b.iloc[rng.integers(0, n, n)], F_k) * 100 for _ in range(3000)])
result["oct_prob"] = {"点估计%": round(ht_estimate(oct_b, F_k) * 100, 1),
                      "中位%": round(float(np.median(vals)), 1),
                      "P(<65%)": round(float((vals < 65).mean()) * 100, 1),
                      "P(<60%)": round(float((vals < 60).mean()) * 100, 1),
                      "P(<50%)": round(float((vals < 50).mean()) * 100, 1),
                      "分位": {q: round(float(np.percentile(vals, q)), 1) for q in (5, 10, 25, 50, 75, 90, 95)}}
result["oct_boot"] = [round(float(v), 1) for v in vals[:600]]
result_oct_boot = [round(float(v), 1) for v in vals]

# ── 模型混合：三种估计等权，各自带抽样波动 ──
_b = np.array(result_oct_boot, dtype=float)
_sd1 = (np.percentile(_b, 84) - np.percentile(_b, 16)) / 2
from math import erf, sqrt
_cdf = lambda z: 0.5 * (1 + erf(z / sqrt(2)))
_ms = [(ht_estimate(oct_b, F_k) * 100, _sd1),
       (ht_estimate(oct_b, F_k, lag_days=3) * 100, _sd1),
       (np_est * 100, 8.0)]
_ps = [_cdf((65 - m) / max(sd, 1e-6)) for m, sd in _ms]
result["mix"] = {"P_lt_65": round(float(np.mean(_ps)) * 100, 1),
                 "P_lt_60": round(float(np.mean([_cdf((60 - m) / max(sd, 1e-6)) for m, sd in _ms])) * 100, 1),
                 "中心": round(float(np.mean([m for m, _ in _ms])), 1),
                 "各模型P": [round(float(p_) * 100, 1) for p_ in _ps],
                 "sd": round(float(_sd1), 1)}

result["estimates"] = [
    {"模型": "删失校正（基准）", "估计%": round(ht_estimate(oct_b, F_k) * 100, 1),
     "低%": round(np.percentile(vals, 2.5), 1), "高%": round(np.percentile(vals, 97.5), 1),
     "假设": "2–8 月成交节奏不变；10 月记录已录完", "性质": "下限倾向"},
    {"模型": "删失校正 + 3 天录入滞后", "估计%": round(ht_estimate(oct_b, F_k, lag_days=3) * 100, 1),
     "低%": None, "高%": None,
     "假设": "同期约见多看 3 天", "性质": "中枢"},
    {"模型": "非参数条件率", "估计%": round(np_est * 100, 1),
     "低%": None, "高%": None,
     "假设": "10 月与成熟月同条件成交率，只是慢", "性质": "上限倾向"},
]

# ── 8. 录入滞后敏感性 ──
sens = []
for L in [0, 2, 3, 5, 7, 10, 14]:
    sens.append({"滞后天": L,
                 "客场估计%": round(ht_estimate(a[(a["m"] == 10) & a["客场"]], F_k, lag_days=L) * 100, 1),
                 "全口径估计%": round(ht_estimate(a[a["m"] == 10], F_all, lag_days=L) * 100, 1)})
result["lag_sens"] = sens

# ── 9. 10 月逐场明细（供人工追踪）──
od = a[(a["m"] == 10)].copy()
od["已成交"] = np.where(od["成交"].notna(), "是", "否")
od["F(已观察)"] = [round(F_k[min(int(np.floor(x)), HORIZON)] if c else
                         F_h[min(int(np.floor(x)), HORIZON)], 3)
                  for x, c in zip(od["已观察"], od["客场"])]
od = od.sort_values("约见")
result["oct_detail"] = [{
    "约见时间": str(r["约见"])[:16], "主场客场": "客场" if r["客场"] else "主场",
    "来源": str(r["来源"]), "成交时间": "" if pd.isna(r["成交"]) else str(r["成交"])[:10],
    "已成交": r["已成交"], "延迟天数": None if pd.isna(r["lag"]) else int(r["lag"]),
    "已观察天": int(r["已观察"]), "F(已观察)%": round(r["F(已观察)"] * 100, 0),
} for _, r in od.iterrows()]

json.dump(result, open(os.path.join(BUILD, "payload_censoring.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)

# ── 打印 ──
print("F(t) 客场，成熟月 2–8 月：")
for t in [0, 1, 3, 7, 14, 21, 30]:
    print(f"  t={t:2d} 天: {F_k[t]*100:5.1f}%")
print("\n逐月（客场）：")
for r in month_rows:
    if r["口径"] != "客场":
        continue
    ci = f"  CI[{r.get('95%CI低','—')}, {r.get('95%CI高','—')}]" if "95%CI低" in r else ""
    print(f"  {r['月']:2d}月 约见{r['约见']:3d} 原始{r['原始率%']:5.1f}% "
          f"7天窗{r['7天窗率%'] if r['7天窗率%'] is not None else '—':>5}% "
          f"校正{r['校正后最终率%']:5.1f}%{ci}  中位观察{r['中位观察天']:.0f}天")
print("\n回测：")
for b in back:
    print(f"  {b['月']}月 真实{b['真实最终率%']:5.1f}% | 截断后{b['模拟截断原始率%']:5.1f}% "
          f"| 模型{b['模型估计%']:5.1f}% | 误差{b['误差pp']:+.1f}pp")
print("\n10 月客场分周：")
for w in result["oct_week"]:
    print(f"  第{int(w['周序'])}周 约见{int(w['约见']):2d} 成交{int(w['成交']):2d} "
          f"原始{w['原始率%']:5.1f}% 中位观察{w['中位观察天']:.0f}天")
print("\n上半月可比（14 天窗）：")
for c in cmp_rows:
    print(f"  {c['月']}月 n={c['上半月约见']:2d} 14天内成交{c['14天内成交']:2d} → {c['14天窗率%']:.1f}%")
print("\n10 月全口径:", result["oct_all"])
print("\n10 月概率分布（客场，3000 次 bootstrap）:", result["oct_prob"])
print("\n模型混合:", result.get("mix"))
print("\n非参数估计:", result["np_est"], "%")
for w in result["np_week"]:
    print("  ", w)
print("\n支付口径 vs 成交时间口径：")
for r in result["pay_def"]:
    print(f"  {r['月']}月 {r['口径']}: 成交时间 {r['成交时间率%']}% | 支付口径 {r['支付口径率%']}%")
print("\n10 月订单状态:", result["oct_state"])
print("10 月客户进展:", result["oct_progress"])
print("\n录入滞后敏感性：")
for x in sens:
    print(f"  L={x['滞后天']:2d} 天 → 客场 {x['客场估计%']:5.1f}% | 全口径 {x['全口径估计%']:5.1f}%")
