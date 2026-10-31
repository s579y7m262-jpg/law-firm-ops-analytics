# -*- coding: utf-8 -*-
"""
2026-11 ~ 2027-03 业绩预测（漏斗驱动 + 三情景）

业绩口径：收入表「业绩金额」全部行求和（与既有看板一致，单位：元）
结构：业绩 = 新接认款 × (1 + 非新接系数)
      新接认款 = 合同额 × 认款率
      合同额   = 约见量(客/主) × 成交率(客/主) × 单均合同(客/主)

关键判断：成交率不直接用成熟月的 74%，而用「恢复系数 ρ」作用于成熟水平。
  10 月客场实录 30.2%，上一轮删失校正给出 47.1% / 56.6% / 64.3%。
  三情景的 ρ 直接锚在这三个校正值上：0.633 / 0.755 / 0.866。

约见外推：7–10 月线性拟合 + 阻尼（φ=0.75）+ 人为下限
季节性：2027-02 春节（2/6）打 0.85；2027-03 回补 1.05
"""
import os, json, re, warnings
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA = os.path.join(ROOT, "data", "raw")
BUILD = os.path.join(ROOT, "build")
os.makedirs(BUILD, exist_ok=True)

F_REV = os.path.join(DATA, "收入表_2026.1-2026.10.xlsx")
F_ORD = [os.path.join(DATA, "订单表_2026.1-2026.9.xlsx"),
         os.path.join(DATA, "订单表_2026.10.xlsx")]

CUT = pd.Timestamp("2026-10-29")
PHI = 0.75
FLOOR = {"客场": 35.0, "主场": 5.0}
SEASON = {1: 1.00, 2: 0.85, 3: 1.05}

# 上一轮删失校正的三个结果（10 月客场成交率）
OCT_HT = {"低": 47.1, "中": 56.6, "高": 64.3}
MATURE_K = 74.2          # 5–9 月客场成交率（加权），用于换算 ρ


def num(s):
    if isinstance(s, (int, float, np.integer, np.floating)):
        return float(s) if pd.notna(s) else np.nan
    t = re.sub(r"[^\d.\-]", "", str(s))
    try:
        return float(t) if t not in ("", "-", ".") else np.nan
    except ValueError:
        return np.nan


def pay_kind(t):
    t = str(t)
    if t.startswith("新接"):
        return "新接"
    if t.startswith("新增户"):
        return "新增户"
    if "中期" in t or "尾款" in t or "尾期" in t:
        return "中期尾款"
    if "退费" in t:
        return "退费"
    if "续" in t:
        return "续约"
    if "后" in t:
        return "后期效果费"
    return "其他"


# ══════════════════════════════════════════ 1. 收入表
xl = pd.ExcelFile(F_REV)
frames = []
for sh in xl.sheet_names:
    raw = pd.read_excel(xl, sheet_name=sh, header=None, usecols=range(0, 20))
    d = raw.iloc[2:].copy()
    d.columns = [str(c).strip() for c in raw.iloc[1].tolist()]
    d["_sheet"] = sh
    frames.append(d)
rev = pd.concat(frames, ignore_index=True)
rev = rev[rev["序号"].notna() | rev["案件编号"].notna()].copy()
rev["amt"] = pd.to_numeric(rev["业绩金额"], errors="coerce").fillna(0)
rev["mon"] = rev["_sheet"].str.replace("月", "").astype(int)
rev["kind"] = rev["新接/尾款"].fillna("").astype(str).map(pay_kind)

KINDS = ["新接", "新增户", "中期尾款", "其他", "后期效果费", "续约", "退费"]
rev_m = rev.pivot_table(index="mon", columns="kind", values="amt", aggfunc="sum").fillna(0)
for c in KINDS:
    if c not in rev_m.columns:
        rev_m[c] = 0.0
rev_m["业绩"] = rev_m[KINDS].sum(axis=1)
rev_m["非新接"] = rev_m["业绩"] - rev_m["新接"]

# ══════════════════════════════════════════ 2. 订单表漏斗
o1 = pd.read_excel(F_ORD[0], sheet_name="订单导出明细")
o2 = pd.read_excel(F_ORD[1], sheet_name="订单导出明细")
o = pd.concat([o1, o2], ignore_index=True)
dup = o["订单编号"].notna() & o.duplicated(subset=["订单编号"], keep="first")
o = o[~dup].copy()
o["约见"] = pd.to_datetime(o["约见时间"], errors="coerce")
o["成交"] = pd.to_datetime(o["成交时间"], errors="coerce")
o = o[o["约见"].notna()].copy()
o["m"] = o["约见"].dt.month
o["seg"] = np.where(o["约见类型"].astype(str).str.contains("主场"), "主场", "客场")
o["合同"] = o["合同总金额"].map(num).fillna(0)

rows = []
for m in range(1, 11):
    for s in ("客场", "主场"):
        x = o[(o["m"] == m) & (o["seg"] == s)]
        q = len(x)
        dd = int(x["成交"].notna().sum())
        amt = float(x.loc[x["成交"].notna(), "合同"].sum())
        rows.append({"mon": m, "seg": s, "meet": q, "deal": dd, "amt": amt,
                     "rate": dd / q if q else None, "avg": amt / dd if dd else None})
fun = pd.DataFrame(rows)

# ══════════════════════════════════════════ 3. 参数
rate_ref, avg_ref = {}, {}
for s in ("客场", "主场"):
    x = fun[(fun["seg"] == s) & (fun["mon"].between(5, 9))]
    rate_ref[s] = {"weighted": float(x["deal"].sum() / x["meet"].sum()),
                   "min": float(x["rate"].min()), "max": float(x["rate"].max())}
    y = fun[(fun["seg"] == s) & (fun["mon"].between(7, 10))]
    avg_ref[s] = {"weighted": float(y["amt"].sum() / y["deal"].sum())}

# 认款率
rec = pd.DataFrame([{"mon": m,
                     "contract": float(fun[fun["mon"] == m]["amt"].sum()),
                     "new_recog": float(rev_m.loc[m, "新接"])} for m in range(1, 11)])
rec["ratio"] = rec["new_recog"] / rec["contract"].replace(0, np.nan)
REC_RATE = float(rec[rec["mon"].between(3, 9)]["new_recog"].sum() /
                 rec[rec["mon"].between(3, 9)]["contract"].sum())

NONNEW = float(rev_m.loc[1:9, "非新接"].sum() / rev_m.loc[1:9, "新接"].sum())
NONNEW_RECENT = float(rev_m.loc[7:9, "非新接"].sum() / rev_m.loc[7:9, "新接"].sum())
FACTOR = REC_RATE * (1 + NONNEW)

# ══════════════════════════════════════════ 4. 约见外推
def fit_trend(seg):
    x = fun[(fun["seg"] == seg) & (fun["mon"].between(7, 10))][["mon", "meet"]].values
    k = np.polyfit(x[:, 0], x[:, 1], 1)
    return float(k[0])


trend = {s: fit_trend(s) for s in ("客场", "主场")}
base_q = {s: float(fun[(fun["seg"] == s) & (fun["mon"] == 10)]["meet"].iloc[0]) for s in ("客场", "主场")}


def proj_meet(seg, h):
    v = base_q[seg] + trend[seg] * sum(PHI ** i for i in range(1, h + 1))
    return max(v, FLOOR[seg])


# 情景：ρ = 成交率恢复系数
RHO = {"悲观": OCT_HT["低"] / MATURE_K,
       "中性": OCT_HT["中"] / MATURE_K,
       "乐观": OCT_HT["高"] / MATURE_K}
SCEN = {"悲观": {"q": 0.85, "avg": 0.92},
        "中性": {"q": 1.00, "avg": 1.00},
        "乐观": {"q": 1.20, "avg": 1.05}}

HORIZON = [(2026, 11), (2026, 12), (2027, 1), (2027, 2), (2027, 3)]

FUT = []
for h, (yy, mm) in enumerate(HORIZON, start=1):
    for sc, mul in SCEN.items():
        for seg in ("客场", "主场"):
            q = proj_meet(seg, h) * mul["q"]
            if yy == 2027:
                q *= SEASON.get(mm, 1.0)
            r = rate_ref[seg]["weighted"] * RHO[sc]
            a = avg_ref[seg]["weighted"] * mul["avg"]
            FUT.append({"ym": f"{yy}-{mm:02d}", "h": h, "scen": sc, "seg": seg,
                        "meet": q, "rate": r, "avg": a, "amt": q * r * a})
fut = pd.DataFrame(FUT)

mon_fc = []
for (ym, sc), g in fut.groupby(["ym", "scen"], sort=False):
    cont = float(g["amt"].sum())
    new_rec = cont * REC_RATE
    mon_fc.append({"ym": ym, "scen": sc,
                   "meet_k": float(g[g["seg"] == "客场"]["meet"].sum()),
                   "meet_h": float(g[g["seg"] == "主场"]["meet"].sum()),
                   "ramt_k": float(g[g["seg"] == "客场"]["amt"].sum()),
                   "ramt_h": float(g[g["seg"] == "主场"]["amt"].sum()),
                   "contract": cont, "new_recog": new_rec,
                   "nonnew": new_rec * NONNEW,
                   "revenue": new_rec * (1 + NONNEW)})
mfc = pd.DataFrame(mon_fc)

oct_recorded = float(rev_m.loc[10, "业绩"])
OCT_FULL = oct_recorded / 27 * 31

annual = []
for sc in ("悲观", "中性", "乐观"):
    q1_9 = float(rev_m.loc[1:9, "业绩"].sum())
    nov = float(mfc[(mfc["ym"] == "2026-11") & (mfc["scen"] == sc)]["revenue"].iloc[0])
    dec = float(mfc[(mfc["ym"] == "2026-12") & (mfc["scen"] == sc)]["revenue"].iloc[0])
    annual.append({"scen": sc, "q1_9": q1_9, "oct": OCT_FULL,
                   "nov": nov, "dec": dec, "total": q1_9 + OCT_FULL + nov + dec})

q1_2027 = {}
for sc in ("悲观", "中性", "乐观"):
    q1_2027[sc] = float(mfc[(mfc["scen"] == sc) & (mfc["ym"].str.startswith("2027"))]["revenue"].sum())

# ══════════════════════════════════════════ 5. 回测
bt = []
for origin in (6, 7):
    for seg in ("客场", "主场"):
        hist = fun[(fun["seg"] == seg) & (fun["mon"].between(origin - 3, origin))]
        k = np.polyfit(hist["mon"].values, hist["meet"].values, 1)[0]
        bq = float(fun[(fun["seg"] == seg) & (fun["mon"] == origin)]["meet"].iloc[0])
        for h in (1, 2):
            tgt = origin + h
            if tgt > 9:
                continue
            pred = max(bq + k * sum(PHI ** i for i in range(1, h + 1)), FLOOR[seg])
            act = float(fun[(fun["seg"] == seg) & (fun["mon"] == tgt)]["meet"].iloc[0])
            bt.append({"origin": origin, "seg": seg, "target": tgt,
                       "pred": round(pred, 1), "actual": act, "err": round(pred - act, 1)})

# ══════════════════════════════════════════ 6. 每场产出 / 业绩矩阵
RHO_N = RHO["中性"]
RK_N = rate_ref["客场"]["weighted"] * RHO_N
RH_N = rate_ref["主场"]["weighted"] * RHO_N
AK = avg_ref["客场"]["weighted"] / 1e4
AH = avg_ref["主场"]["weighted"] / 1e4
per_k = RK_N * AK
per_h = RH_N * AH
QH_N = 5.0


def revenue_of(qk, rk, qh=QH_N, rh=RH_N, ak=AK, ah=AH, factor=FACTOR):
    """给定约见量与成交率，反推业绩（万元）"""
    return (qk * rk * ak + qh * rh * ah) * factor


GRID_Q = [40, 50, 60, 70, 80, 90]
GRID_R = [0.45, 0.50, 0.55, 0.60, 0.65, 0.70, 0.75]
grid = []
for rk in GRID_R:
    grid.append({"rate": round(rk * 100, 0),
                 "vals": [round(revenue_of(q, rk), 0) for q in GRID_Q],
                 "hit600": next((q for q in GRID_Q if revenue_of(q, rk) >= 600), None)})

# 守 600 万所需的最小客场约见（各情景）
need600 = []
for sc in ("悲观", "中性", "乐观"):
    rk = rate_ref["客场"]["weighted"] * RHO[sc]
    rh = rate_ref["主场"]["weighted"] * RHO[sc]
    q = np.linspace(0, 200, 4001)
    v = (q * rk * AK + QH_N * rh * AH) * FACTOR
    need600.append({"scen": sc, "rate_k": round(rk * 100, 1),
                    "need_k": round(float(q[np.argmax(v >= 600)]), 0) if (v >= 600).any() else None})

# ══════════════════════════════════════════ 7. 敏感性（中性 11 月）
base_nov = float(mfc[(mfc["ym"] == "2026-11") & (mfc["scen"] == "中性")]["revenue"].iloc[0]) / 1e4
base_dec = float(mfc[(mfc["ym"] == "2026-12") & (mfc["scen"] == "中性")]["revenue"].iloc[0]) / 1e4
nk = float(fut[(fut.ym == "2026-11") & (fut.scen == "中性") & (fut.seg == "客场")]["meet"].iloc[0])
nh = float(fut[(fut.ym == "2026-11") & (fut.scen == "中性") & (fut.seg == "主场")]["meet"].iloc[0])
nk2 = float(fut[(fut.ym == "2026-12") & (fut.scen == "中性") & (fut.seg == "客场")]["meet"].iloc[0])
nh2 = float(fut[(fut.ym == "2026-12") & (fut.scen == "中性") & (fut.seg == "主场")]["meet"].iloc[0])


def with_shock(q, nh_, dq=0.0, dr=0.0, dp=0.0, nonnew=NONNEW):
    c = 0.0
    for qq, r0, a0 in [(q, RK_N, AK), (nh_, RH_N, AH)]:
        c += qq * (1 + dq) * (r0 + dr) * a0 * (1 + dp)
    return c * REC_RATE * (1 + nonnew)


sens = []
for name, kw in [
        ("约见量 −10%", dict(dq=-0.10)),
        ("约见量 +10%", dict(dq=0.10)),
        ("成交率 −5pp", dict(dr=-0.05)),
        ("成交率 +5pp", dict(dr=0.05)),
        ("单均合同 −10%", dict(dp=-0.10)),
        ("单均合同 +10%", dict(dp=0.10)),
        ("非新接系数 16.5%→11.5%", dict(nonnew=0.115)),
        ("非新接系数 16.5%→21.5%", dict(nonnew=0.215))]:
    v1 = with_shock(nk, nh, **kw)
    v2 = with_shock(nk2, nh2, **kw)
    sens.append({"item": name, "nov": round(v1, 1), "dec": round(v2, 1),
                 "delta": round(v1 - base_nov, 1)})

# ══════════════════════════════════════════ 8. payload
def W(v):
    return round(v / 1e4, 1)

payload = {
    "meta": {
        "cut": "2026-10-29",
        "window": "2026.1–10 实际 · 2026.11–2027.3 预测",
        "phi": PHI,
        "rec_rate": round(REC_RATE * 100, 1),
        "nonnew": round(NONNEW * 100, 1),
        "nonnew_recent": round(NONNEW_RECENT * 100, 1),
        "factor": round(FACTOR, 3),
        "oct_recorded": W(oct_recorded),
        "oct_full": W(OCT_FULL),
        "oct_days": 27,
        "oct_ht": OCT_HT,
        "mature_k": MATURE_K,
        "rho": {k: round(v * 100, 1) for k, v in RHO.items()},
        "floor": FLOOR,
    },
    "actual": [{
        "mon": int(m), "revenue": W(rev_m.loc[m, "业绩"]),
        "new_recog": W(rev_m.loc[m, "新接"]), "nonnew": W(rev_m.loc[m, "非新接"]),
        "meet": int(fun[fun["mon"] == m]["meet"].sum()),
        "meet_k": int(fun[(fun["mon"] == m) & (fun["seg"] == "客场")]["meet"].sum()),
        "meet_h": int(fun[(fun["mon"] == m) & (fun["seg"] == "主场")]["meet"].sum()),
        "deal": int(fun[fun["mon"] == m]["deal"].sum()),
        "contract": W(float(fun[fun["mon"] == m]["amt"].sum())),
        "rate": round(float(fun[fun["mon"] == m]["deal"].sum()) /
                      float(fun[fun["mon"] == m]["meet"].sum()) * 100, 1),
        "avg": round(float(fun[fun["mon"] == m]["amt"].sum()) /
                     max(float(fun[fun["mon"] == m]["deal"].sum()), 1) / 1e4, 2),
        "complete": m <= 9,
    } for m in range(1, 11)],
    "trend": {s: {"slope": round(trend[s], 2), "base": base_q[s],
                  "path": [round(proj_meet(s, h), 1) for h in range(1, 6)]}
              for s in ("客场", "主场")},
    "params": {
        "rate_k": round(rate_ref["客场"]["weighted"] * 100, 1),
        "rate_h": round(rate_ref["主场"]["weighted"] * 100, 1),
        "rate_k_range": [round(rate_ref["客场"]["min"] * 100, 1),
                         round(rate_ref["客场"]["max"] * 100, 1)],
        "avg_k": round(avg_ref["客场"]["weighted"] / 1e4, 2),
        "avg_h": round(avg_ref["主场"]["weighted"] / 1e4, 2),
        "per_k": round(per_k, 2), "per_h": round(per_h, 2),
        "rec_rate": round(REC_RATE * 100, 1),
        "nonnew": round(NONNEW * 100, 1),
        "factor": round(FACTOR, 3),
    },
    "fc": [{"ym": r.ym, "scen": r.scen, "meet_k": round(r.meet_k, 1), "meet_h": round(r.meet_h, 1),
            "contract": W(r.contract), "new_recog": W(r.new_recog),
            "nonnew": W(r.nonnew), "revenue": W(r.revenue)} for r in mfc.itertuples()],
    "annual": [{"scen": a["scen"], "q1_9": W(a["q1_9"]), "oct": W(a["oct"]),
                "nov": W(a["nov"]), "dec": W(a["dec"]), "total": W(a["total"]),
                "q1_2027": W(q1_2027[a["scen"]])} for a in annual],
    "backtest": bt,
    "rec_ratio": [{"mon": int(r.mon), "contract": W(r.contract),
                   "new_recog": W(r.new_recog),
                   "ratio": None if pd.isna(r.ratio) else round(r.ratio * 100, 1)}
                  for r in rec.itertuples()],
    "grid": {"q": GRID_Q, "rate": [round(r * 100, 0) for r in GRID_R], "rows": grid},
    "need600": need600,
    "per_neutral": {"per_k": round(per_k, 2), "per_h": round(per_h, 2),
                    "rk": round(RK_N * 100, 1), "rh": round(RH_N * 100, 1),
                    "q_h": QH_N},
    "sens": sens,
    "rev_mon": [{"mon": int(m), **{k: W(rev_m.loc[m, k]) for k in KINDS}} for m in range(1, 11)],
}

with open(os.path.join(BUILD, "payload_scenarios.json"), "w", encoding="utf-8") as f:
    json.dump(payload, f, ensure_ascii=False)

# ══════════════════════════════════════════ 9. 控制台
print("=" * 86)
print(f"认款率 {REC_RATE*100:.1f}% | 非新接 {NONNEW*100:.1f}%（近3月 {NONNEW_RECENT*100:.1f}%）| 综合乘数 {FACTOR:.3f}")
print(f"约见趋势 客场 {trend['客场']:+.2f}/月（10月 {base_q['客场']:.0f}）| 主场 {trend['主场']:+.2f}/月（10月 {base_q['主场']:.0f}）")
print(f"约见路径 客场 {[round(proj_meet('客场',h),1) for h in range(1,6)]}")
print(f"         主场 {[round(proj_meet('主场',h),1) for h in range(1,6)]}")
print(f"成熟成交率 客 {rate_ref['客场']['weighted']*100:.1f}% / 主 {rate_ref['主场']['weighted']*100:.1f}% | 恢复系数 {({k:round(v,3) for k,v in RHO.items()})}")
print(f"单均合同 客 {avg_ref['客场']['weighted']/1e4:.2f}万 / 主 {avg_ref['主场']['weighted']/1e4:.2f}万 | 每场产出 客 {per_k:.2f}万 / 主 {per_h:.2f}万")
print("-" * 86)
print(f"{'月份':<9}{'情景':<5}{'约见客':>7}{'约见主':>7}{'合同万':>9}{'新接万':>9}{'业绩万':>9}")
for r in mfc.itertuples(index=False):
    print(f"{r.ym:<9}{r.scen:<5}{r.meet_k:>7.1f}{r.meet_h:>7.1f}{r.contract/1e4:>9.1f}"
          f"{r.new_recog/1e4:>9.1f}{r.revenue/1e4:>9.1f}")
print("-" * 86)
print(f"{'情景':<5}{'1-9月':>9}{'10月':>8}{'11月':>8}{'12月':>8}{'全年':>9}{'2027Q1':>9}")
for a in payload["annual"]:
    print(f"{a['scen']:<5}{a['q1_9']:>9,.0f}{a['oct']:>8,.0f}{a['nov']:>8,.0f}"
          f"{a['dec']:>8,.0f}{a['total']:>9,.0f}{a['q1_2027']:>9,.0f}")
print("-" * 86)
print("回测（滚动原点，预测约见量）：")
for b in bt:
    print(f"  {b['origin']}月→{b['target']}月 {b['seg']}: 预测 {b['pred']:.0f} 实际 {b['actual']:.0f} 误差 {b['err']:+.0f}")
print("-" * 86)
print("中性条件下每场产出：客场 %.2f 万 / 主场 %.2f 万（成交率 客 %.1f%% 主 %.1f%%）" % (per_k, per_h, RK_N*100, RH_N*100))
print("守 600 万所需客场约见：", [(x["scen"], x["need_k"]) for x in need600])
print("敏感性（中性 11 月 / 12 月，万元）：")
for s_ in sens:
    print(f"  {s_['item']:<22} {s_['nov']:>8.1f} {s_['dec']:>8.1f}  Δ {s_['delta']:>+7.1f}")
print("=" * 86)
