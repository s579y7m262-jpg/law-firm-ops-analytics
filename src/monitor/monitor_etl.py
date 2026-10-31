# -*- coding: utf-8 -*-
"""
经营监测仪表盘 —— 自动化 ETL
================================================
用法：
    python etl/monitor_etl.py            # 正常跑（源文件没变则跳过）
    python etl/monitor_etl.py --force    # 强制重算

输入（目录内自动发现，新增月份文件丢进来即可，无需改代码）：
    订单表_2026*.xlsx                      约见量 / 成交 / 合同额
    收入表_2026*.xlsx                      业绩（「业绩金额」按月）
    给数据组结案，退费，返点汇总表*.xlsx    退费（取最新一份）

输出：
    2026经营监测仪表盘.html        自包含单文件（可直接发人）
    monitor_data.js     供页面 2 分钟轮询热更新
    etl/monitor_state.json         源文件指纹，用于跳过无变化的重算
    etl/monitor.log                追加日志

口径（与《业绩下滑归因与监测方案》《2026-2027 业绩预测》完全一致）：
    业绩   = 收入表「业绩金额」按 sheet 月份汇总
    约见   = 订单表有「约见时间」的记录数，按约见时间归月
             约见类型含"主场"→主场，其余→客场
    成交   = 「成交时间」非空
    合同额 = 「合同总金额」按约见月汇总
    退费率 = 退费金额 ÷ 当月业绩
"""
import os, sys, re, json, glob, hashlib, traceback
from datetime import datetime, timedelta

import warnings
warnings.filterwarnings("ignore")
import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(ROOT, "data", "raw")
BUILD = os.path.join(ROOT, "build")
os.makedirs(BUILD, exist_ok=True)
STATE = os.path.join(BUILD, "monitor_state.json")
LOG = os.path.join(BUILD, "monitor.log")
OUT_HTML = os.path.join(ROOT, "dashboards", "monitor.html")
OUT_JS = os.path.join(ROOT, "dashboards", "monitor_data.js")
TEMPLATE = os.path.join(HERE, "monitor_template.html")

TARGET_FLOOR = 600.0      # 底线目标（万元/月）
TARGET_IDEAL = 874.0      # 理想目标（万元/月，10,486÷12）
MEET_WARN = 70            # 客场约见预警线（场/月）
MEET_OK = 80              # 客场约见健康线
RATE_WARN = 65.0          # 成交率预警线（%）
RATE_OK_LO, RATE_OK_HI = 64.5, 80.0
REFUND_WARN = 8.0         # 退费率警戒（%）
REFUND_RED = 12.0         # 退费率红线（%）
PER_WARN = 7.0            # 客场单场产出预警（万元/场）


# ══════════════════════════════════════════════════════
# 0. 工具
# ══════════════════════════════════════════════════════
def log(msg):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line)
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def num(s):
    if isinstance(s, (int, float, np.integer, np.floating)):
        return float(s) if pd.notna(s) else np.nan
    t = re.sub(r"[^\d.\-]", "", str(s))
    try:
        return float(t) if t not in ("", "-", ".") else np.nan
    except Exception:
        return np.nan


def newest(pat):
    fs = glob.glob(os.path.join(DATA, pat))
    fs = [f for f in fs if not os.path.basename(f).startswith("~$")]
    return max(fs, key=os.path.getmtime) if fs else None


def fingerprint(files):
    h = hashlib.md5()
    for f in sorted(files):
        if not f:
            continue
        st = os.stat(f)
        h.update(f"{os.path.basename(f)}|{int(st.st_mtime)}|{st.st_size}".encode())
    return h.hexdigest()


# ══════════════════════════════════════════════════════
# 1. 发现源文件
# ══════════════════════════════════════════════════════
def discover():
    """自动发现源文件。年份从收入表文件名推断，跨年无需改代码。"""
    orders = [f for f in glob.glob(os.path.join(DATA, "订单表_20*.xlsx"))
              if not os.path.basename(f).startswith("~$")]
    orders.sort(key=lambda f: os.path.getmtime(f))
    revs = [f for f in glob.glob(os.path.join(DATA, "收入表_20*.xlsx"))
            if not os.path.basename(f).startswith("~$")]
    revs.sort(key=lambda f: os.path.getmtime(f))
    rev = revs[-1] if revs else None          # 收入表取最新一份（全量替换，不合并）
    year = None
    if rev:
        m = re.search(r"(20\d{2})", os.path.basename(rev))
        year = int(m.group(1)) if m else datetime.now().year
    # 订单表只保留同一年份的文件，避免混入往年全量导出
    if year:
        orders = [f for f in orders if f"_{year}" in os.path.basename(f)]
    ledger = newest("给数据组结案，退费，返点汇总表*.xlsx")
    return orders, rev, ledger, year


# ══════════════════════════════════════════════════════
# 2. 收入表 —— 业绩
# ══════════════════════════════════════════════════════
def load_revenue(path):
    xl = pd.ExcelFile(path)
    frames = []
    for sh in xl.sheet_names:
        m = re.search(r"(\d{1,2})\s*月", sh)
        if not m:
            continue
        raw = pd.read_excel(xl, sheet_name=sh, header=None, usecols=range(0, 20))
        if raw.shape[0] < 3:
            continue
        d = raw.iloc[2:].copy()
        d.columns = [str(c).strip() for c in raw.iloc[1].tolist()]
        d["_mon"] = int(m.group(1))
        frames.append(d)
    df = pd.concat(frames, ignore_index=True)
    keep = [c for c in ("序号", "案件编号") if c in df.columns]
    df = df[df[keep].notna().any(axis=1)].copy()
    df["amt"] = pd.to_numeric(df["业绩金额"], errors="coerce").fillna(0) \
        if "业绩金额" in df.columns else 0.0
    df["d"] = pd.to_datetime(df["时间"], errors="coerce") if "时间" in df.columns else pd.NaT
    df["tag"] = df["新接/尾款"].fillna("").astype(str).str.strip() if "新接/尾款" in df.columns else ""
    df["mon"] = df["_mon"]
    # 行日期与 sheet 月不符时，以 sheet 月为准（收入表按账期分表）
    df["day"] = df["d"].dt.day.fillna(0).astype(int)
    df.loc[df["d"].isna() | (df["d"].dt.year != 2026), "day"] = 0
    return df


# ══════════════════════════════════════════════════════
# 3. 订单表 —— 约见 / 成交 / 合同额
# ══════════════════════════════════════════════════════
def load_orders(paths, year=None):
    parts = []
    for f in paths:
        d = pd.read_excel(f, sheet_name="订单导出明细")
        d["_src"] = os.path.basename(f)
        parts.append(d)
    o = pd.concat(parts, ignore_index=True)
    n0 = len(o)
    if "订单编号" in o.columns:
        dup = o["订单编号"].notna() & o.duplicated(subset=["订单编号"], keep="first")
        o = o[~dup].copy()
    o["meet"] = pd.to_datetime(o["约见时间"], errors="coerce")
    o = o[o["meet"].notna()].copy()
    if year:
        o = o[o["meet"].dt.year == year].copy()
    o["deal"] = pd.to_datetime(o["成交时间"], errors="coerce").notna()
    o["ht"] = o["合同总金额"].map(num).fillna(0)
    o["kind"] = np.where(o["约见类型"].astype(str).str.contains("主场"), "主场", "客场")
    o["mon"] = o["meet"].dt.month
    o["day"] = o["meet"].dt.day
    o["ymd"] = o["meet"].dt.normalize()
    o["deal_dt"] = pd.to_datetime(o["成交时间"], errors="coerce")
    o["lag"] = (o["deal_dt"].dt.normalize() - o["ymd"]).dt.days
    o.loc[o["lag"].notna() & (o["lag"] < 0), "lag"] = 0
    return o, {"raw_rows": n0, "meet_rows": len(o)}


HORIZON = 30
REC_LAG = 3          # 录入滞后天数（与预测报告一致）


def build_F(df, horizon=HORIZON):
    """成交延迟 CDF：F[t] = 最终成交者中 t 天内成交的占比"""
    lag = df.loc[df["deal_dt"].notna(), "lag"].dropna().clip(upper=horizon)
    n = len(lag)
    if n < 20:
        return None, n
    F = np.array([(lag <= t).sum() / n for t in range(horizon + 1)])
    return np.maximum.accumulate(F), n


def ht_rate(df, F, cut, horizon=HORIZON, lag_days=REC_LAG, max_iter=200):
    """删失校正后的最终成交率（不动点迭代，口径同《10月约见成交率预测》）"""
    if F is None or len(df) < 8:
        return None
    d = ((cut - df["ymd"]).dt.total_seconds() / 86400.0).values - lag_days
    dealt = (df["deal_dt"].notna() & (df["lag"] <= horizon)).values
    Fd = np.array([F[min(int(np.floor(x)), horizon)] if x >= 0 else 0.0 for x in d])
    Fd = np.clip(Fd, 1e-6, 1.0)
    p = dealt.mean() if dealt.mean() > 0 else 0.05
    for _ in range(max_iter):
        p = max(min(p, 0.999), 1e-4)
        cond = p * (1 - Fd) / (1 - p * Fd)
        new_p = float(np.mean(dealt * 1.0 + (~dealt) * cond))
        if abs(new_p - p) < 1e-9:
            p = new_p
            break
        p = new_p
    return float(p)


# ══════════════════════════════════════════════════════
# 4. 退费
# ══════════════════════════════════════════════════════
def load_refund(path):
    from openpyxl import load_workbook
    if not path:
        return {}, []
    wb = load_workbook(path, data_only=True, read_only=True)
    rec = []
    for sn in wb.sheetnames:
        m = re.match(r"^(\d{2})(\d{2})$", str(sn).strip())
        if not m:
            continue
        ym = "20" + m.group(1) + "-" + m.group(2)
        ws = wb[sn]
        for i, row in enumerate(ws.iter_rows(min_row=2, max_col=12, values_only=True), start=2):
            v11 = row[10] if len(row) > 10 else None
            v12 = row[11] if len(row) > 11 else None
            if isinstance(v11, (int, float)) and v12 in ("是", "否"):
                rec.append({"ym": ym, "amt": float(v11), "sheet": sn, "row": i})
    wb.close()
    df = pd.DataFrame(rec)
    if df.empty:
        return {}, []
    df = df[df["ym"].str[:4] == "2026"]
    df["m"] = df["ym"].str[-2:].astype(int)
    g = df.groupby("m")["amt"].apply(lambda x: x.abs().sum() / 1e4)
    return g.to_dict(), df.to_dict("records")


# ══════════════════════════════════════════════════════
# 5. 进度曲线 ρ(N)：历史月份截止第 N 天的累计占全月比例
# ══════════════════════════════════════════════════════
def progress_curve(df, val_col, mon_col="mon", day_col="day", months=None, val_abs=False):
    """返回 {N: 中位比例}，仅用完整月份拟合"""
    out = {}
    sub = df[df[day_col] > 0]
    for m in months:
        s = sub[sub[mon_col] == m]
        tot = s[val_col].abs().sum() if val_abs else s[val_col].sum()
        if tot <= 0:
            continue
        cum = (s.groupby(day_col)[val_col].apply(lambda x: x.abs().sum() if val_abs else x.sum())
               .sort_index().cumsum() / tot)
        for d, v in cum.items():
            out.setdefault(int(d), []).append(float(v))
    return {d: float(np.median(v)) for d, v in out.items() if v}


def pick_rho(curve, n):
    """取第 n 天的比例；缺该天则取最近已知（向下）"""
    if n in curve:
        return curve[n]
    ks = [k for k in curve if k <= n]
    return curve[max(ks)] if ks else None


# ══════════════════════════════════════════════════════
# 主流程
# ══════════════════════════════════════════════════════
def build(force=False):
    orders_f, rev_f, ledger_f, YEAR = discover()
    if not rev_f:
        raise SystemExit("未找到 收入表_2026*.xlsx")
    if not orders_f:
        raise SystemExit("未找到 订单表_2026*.xlsx")

    fp = fingerprint(orders_f + [rev_f, ledger_f])
    if not force and os.path.exists(STATE):
        try:
            old = json.load(open(STATE, encoding="utf-8"))
        except Exception:
            old = {}
        if old.get("fp") == fp and os.path.exists(OUT_HTML):
            log(f"源文件无变化（fp={fp[:8]}），跳过重算。如需强制：--force")
            return False

    t0 = datetime.now()
    log("=" * 62)
    log(f"开始重算  fp={fp[:8]}")

    rv = load_revenue(rev_f)
    od, ometa = load_orders(orders_f, YEAR)
    refund_map, refund_rec = load_refund(ledger_f)

    # ── 数据截止日 ──
    last_rev = rv["d"].max()
    last_meet = od["meet"].max()
    cut = max([x for x in (last_rev, last_meet) if pd.notna(x)])
    cut = pd.Timestamp(cut).normalize()
    log(f"数据截止 {cut.date()}（收入表 {last_rev.date() if pd.notna(last_rev) else '—'}，"
        f"订单约见 {last_meet.date() if pd.notna(last_meet) else '—'}）")

    cur_mon = cut.month
    mon_days = pd.Timestamp(cut.year, cut.month, 1).days_in_month
    day_n = cut.day

    # ── 月度聚合 ──
    rmg = rv.groupby("mon")["amt"].sum() / 1e4
    og = od.groupby(["mon", "kind"]).agg(
        meet=("订单编号", "size") if "订单编号" in od.columns else ("meet", "size"),
        deal=("deal", "sum"),
        ht=("ht", lambda x: x.sum() / 1e4)).reset_index()

    months = sorted(set(rmg.index) | set(od["mon"].unique()))
    months = [m for m in months if 1 <= m <= 12]
    month = []
    for m in months:
        h = og[(og["mon"] == m) & (og["kind"] == "主场")]
        k = og[(og["mon"] == m) & (og["kind"] == "客场")]
        hv, kv = float(h["meet"].sum()), float(k["meet"].sum())
        hd, kd = float(h["deal"].sum()), float(k["deal"].sum())
        hc, kc = float(h["ht"].sum()), float(k["ht"].sum())
        rev = float(rmg.get(m, 0.0))
        mt, dl, ct = hv + kv, hd + kd, hc + kc
        ref = refund_map.get(m)
        month.append({
            "mon": m, "rev": round(rev, 1),
            "meet": int(mt), "meet_h": int(hv), "meet_k": int(kv),
            "deal": int(dl), "deal_h": int(hd), "deal_k": int(kd),
            "cont": round(ct, 1), "cont_h": round(hc, 1), "cont_k": round(kc, 1),
            "rate": round(dl / mt * 100, 1) if mt else None,
            "rate_k": round(kd / kv * 100, 1) if kv else None,
            "per_k": round(kc / kv, 2) if kv else None,
            "per_h": round(hc / hv, 2) if hv else None,
            "arpu": round(ct / dl, 1) if dl else None,
            "refund": round(float(ref), 1) if ref is not None else None,
            "refund_rate": round(float(ref) / rev * 100, 1) if (ref is not None and rev > 0) else None,
            "complete": m < cur_mon or (m == cur_mon and cut.day >= mon_days),
        })
    cur = next((x for x in month if x["mon"] == cur_mon), None)
    prev = next((x for x in month if x["mon"] == cur_mon - 1), None)

    # ── 删失校正成交率：当月样本未成熟，原始比率必然低估 ──
    mature_mons = [m for m in months if 2 <= m <= cur_mon - 2]
    F_all, nF = build_F(od[od["mon"].isin(mature_mons)])
    F_k, nFk = build_F(od[(od["mon"].isin(mature_mons)) & (od["kind"] == "客场")])
    for x in month:
        if x["mon"] >= cur_mon - 1 and not x["complete"]:
            s = od[od["mon"] == x["mon"]]
            x["rate_adj"] = round(ht_rate(s, F_all, cut) * 100, 1) if len(s) else None
            sk = s[s["kind"] == "客场"]
            x["rate_k_adj"] = round(ht_rate(sk, F_k, cut) * 100, 1) if len(sk) else None
        else:
            x["rate_adj"] = x["rate"]
            x["rate_k_adj"] = x["rate_k"]

    # ── 进度曲线（完整月份拟合）──
    full_months = [x["mon"] for x in month if x["complete"] and x["mon"] >= 2]
    rho_rev = progress_curve(rv, "amt", months=full_months)
    rho_meet = progress_curve(od.assign(_one=1.0), "_one", months=full_months)

    r_rev = pick_rho(rho_rev, day_n)
    r_meet = pick_rho(rho_meet, day_n)
    rev_mtd = cur["rev"] if cur else 0.0
    meet_k_mtd = cur["meet_k"] if cur else 0
    meet_mtd = cur["meet"] if cur else 0
    deal_mtd = cur["deal"] if cur else 0

    # 折算（进度曲线法；进度太早不外推）
    def proj(v, rho, floor_rho=0.55):
        if rho and rho >= floor_rho and v > 0:
            return v / rho
        return None
    rev_proj = proj(rev_mtd, r_rev)
    meet_k_proj = proj(meet_k_mtd, r_meet)
    meet_proj = proj(meet_mtd, r_meet)
    rev_linear = rev_mtd / max(day_n, 1) * mon_days
    meet_k_linear = meet_k_mtd / max(day_n, 1) * mon_days

    # ── 剩余时间 ──
    month_end = pd.Timestamp(cut.year, cut.month, mon_days)
    left_days = (month_end - cut).days
    left_biz = int(sum(1 for i in range(1, left_days + 1)
                       if (cut + timedelta(days=i)).weekday() < 5))
    time_pct = day_n / mon_days * 100

    # ── 周粒度（近 14 周，客场约见）──
    w = od[od["kind"] == "客场"].copy()
    w["wk"] = w["ymd"] - pd.to_timedelta(w["ymd"].dt.weekday, unit="D")   # 周一
    wg = w.groupby("wk").agg(meet=("ymd", "size"), deal=("deal", "sum")).reset_index()
    wg = wg.sort_values("wk")
    wg = wg[wg["wk"] <= cut]
    weeks = []
    for _, r in wg.tail(14).iterrows():
        weeks.append({"wk": r["wk"].strftime("%m/%d"), "iso": r["wk"].strftime("%Y-%m-%d"),
                      "meet": int(r["meet"]), "deal": int(r["deal"]),
                      "rate": round(r["deal"] / r["meet"] * 100, 1) if r["meet"] else None})
    # 当前周未走完，标记
    if weeks and weeks[-1]["iso"] == (cut - timedelta(days=cut.weekday())).strftime("%Y-%m-%d"):
        weeks[-1]["partial"] = True

    wm = wg["meet"]
    # 当前周未走完，均值只用完整周计算，避免把半周拉低
    wm_full = wm.iloc[:-1] if (weeks and weeks[-1].get("partial")) else wm
    wk_last4 = float(wm_full.tail(4).mean()) if len(wm_full) >= 4 else None
    wk_prev4 = float(wm_full.tail(8).head(4).mean()) if len(wm_full) >= 8 else None
    wk_run = float(wm_full.iloc[-1]) if len(wm_full) >= 1 else None
    wk_cur = weeks[-1] if (weeks and weeks[-1].get("partial")) else None

    # ── 退费 ──
    ref_mtd = cur["refund"] if cur else None

    # ── 预测对照 ──
    fc_path = os.path.join(BUILD, "payload_scenarios.json")
    fc_rows, annual, fc_meta, need600 = [], None, {}, []
    if os.path.exists(fc_path):
        p = json.load(open(fc_path, encoding="utf-8"))
        fc_meta = p.get("meta", {})
        for d in (p.get("fc") or []):
            if isinstance(d, dict) and str(d.get("ym", "")) >= "2026-11":
                fc_rows.append({"ym": d["ym"], "scen": d["scen"],
                                "rev": round(float(d.get("revenue", 0)), 1),
                                "meet_k": d.get("meet_k")})
        annual = [a for a in (p.get("annual") or []) if isinstance(a, dict)]
        need600 = [d for d in (p.get("need600") or []) if isinstance(d, dict)]

    fc_neutral = {}
    for r in fc_rows:
        if r["scen"] == "中性":
            fc_neutral[r["ym"]] = r["rev"]
    annual_neutral = next((a for a in annual if a.get("scen") == "中性"), None) if annual else None
    annual_floor = next((a for a in annual if a.get("scen") == "悲观"), None) if annual else None

    # ── 累计 ──
    ytd_rev = round(sum(x["rev"] for x in month), 1)
    ytd_meet = sum(x["meet"] for x in month)
    ytd_meet_k = sum(x["meet_k"] for x in month)
    ytd_deal = sum(x["deal"] for x in month)
    # 年内时间进度（用于和业绩进度比）
    yday = (cut - pd.Timestamp(cut.year, 1, 1)).days + 1
    ydays = 366 if pd.Timestamp(cut.year, 1, 1).is_leap_year else 365
    ytime_pct = yday / ydays * 100

    # 最近一个完整月（用于滞后指标判定）
    last_full = next((x for x in reversed(month) if x["complete"]), None)

    # ── 信号灯 ──
    def st(level, val, thr, act):
        return {"level": level, "val": val, "thr": thr, "act": act}

    sig = []
    # 主指标：客场约见（用折算全月）
    mk = meet_k_proj if meet_k_proj else meet_k_mtd
    if mk < MEET_WARN * 0.9:
        s = st("red", round(mk, 0), f"<{MEET_WARN*0.9:.0f} 场", "客场约见严重不足，需立刻加投放/加排期")
    elif mk < MEET_WARN:
        s = st("amber", round(mk, 0), f"<{MEET_WARN} 场", "贴着预警线，本周需补约见")
    elif mk < MEET_OK:
        s = st("amber", round(mk, 0), f"<{MEET_OK} 场", "未达健康线，维持投放强度")
    else:
        s = st("green", round(mk, 0), f"≥{MEET_OK} 场", "达标，保持")
    sig.append({"key": "meet_k", "layer": "主指标", "name": "客场约见量",
                "note": "本月折算，场", **s})

    # 护栏1：成交率（用删失校正值判定；原始值当月必然低估，仅作展示）
    rt_raw = cur["rate"] if cur else None
    rt = cur["rate_adj"] if cur else None
    if rt is None:
        s = st("gray", None, f"<{RATE_WARN}%", "样本不足，暂不判定")
    elif meet_mtd < 25:
        s = st("gray", rt, f"<{RATE_WARN}%", "本月约见样本 <25 场，校正值不稳，累计后再判")
    elif rt < RATE_WARN - 8:
        s = st("red", rt, f"<{RATE_WARN}%", "质的问题：抽查约见录音与主谈配置，非加大投放能解")
    elif rt < RATE_WARN:
        s = st("amber", rt, f"<{RATE_WARN}%", "跌破预警线，核对客场来源质量与主谈人")
    else:
        s = st("green", rt, f"≥{RATE_WARN}%", "正常")
    sig.append({"key": "rate", "layer": "护栏 1", "name": "约见成交率（校正）",
                "note": "本月，%", "raw": rt_raw, **s})

    # 护栏2：退费率（台账可能晚于收入表，需区分「本月没退费」和「台账没到本月」）
    rr = cur["refund_rate"] if cur else None
    ledger_months = sorted(refund_map.keys())
    covered = cur_mon in ledger_months
    if not covered:
        s = st("gray", None, f">{REFUND_RED}%",
               f"退费台账只到 {ledger_months[-1] if ledger_months else '—'} 月，本月未覆盖；"
               f"台账更新后需复核")
    elif rr is None:
        s = st("green", 0.0, f"≤{REFUND_WARN}%", "本月无退费记录")
    elif rr > REFUND_RED:
        s = st("red", rr, f">{REFUND_RED}%", "退费抬头，前序量价结论需重算")
    elif rr > REFUND_WARN:
        s = st("amber", rr, f">{REFUND_WARN}%", "关注退费明细，按案件排查")
    else:
        s = st("green", rr, f"≤{REFUND_WARN}%", "正常")
    sig.append({"key": "refund", "layer": "护栏 2", "name": "退费率",
                "note": "本月，%", "ledger_to": ledger_months[-1] if ledger_months else None, **s})

    # 参考：客场单场产出（滞后指标，当月分子还没长齐，判定用最近完整月）
    pk_now = cur.get("per_k") if cur else None
    pk = last_full.get("per_k") if last_full else None
    if pk is None:
        s = st("gray", None, f"<{PER_WARN} 万", "无完整月样本")
    elif pk < PER_WARN:
        s = st("amber", pk, f"<{PER_WARN} 万", "价在补量但补不动了，需查客单价结构")
    else:
        s = st("green", pk, f"≥{PER_WARN} 万", "正常")
    sig.append({"key": "per_k", "layer": "参考", "name": "客场单场产出",
                "note": f"{last_full['mon'] if last_full else '—'} 月完整值，万元/场",
                "now": pk_now, **s})

    # 达成率：预计全月 vs 时间进度
    ach_mtd = rev_mtd / TARGET_FLOOR * 100
    ach_proj = (rev_proj / TARGET_FLOOR * 100) if rev_proj else None
    ach_ideal = (rev_proj / TARGET_IDEAL * 100) if rev_proj else (rev_mtd / TARGET_IDEAL * 100)
    ach = ach_proj if ach_proj is not None else ach_mtd
    gap = ach - time_pct
    if gap < -12:
        s = st("red", round(ach, 1), f"落后时间进度 >12pt", "本月守不住 600 万，需重排下月目标")
    elif gap < -4:
        s = st("amber", round(ach, 1), f"落后时间进度 >4pt", "后半程需补确认")
    else:
        s = st("green", round(ach, 1), "与时间进度同步", "节奏正常")
    sig.append({"key": "ach", "layer": "业绩", "name": "月目标达成率(600万)",
                "note": "预计全月/600万，%", "now": round(ach_mtd, 1), **s})

    # 全年：YTD 进度 vs 年内时间进度
    ann_target = float(annual_neutral["total"]) if annual_neutral else None
    ann_pct = ytd_rev / ann_target * 100 if ann_target else None
    if ann_pct is None:
        s = st("gray", None, "—", "无预测基线")
    else:
        gy = ann_pct - ytime_pct
        if gy < -5:
            s = st("red", round(ann_pct, 1), f"落后年内进度 >5pt", "全年会落到悲观情景")
        elif gy < 0:
            s = st("amber", round(ann_pct, 1), "落后年内进度 <5pt", "全年略低于中性")
        else:
            s = st("green", round(ann_pct, 1), "不低于年内进度", "全年可达中性")
    sig.append({"key": "annual", "layer": "全年", "name": "全年进度(中性基线)",
                "note": "YTD/中性预测，%", "ytime": round(ytime_pct, 1), **s})

    reds = [x for x in sig if x["level"] == "red"]
    ambers = [x for x in sig if x["level"] == "amber"]
    overall = "red" if reds else ("amber" if ambers else "green")

    # ── 摘要 ──
    summary = []
    if cur:
        head = (f"本月（{cur_mon} 月）截至 {cut.month}/{cut.day} 入账 "
                f"<b>{rev_mtd:,.0f} 万</b>")
        if rev_proj:
            head += (f"，按历史同进度折算全月约 <b>{rev_proj:,.0f} 万</b>"
                     f"（{'站上' if rev_proj >= TARGET_FLOOR else '未达'} 600 万底线，"
                     f"{'站上' if rev_proj >= TARGET_IDEAL else '未达'} 874 万理想线）")
        else:
            head += "（月初样本少，暂不折算全月）"
        head += f"；时间已过 <b>{time_pct:.0f}%</b>，业绩已入账进度 <b>{ach_mtd:.0f}%</b>。"
        summary.append(head)
    if cur:
        summary.append(
            f"客场约见本月 <b>{meet_k_mtd} 场</b>"
            + (f"，折算全月 <b>{meet_k_proj:.0f} 场</b>" if meet_k_proj else "")
            + f"，{'高于' if (meet_k_proj or meet_k_mtd) >= MEET_WARN else '低于'} {MEET_WARN} 场预警线；"
            f"成交率删失校正后 <b>{rt if rt is not None else '—'}%</b>"
            f"（原始 {rt_raw if rt_raw is not None else '—'}%，未成熟），"
            f"近 4 个完整周客场周均 <b>{wk_last4:.0f} 场</b>"
            + (f"（前 4 周 {wk_prev4:.0f} 场）" if wk_prev4 else "") + "。")
    if reds:
        summary.append("触红：" + "、".join(f"<b>{x['name']}</b>" for x in reds) + "。按下方动作清单先处理红项。")
    elif ambers:
        summary.append("无红灯，黄灯项：" + "、".join(f"<b>{x['name']}</b>" for x in ambers) + "，需持续盯。")
    else:
        summary.append("全部指标在健康区间，维持当前投放与排期节奏。")

    # ── 动作清单 ──
    actions = [{"level": x["level"], "name": x["name"], "act": x["act"]}
               for x in sig if x["level"] in ("red", "amber")]
    if not actions:
        actions = [{"level": "green", "name": "全部指标", "act": "无需动作，按周复核"}]
    if left_biz > 0:
        need = max(TARGET_FLOOR - rev_mtd, 0)
        if left_biz >= 3:
            actions.append({"level": "info", "name": "剩余目标",
                            "act": f"本月剩 {left_biz} 个工作日，守 600 万还需确认 {need:,.0f} 万，"
                                   f"日均 {need / max(left_biz, 1):,.1f} 万"})
        else:
            actions.append({"level": "info", "name": "本月已定型",
                            "act": f"本月只剩 {left_biz} 个工作日，缺口 {need:,.0f} 万已难自然补齐，"
                                   f"建议并入下月目标重排，并把客场排期前移"})
    if need600:
        n6 = {d.get("scen"): d.get("need_k") for d in need600 if isinstance(d, dict)}
        if n6.get("中性"):
            actions.append({"level": "info", "name": "下月配档",
                            "act": f"守 600 万所需客场约见：乐观 {n6.get('乐观')} 场 / "
                                   f"中性 {n6.get('中性')} 场 / 悲观 {n6.get('悲观')} 场（按对应成交率情景）"})

    payload = {
        "meta": {
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "cut": str(cut.date()),
            "cur_mon": cur_mon,
            "day_n": day_n,
            "mon_days": mon_days,
            "left_days": left_days,
            "left_biz": left_biz,
            "time_pct": round(time_pct, 1),
            "sources": [{"name": os.path.basename(f), "mtime": datetime.fromtimestamp(
                os.path.getmtime(f)).strftime("%Y-%m-%d %H:%M"), "size": os.path.getsize(f)}
                for f in orders_f + [rev_f, ledger_f] if f],
            "rows": ometa,
            "targets": {"floor": TARGET_FLOOR, "ideal": TARGET_IDEAL},
            "thr": {"meet_warn": MEET_WARN, "meet_ok": MEET_OK, "rate_warn": RATE_WARN,
                    "refund_warn": REFUND_WARN, "refund_red": REFUND_RED, "per_warn": PER_WARN},
        },
        "month": month,
        "mtd": {
            "rev": round(rev_mtd, 1), "rev_proj": round(rev_proj, 1) if rev_proj else None,
            "rev_linear": round(rev_linear, 1),
            "rho_rev": round(r_rev, 3) if r_rev else None,
            "meet": meet_mtd, "meet_k": meet_k_mtd, "deal": deal_mtd,
            "meet_k_proj": round(meet_k_proj, 0) if meet_k_proj else None,
            "meet_k_linear": round(meet_k_linear, 0),
            "rho_meet": round(r_meet, 3) if r_meet else None,
            "rate": rt, "rate_raw": rt_raw,
            "rate_k_adj": cur.get("rate_k_adj") if cur else None,
            "refund": ref_mtd, "refund_rate": rr,
        },
        "ytd": {"rev": ytd_rev, "meet": ytd_meet, "meet_k": ytd_meet_k, "deal": ytd_deal,
                "annual_target": ann_target,
                "annual_floor": float(annual_floor["total"]) if annual_floor else None,
                "ytime_pct": round(ytime_pct, 1)},
        "sig": sig, "overall": overall, "summary": summary, "actions": actions,
        "weeks": weeks, "wk_stats": {"last4": round(wk_last4, 1) if wk_last4 else None,
                                     "prev4": round(wk_prev4, 1) if wk_prev4 else None,
                                     "last_full": wk_run, "cur": wk_cur},
        "fc": fc_rows, "fc_neutral": fc_neutral,
        "annual": annual,
        "rho": {"rev": {str(k): round(v, 3) for k, v in sorted(rho_rev.items())},
                "meet": {str(k): round(v, 3) for k, v in sorted(rho_meet.items())}},
        "prev": prev,
    }

    # ── 输出 ──
    jsdata = json.dumps(payload, ensure_ascii=False, default=str)
    with open(OUT_JS, "w", encoding="utf-8") as f:
        f.write("window.__MON_EXT__=" + jsdata + ";\n")

    if os.path.exists(TEMPLATE):
        tpl = open(TEMPLATE, encoding="utf-8").read()
        css_path = os.path.join(ROOT, "src", "templates", "base.css.html")
        css = open(css_path, encoding="utf-8").read() if os.path.exists(css_path) else ""
        html = tpl.replace("__CSS__", css).replace("__PAYLOAD__", jsdata)
        with open(OUT_HTML, "w", encoding="utf-8") as f:
            f.write(html)
        log(f"已生成 {os.path.basename(OUT_HTML)}（{os.path.getsize(OUT_HTML)/1024:.0f} KB）")
    else:
        log(f"⚠ 模板缺失：{TEMPLATE}（仅输出 data.js）")

    json.dump({"fp": fp, "at": datetime.now().isoformat(timespec="seconds")},
              open(STATE, "w", encoding="utf-8"))

    log(f"完成，用时 {(datetime.now()-t0).total_seconds():.1f}s | "
        f"本月业绩 {rev_mtd:,.0f} 万 / 客场约见 {meet_k_mtd} 场 / 成交率 {rt}% / 总体 {overall}")
    return True


if __name__ == "__main__":
    try:
        build(force="--force" in sys.argv or "-f" in sys.argv)
    except Exception:
        log("运行失败：\n" + traceback.format_exc())
        sys.exit(1)
