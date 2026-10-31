# -*- coding: utf-8 -*-
"""
Gold 层构建：把 CRM 原始导出（订单表 / 收入表 / 花名册）加工成看板用的统一事实表。

输入  data/raw/  收入表、订单表（可多份按年月）、花名册
输出  build/payload.json   —— 看板唯一数据源
      build/new_cases.csv  —— 案件级新接事实（区域专项分析复用）
      build/consult.csv    —— 线索级咨询费事实

三张事实表
  fact_new   成案事实·案件级：金额按行归属，件数按业绩比例拆分（每案恰 1 件）
  fact_all   业绩事实·行级：收入表全部行求和，不剔除撞单
  fact_fee   咨询费事实·线索级：三条规则并集，可重复计入

关键口径（详见 docs/03-metric-dictionary.md）
  1. 件数按业绩比例拆分
        W = 同案件「含"新接"」金额合计
        X = 同案件「新接案件」类记录数
        V = 0        若该行不是新接案件类
          = K/W      若 X>=2 且 W>0（撞单按金额比例拆）
          = 1        其余（单记录案件整件）
     参考公式「E=新接案件 即给 1」会使 43 个多记录案件各算 2 件；
     「孤立撞单标签(X=1)」会使 2 个案件丢件。此处统一为「每案恰 1 件」。
  2. 咨询费区域：订单归属人 → 花名册小组；是否合作约见=是 时用约见归属人1/2 按份额拆
  3. 咨询费覆盖 1–10 月（订单表 1–9 月为一份导出，10 月单独一份）
"""
import os, json, warnings
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DATA = os.path.join(ROOT, "data", "raw")
OUT = os.path.join(ROOT, "build")
os.makedirs(OUT, exist_ok=True)

F_REV = os.path.join(DATA, "收入表_2026.1-2026.10.xlsx")
F_ORDER = [os.path.join(DATA, "订单表_2026.1-2026.9.xlsx"),
           os.path.join(DATA, "订单表_2026.10.xlsx")]
F_ROSTER = os.path.join(DATA, "花名册.xlsx")

# 区域代号：需与 data/raw 中「区域」列取值一致。仓库内已代号化，
# 实际使用时改这里、或设环境变量 REGIONS="A区,B区,C区,D区"。
REGIONS = [r.strip() for r in os.environ.get("REGIONS", "A区,B区,C区,D区").split(",") if r.strip()]

# 小组名不以区域名开头时（例如「河北…组」），靠别名词表回挂到区域。
# 词表放 data/raw/region_alias.json，属机构敏感信息，不入仓。
_alias_file = os.path.join(DATA, "region_alias.json")
REGION_ALIAS = json.load(open(_alias_file, encoding="utf-8")) if os.path.exists(_alias_file) else {}
OTHER = "其他"
UNASSIGNED = "其他（未归属）"
CONSULT_THRESHOLD = 5000

DQ = []
def dq(level, item, detail, impact=""):
    DQ.append({"level": level, "item": item, "detail": detail, "impact": impact})


def money(s):
    if pd.isna(s):
        return 0.0
    if isinstance(s, (int, float, np.integer, np.floating)):
        return float(s)
    t = str(s).replace("人民币(CNY)", "").replace("CNY", "").replace(",", "").strip()
    if t in ("", "/", "-", "nan", "None", "."):
        return 0.0
    try:
        return float(t)
    except ValueError:
        return 0.0


def pay_kind(t):
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


def to_region(t):
    """小组名 / 部门名 → 区域。先用区域代号前缀匹配，再退到别名词表。"""
    for k in REGIONS:
        if t.startswith(k):
            return k
    for reg, keys in REGION_ALIAS.items():
        for k in keys:
            if k in t:
                return reg
    return OTHER


def norm_roster_region(x):
    if not isinstance(x, str):
        return None
    return to_region(x.strip())


# ══════════════════════════════════════════════
# 1. 收入表
# ══════════════════════════════════════════════
def load_revenue():
    xl = pd.ExcelFile(F_REV)
    frames = []
    for sh in xl.sheet_names:
        raw = pd.read_excel(xl, sheet_name=sh, header=None, usecols=range(0, 20))
        d = raw.iloc[2:].copy()
        d.columns = [str(c).strip() for c in raw.iloc[1].tolist()]
        d["_sheet"] = sh
        frames.append(d)
    df = pd.concat(frames, ignore_index=True)
    df = df[df["序号"].notna() | df["案件编号"].notna()].copy()

    df["amt"] = pd.to_numeric(df["业绩金额"], errors="coerce")
    bad = df["amt"].isna() & df["业绩金额"].notna()
    df["amt"] = df["amt"].fillna(0)
    df["case_id"] = df["案件编号"].astype(str).str.strip().replace({"nan": None})
    df["tag"] = df["新接/尾款"].fillna("").astype(str).str.strip()
    df["team"] = df["所属小组"].fillna("未填写").astype(str).str.strip()
    df["company"] = df["所属公司"].fillna("未填写").astype(str).str.strip()
    df["biz_date"] = pd.to_datetime(df["时间"], errors="coerce")
    df["mon"] = df["_sheet"].str.replace("月", "").astype(int)
    df["kind"] = df["tag"].map(pay_kind)
    df["region"] = df["team"].map(to_region)

    dq("info", "收入表读取",
       f"{len(df):,} 行有效记录，覆盖 {'、'.join(xl.sheet_names)}",
       f"业绩合计 {df['amt'].sum():,.0f} 元（全部行求和，不做去重）")
    if bad.sum():
        dq("warn", "业绩金额非数值", f"{int(bad.sum())} 行无法解析为数字", "已按 0 处理")
    noid = int(df["case_id"].isna().sum())
    if noid:
        dq("warn", "无案件编号的记录", f"{noid} 行业绩未填案件编号（人工调整、回款提点等）",
           "计入总业绩，但无法归属成案量")
    return df


# ══════════════════════════════════════════════
# 2. 成案：金额按行求和，件数按业绩比例拆分（V 列）
# ══════════════════════════════════════════════
NEW_TAGS = ("新接案件", "新接案件-撞单")

def build_cases(df):
    new = df[df["kind"] == "新接"].copy()

    W = (df[df["tag"].str.contains("新接", na=False)]
         .groupby("case_id")["amt"].sum())
    X = (df[df["tag"].str.contains("新接案件", na=False)]
         .groupby("case_id").size())

    new["W"] = new["case_id"].map(W).fillna(0.0)
    new["X"] = new["case_id"].map(X).fillna(0).astype(int)

    def calc_v(tag, x, k, w):
        if tag not in NEW_TAGS:
            return 0.0
        if x >= 2:
            if w > 0:
                return k / w
            return 1.0 / x if x else 0.0
        return 1.0

    new["v"] = [calc_v(t, x, k, w)
                for t, x, k, w in zip(new["tag"], new["X"], new["amt"], new["W"])]

    cases = new[new["v"] > 0]
    cnt_share = cases["v"].sum()
    per = cases.groupby("case_id")["v"].sum()

    dq("info", "件数按业绩比例拆分（V 列口径）",
       f"{int((new.groupby('case_id').size() > 1).sum()):,} 个案件存在多条新接记录，"
       f"按 K/W 金额比例拆分件数；拆分后件数合计 {cnt_share:,.2f} 件，案件 {cases['case_id'].nunique():,} 个",
       "每案件件数份额之和恰为 1，区域/小组件数可出现小数")
    dq("info", "参考公式边界修正",
       "参考表 V 列对「E=新接案件」直接给 1，会使 43 个多记录案件各算 2 件；"
       "另有 2 个案件仅有 1 条「新接案件-撞单」记录（202608031、GS2605034）会被判 0 件",
       "已按「X≥2 按金额比例拆、否则算 1 件」统一，确保每案恰 1 件")

    off = per[(per - 1).abs() > 1e-6]
    if len(off):
        dq("warn", "件数份额异常案件", f"{len(off)} 个案件 V 合计不等于 1", "请核查")

    zero_w = cases[(cases["X"] >= 2) & (cases["W"] <= 0)]["case_id"].nunique()
    if zero_w:
        dq("warn", "新接金额为 0 的撞单案件", f"{zero_w} 个案件按 1/X 均分件数", "金额无法提供比例")

    n0 = cases[cases["amt"] == 0]["case_id"].nunique()
    if n0:
        dq("warn", "零金额成案", f"{n0} 个案件新接金额为 0（律师条线为主）",
           "计入成案量但不贡献成案金额，会拉低客单价")
    return new, cases


def cross_check(new):
    o = pd.concat([pd.read_excel(f, sheet_name="订单导出明细") for f in F_ORDER],
                  ignore_index=True)
    o["case_id"] = o["案件编号"].astype(str).str.strip()
    o["ht"] = o["合同总金额"].map(money)
    om = o.groupby("case_id")["ht"].max()
    s = new.groupby("case_id")["amt"].sum().rename("sum")
    cmp = s.to_frame().join(om, how="inner")
    dq("info", "拆分规则校验",
       f"可与订单表匹配的 {len(cmp):,} 个新接案件，求和金额 vs『合同总金额』"
       f"完全一致率 {((cmp['sum']-cmp['ht']).abs()<=1).mean():.1%}，"
       f"中位差 {(cmp['sum']-cmp['ht']).abs().median():,.0f} 元",
       "交叉验证支持「多行为业绩拆分，求和即为案件金额」")


# ══════════════════════════════════════════════
# 3. 咨询费（订单表 1–9 + 10 月）
# ══════════════════════════════════════════════
def load_consult():
    parts = []
    for f in F_ORDER:
        d = pd.read_excel(f, sheet_name="订单导出明细")
        parts.append(d)
    o = pd.concat(parts, ignore_index=True)
    n_raw = len(o)
    dup = o["订单编号"].notna() & o.duplicated(subset=["订单编号"], keep="first")
    n_dup = int(dup.sum())
    o = o[~dup].copy()
    if n_dup:
        dq("info", "订单编号去重", f"合并 {n_dup} 行跨文件重复订单（订单编号相同）", "保留首行")

    o["lv"] = o["差旅费金额"].map(money)
    o["ht"] = o["合同总金额"].map(money)
    o["visit"] = pd.to_datetime(o["约见时间"], errors="coerce")
    o["pay"] = pd.to_datetime(o["费用支付时间"], errors="coerce")
    o["mon"] = o["visit"].dt.month
    o = o[o["mon"].notna()].copy()
    o["mon"] = o["mon"].astype(int)

    # ── 区域份额：订单归属人→花名册；合作约见→约见归属人1/2 ──
    ros = pd.read_excel(F_ROSTER, header=0)
    ros["坐席"] = ros["坐席"].astype(str).str.strip()
    ros_team = dict(zip(ros["坐席"], ros["小组"].astype(str).str.strip()))
    def look(x):
        if not isinstance(x, str) or x in ("nan", ""):
            return None
        t = ros_team.get(x.strip())
        if not isinstance(t, str) or t in ("", "nan"):
            return None
        return t

    o["coop"] = o["是否合作约见"].astype(str).str.strip().eq("是")
    o["o1"] = o["约见归属人1"].astype(str).str.strip().replace({"nan": None})
    o["o2"] = o["约见归属人2"].astype(str).str.strip().replace({"nan": None})
    o["w1"] = pd.to_numeric(o["约见归属人1个数"], errors="coerce")
    o["w2"] = pd.to_numeric(o["约见归属人2个数"], errors="coerce")
    o["r_owner"] = o["订单归属人"].astype(str).str.strip().replace({"nan": None}).map(look)
    o["r_o1"] = o["o1"].map(look)
    o["r_o2"] = o["o2"].map(look)

    def shares(row):
        if row["coop"]:
            # iterrows 会把 None 转成 nan，必须先强制校验类型
            a = row["r_o1"] if isinstance(row["r_o1"], str) else None
            b = row["r_o2"] if isinstance(row["r_o2"], str) else None
            w1 = float(row["w1"]) if pd.notna(row["w1"]) else 0.5
            w2 = float(row["w2"]) if pd.notna(row["w2"]) else 0.5
            if a and b:
                return ((a, w1), (b, w2))
            if a:
                return ((a, 1.0),)
            if b:
                return ((b, 1.0),)
        if isinstance(row["r_owner"], str):
            return ((row["r_owner"], 1.0),)
        return ((UNASSIGNED, 1.0),)

    o["shares"] = [shares(r) for _, r in o.iterrows()]

    # ── 线索 ──
    lead = o["差旅码（线索id）"].astype(str).str.strip().replace({"nan": None})
    o["lead"] = lead.where(lead.notna(), "C" + o["案件编号"].astype(str)).copy()
    o["lead"] = o["lead"].where(o["lead"].notna() & o["lead"].ne("Cnan"),
                                "O" + o["订单编号"].astype(str))

    a = o["lv"] > 0
    b = (o["约见类型"] == "主场约见") & (o["ht"] >= CONSULT_THRESHOLD)
    c = (o["约见类型"] == "直播签约") & (o["ht"] >= CONSULT_THRESHOLD)
    o["f_travel"], o["f_main"], o["f_live"] = a, b, c
    o["hit"] = a | b | c

    f = o[o["hit"]].copy()
    f = f.sort_values(["lv", "ht"], ascending=False)
    g = f.groupby("lead").agg(
        mon=("mon", "min"), owner=("订单归属人", "first"),
        shares=("shares", "first"), coop=("coop", "max"),
        r_owner=("r_owner", "first"), r_o1=("r_o1", "first"),
        f_travel=("f_travel", "max"), f_main=("f_main", "max"), f_live=("f_live", "max"),
        lv=("lv", "sum"), ht=("ht", "max"),
    ).reset_index()

    n_dup_lead = len(f) - len(g)
    if n_dup_lead:
        dq("info", "咨询费线索多行",
           f"{n_dup_lead} 行属于重复线索（撞单/重复约见）",
           "按差旅费金额最大的行确定区域归属，线索数不重复计")

    rows = []
    for _, r in g.iterrows():
        for tm, w in r["shares"]:
            rows.append({"mon": int(r["mon"]), "region": to_region(tm),
                         "team": tm, "lead": r["lead"], "w": float(w),
                         "owner": (r["owner"] if isinstance(r["owner"], str) else "未填写"),
                         "f_travel": bool(r["f_travel"]), "f_main": bool(r["f_main"]),
                         "f_live": bool(r["f_live"]), "lv": float(r["lv"]), "ht": float(r["ht"])})
    fee = pd.DataFrame(rows)

    cov = sorted(o["mon"].unique().tolist())
    paymon = pd.to_datetime(o.loc[a, "费用支付时间"], errors="coerce").dt.month
    same = float((paymon == o.loc[a, "mon"]).mean())
    dq("info", "咨询费口径",
       f"线索 {len(g):,} 条（区域份额合计 {fee['w'].sum():,.1f}）："
       f"差旅费驱动 {int(g['f_travel'].sum()):,}、主场成案 {int(g['f_main'].sum()):,}、"
       f"直播签约 {int(g['f_live'].sum()):,}（按线索计，可重叠）",
       f"门槛为合同总金额 ≥ {CONSULT_THRESHOLD:,} 元；月份按「约见时间」；"
       f"差旅费收取与约见同月率 {same:.0%}")

    dq("info", "咨询费小组归属",
       f"人员→花名册「小组」：合作约见 {int(g['coop'].sum()):,} 条按「约见归属人1/2」拆分，其余按「订单归属人」",
       f"订单归属人可映射到小组 {g['r_owner'].notna().mean():.1%}；"
       f"合作线索中约见归属人可映射 {g[g['coop']]['r_o1'].notna().mean():.0%}；"
       f"未映射的归入「{UNASSIGNED}」")

    nun = fee[fee["team"] == UNASSIGNED]
    if len(nun):
        dq("warn", "咨询费未映射到小组",
           f"{nun['lead'].nunique():,} 条线索（{nun['w'].sum():.1f} 条份额）的归属人不在花名册中",
           f"单列「{UNASSIGNED}」，计入总量但不进任何小组")
    oth = fee[fee["region"] == OTHER]
    if len(oth):
        dq("warn", "咨询费落在他区域",
           f"{oth['w'].sum():,.1f} 条来自律师条线/已解散区域",
           "归入「其他」，与业绩口径一致")

    dq("info", "订单表合并",
       f"1–9 月 + 10 月合计 {n_raw:,} 行，去重后 {len(o):,} 行有约见时间",
       f"咨询费覆盖 {'、'.join(str(x) for x in cov)} 月")

    return fee, cov


# ══════════════════════════════════════════════
# 4. 主流程
# ══════════════════════════════════════════════
def main():
    raw = load_revenue()
    new, cases = build_cases(raw)
    cross_check(new)
    fee, fee_months = load_consult()

    tot_rev = float(raw["amt"].sum())
    case_amt = float(cases["amt"].sum())
    case_share = float(cases["v"].sum())
    case_n = int(cases["case_id"].nunique())
    fee_n = int(fee["lead"].nunique())

    bylead = fee.groupby("lead").agg(ft=("f_travel", "max"), fm=("f_main", "max"), fl=("f_live", "max"))
    kpi = {
        "total_revenue": round(tot_rev, 2),
        "case_count": case_n,
        "case_share": round(case_share, 3),
        "case_amount": round(case_amt, 2),
        "arpu": round(case_amt / case_share, 1) if case_share else None,
        "fee_count": fee_n,
        "fee_travel": int(bylead["ft"].sum()),
        "fee_main": int(bylead["fm"].sum()),
        "fee_live": int(bylead["fl"].sum()),
        "fee_map_rate": round(float(fee[fee["team"] != UNASSIGNED]["lead"].nunique()) / max(fee_n, 1) * 100, 1),
        "case_rows": int(len(cases)),
        "new_share": round(case_amt / tot_rev * 100, 1) if tot_rev else None,
    }

    mon_rows = []
    for m in sorted(raw["mon"].unique()):
        a = raw[raw["mon"] == m]
        c = cases[cases["mon"] == m]
        cn = float(c["v"].sum())
        f = fee[fee["mon"] == m]
        mon_rows.append({
            "mon": int(m),
            "revenue": round(float(a["amt"].sum()), 2),
            "case_count": round(cn, 2),
            "case_amount": round(float(c["amt"].sum()), 2),
            "arpu": round(float(c["amt"].sum()) / cn, 1) if cn else None,
            "fee_count": round(float(f["w"].sum()), 2),
            "complete": bool(m < 10),
            "fee_ok": bool(m in fee_months),
        })

    def dim_table(key, order=None):
        rows = []
        keys = set(raw[key]) | set(cases[key])
        if key in fee.columns:
            keys |= set(fee[key])
        for k in sorted(keys, key=str):
            a = raw[raw[key] == k]
            c = cases[cases[key] == k]
            cn = float(c["v"].sum())
            f = fee[fee[key] == k] if key in fee.columns else fee.iloc[0:0]
            rows.append({
                "k": k,
                "revenue": round(float(a["amt"].sum()), 2),
                "case_count": round(cn, 2),
                "case_amount": round(float(c["amt"].sum()), 2),
                "arpu": round(float(c["amt"].sum()) / cn, 1) if cn else None,
                "new_share": round(float(c["amt"].sum()) / a["amt"].sum() * 100, 1) if a["amt"].sum() else None,
                "fee_count": round(float(f["w"].sum()), 2),
            })
        if order:
            rows.sort(key=lambda x: (order.index(x["k"]) if x["k"] in order else 99))
        else:
            rows.sort(key=lambda x: -x["revenue"])
        return rows

    regions = dim_table("region", REGIONS + [OTHER])
    teams = dim_table("team")
    companies = dim_table("company")

    reg_mon = []
    for r in regions:
        k = r["k"]
        for m in sorted(raw["mon"].unique()):
            a = raw[(raw["region"] == k) & (raw["mon"] == m)]
            c = cases[(cases["region"] == k) & (cases["mon"] == m)]
            f = fee[(fee["region"] == k) & (fee["mon"] == m)]
            reg_mon.append({"region": k, "mon": int(m),
                            "revenue": round(float(a["amt"].sum()), 2),
                            "case_count": round(float(c["v"].sum()), 2),
                            "case_amount": round(float(c["amt"].sum()), 2),
                            "fee_count": round(float(f["w"].sum()), 2)})

    fee_rm = []
    for k in REGIONS + [OTHER]:
        for m in fee_months:
            fee_rm.append({"region": k, "mon": int(m),
                           "fee_count": round(float(fee[(fee["region"] == k) &
                                                        (fee["mon"] == m)]["w"].sum()), 2)})

    tr = {}
    for t, g in new.groupby("team"):
        tr[t] = g["region"].iloc[0]
    for t, g in fee.groupby("team"):
        tr.setdefault(t, g["region"].iloc[0])
    team_rows = [{"team": t, "region": v} for t, v in tr.items()]

    fact_all = raw[["mon", "region", "team", "company", "kind", "amt"]].copy()
    fact_all["amt"] = fact_all["amt"].round(2)

    fact_new = cases[["mon", "region", "team", "company", "case_id", "amt", "v"]].copy()
    fact_new["amt"] = fact_new["amt"].round(2)
    fact_new["v"] = fact_new["v"].round(4)
    fact_new = fact_new.rename(columns={"case_id": "cid"})

    fact_fee = fee[["mon", "region", "team", "lead", "w", "f_travel", "f_main", "f_live", "lv", "ht"]].copy()
    fact_fee = fact_fee.rename(columns={"lead": "lid", "w": "wt",
                                        "f_travel": "ft", "f_main": "fm", "f_live": "fl"})
    fact_fee["lv"] = fact_fee["lv"].round(2)
    fact_fee["ht"] = fact_fee["ht"].round(2)

    fee_tag = []
    for k, cond in [("差旅费", "ft"), ("主场成案", "fm"), ("直播签约", "fl")]:
        sub = fact_fee[fact_fee[cond]]
        fee_tag.append({"tag": k, "cnt": round(float(sub["wt"].sum()), 2),
                        "mon": {str(m): round(float(sub[sub["mon"] == m]["wt"].sum()), 2) for m in fee_months}})
    fee_tag.append({"tag": "总线索", "cnt": round(float(fact_fee["wt"].sum()), 2),
                    "mon": {str(m): round(float(fact_fee[fact_fee["mon"] == m]["wt"].sum()), 2) for m in fee_months}})

    oth = raw[raw["region"] == OTHER]
    dq("warn", "非四大区域",
       f"{len(oth):,} 行业绩 {oth['amt'].sum():,.0f} 元归入「其他（已解散）」",
       "已从主视图剥离，仅保留以对齐财务总额")

    last = raw["biz_date"].max()
    if last is not None and pd.notna(last):
        dq("warn", "10 月业绩未完整",
           f"收入表最新记录 {last.strftime('%Y-%m-%d')}，10 月仅 {len(raw[raw['mon']==10]):,} 行业绩",
           "趋势图已用浅色标注；咨询费 10 月已完整（订单表截到 10/29）")

    payload = {
        "meta": {
            "generated_at": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
            "window": "2026.1–10",
            "last_date": last.strftime("%Y-%m-%d") if pd.notna(last) else None,
            "regions": REGIONS,
            "other": OTHER,
            "fee_months": fee_months,
            "threshold": CONSULT_THRESHOLD,
        },
        "kpi": kpi,
        "month": mon_rows,
        "region": regions,
        "team": teams,
        "company": companies,
        "reg_mon": reg_mon,
        "fee_rm": fee_rm,
        "fee_tag": fee_tag,
        "team_region": team_rows,
        "dq": DQ,
        "fact_all": fact_all.to_dict("records"),
        "fact_new": fact_new.to_dict("records"),
        "fact_fee": fact_fee.to_dict("records"),
    }

    def _clean(o):
        if isinstance(o, dict):
            return {k: _clean(v) for k, v in o.items()}
        if isinstance(o, list):
            return [_clean(v) for v in o]
        if isinstance(o, float) and (np.isnan(o) or np.isinf(o)):
            return None
        if isinstance(o, (np.integer,)):
            return int(o)
        if isinstance(o, (np.floating,)):
            return None if np.isnan(o) else float(o)
        if isinstance(o, (np.bool_,)):
            return bool(o)
        return o

    payload = _clean(payload)
    with open(os.path.join(OUT, "payload.json"), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, separators=(",", ":"), allow_nan=False)

    raw.to_csv(os.path.join(OUT, "revenue_raw.csv"), index=False, encoding="utf-8-sig")
    new.to_csv(os.path.join(OUT, "new_cases.csv"), index=False, encoding="utf-8-sig")
    fee.to_csv(os.path.join(OUT, "consult.csv"), index=False, encoding="utf-8-sig")

    print("KPI:", json.dumps(kpi, ensure_ascii=False))
    print("\n月度:")
    print(pd.DataFrame(mon_rows).to_string(index=False))
    print("\n区域:")
    print(pd.DataFrame(regions).to_string(index=False))
    print("\nDQ:")
    for d in DQ:
        print(f"  [{d['level']}] {d['item']}: {d['detail']}")
    return payload


if __name__ == "__main__":
    main()
