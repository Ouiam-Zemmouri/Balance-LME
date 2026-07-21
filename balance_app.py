import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import glob
import os
import re
import openpyxl

# ===== PASSWORD PROTECTION =====
PASSWORD = "Finance26"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 Login Required")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if password == PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password")
    st.stop()

st.set_page_config(page_title="Balance LME", page_icon="⚖️",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
.main,[data-testid="stAppViewContainer"]{background:#0f1b3d;}
[data-testid="stHeader"]{background:transparent;}
.block-container{padding-top:1.5rem;padding-bottom:2rem;}
.title-banner{
  background:linear-gradient(135deg,#0f1b3d 0%,#1B2F6E 50%,#0f1b3d 100%);
  border:1px solid rgba(139,94,60,0.3);border-top:3px solid #8B5E3C;
  border-radius:16px;padding:30px 40px;margin-bottom:28px;text-align:center;
  box-shadow:0 6px 60px rgba(27,47,110,0.3);
}
.title-banner h1{
  font-size:2rem;font-weight:800;letter-spacing:5px;margin:0;text-transform:uppercase;
  background:linear-gradient(90deg,#FFFFFF,#d4a97a,#FFFFFF,#d4a97a,#FFFFFF);
  -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;
}
.title-banner p{color:#8fa3c8;font-size:0.8rem;margin:8px 0 0 0;letter-spacing:2px;}
.kpi-card{
  background:linear-gradient(145deg,#162040,#1a2850);
  border:1px solid rgba(255,255,255,0.06);border-left:3px solid #8B5E3C;
  border-radius:12px;padding:18px 14px;text-align:center;
  height:118px;display:flex;flex-direction:column;justify-content:center;
  box-shadow:0 2px 24px rgba(0,0,0,0.5);transition:all 0.2s ease;
}
.kpi-card:hover{transform:translateY(-2px);box-shadow:0 6px 32px rgba(27,47,110,0.25);}
.section-header{
  color:#8B5E3C;font-size:0.72rem;font-weight:700;text-transform:uppercase;letter-spacing:3px;
  border-bottom:1px solid rgba(139,94,60,0.25);padding-bottom:8px;margin:28px 0 18px 0;
}
[data-testid="stSidebar"]{background:#0d1530;border-right:1px solid rgba(27,47,110,0.2);}
.filter-label{color:#8fa3c8;font-size:0.6rem;font-weight:700;text-transform:uppercase;letter-spacing:2px;margin:14px 0 3px 0;}
[data-baseweb="tag"]{background:rgba(27,47,110,0.3) !important;border:1px solid rgba(27,47,110,0.5) !important;color:#FFFFFF !important;border-radius:6px !important;}
::-webkit-scrollbar{width:3px;height:3px;}
::-webkit-scrollbar-track{background:#162040;}
::-webkit-scrollbar-thumb{background:#1B2F6E;border-radius:4px;}
::-webkit-scrollbar-thumb:hover{background:#8B5E3C;}
.stDataFrame{border-radius:10px;overflow:hidden;border:1px solid rgba(139,94,60,0.1);}
</style>
""", unsafe_allow_html=True)

# ── COLORS ──
NAVY    = "#0f1b3d"; NAVY_MD = "#162040"; NAVY_LT = "#1B2F6E"
COPPER  = "#8B5E3C"; COP_LT  = "#b07d52"; COP_XL  = "#d4a97a"
WHITE   = "#FFFFFF"; ICE     = "#c8d5ea"; SLATE   = "#8fa3c8"
TEAL    = "#0ea5a0"; GOLD    = "#c49040"

LAY = dict(
    paper_bgcolor=NAVY, plot_bgcolor="#0d1730",
    font=dict(color=SLATE, family="Inter", size=12),
    margin=dict(t=55, b=45, l=60, r=25),
    legend=dict(bgcolor="rgba(15,27,61,0.97)", bordercolor="rgba(139,94,60,0.25)",
                borderwidth=1, font=dict(color=ICE, size=11)),
    xaxis=dict(gridcolor="rgba(255,255,255,0.04)", zeroline=False,
               tickfont=dict(color=SLATE, size=11)),
    yaxis=dict(gridcolor="rgba(255,255,255,0.04)", zeroline=False,
               tickfont=dict(color=SLATE, size=11)),
    title_font=dict(color=COP_LT, size=13, family="Inter"),
    hoverlabel=dict(bgcolor=NAVY_MD, bordercolor="rgba(139,94,60,0.4)",
                    font=dict(color=WHITE, size=12)),
)

def kpi(col, label, val, color=COPPER):
    col.markdown(f"""<div class="kpi-card" style="border-left-color:{color};">
      <div style="color:{WHITE};font-size:0.68rem;font-weight:700;text-transform:uppercase;
                  letter-spacing:1.8px;margin-bottom:10px;">{label}</div>
      <div style="color:{WHITE};font-size:1.55rem;font-weight:800;line-height:1.1;">{val}</div>
      </div>""", unsafe_allow_html=True)

def sec(icon, title):
    st.markdown(f'<div class="section-header">{icon}&nbsp; {title}</div>', unsafe_allow_html=True)

def alay(fig, **kw):
    fig.update_layout(**{**LAY, **kw}); return fig

# ── LME BALANCE — PARSING ──
ENTITY_CODE_MAP = {
    "KT": "COFICAB Kenitra", "KEN": "COFICAB Kenitra",
    "MR": "COFICAB Maroc", "MA": "COFICAB Maroc",
    "INTL": "COFICAB International", "INT": "COFICAB International",
}
MONTH_NUM_MAP = {"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"May","06":"Jun",
                  "07":"Jul","08":"Aug","09":"Sep","10":"Oct","11":"Nov","12":"Dec"}
BALANCE_FOLDER = "balance_files"

def parse_lme_balance_file(path):
    """Parse one monthly 'LME balance calculation under FIFO method' Excel file."""
    fname = os.path.basename(path)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        return None, f"❌ {fname}: unable to read file ({e})"

    ws = wb[wb.sheetnames[0]]
    header_row = None
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, 2).value).strip() == "LME Projects":
            header_row = r
            break
    if header_row is None:
        return None, f"❌ {fname}: could not locate the 'LME Projects' balance table."

    rows = []
    r = header_row + 1
    while r <= ws.max_row:
        proj = ws.cell(r, 2).value
        if proj is None or str(proj).strip() == "":
            break
        proj = str(proj).strip()
        rows.append({
            "Fixation":        proj,
            "Qty_Sold_T":      ws.cell(r, 3).value,
            "LME_Sales":       ws.cell(r, 4).value,
            "Sales_Value":     ws.cell(r, 5).value,
            "Qty_Stock_T":     ws.cell(r, 6).value,
            "LME_Stock":       ws.cell(r, 7).value,
            "Stock_Value":     ws.cell(r, 8).value,
            "Qty_Purchase_T":  ws.cell(r, 9).value,
            "LME_Purchase":    ws.cell(r, 10).value,
            "Purchase_Value":  ws.cell(r, 11).value,
            "Needs_Exceed_T":  ws.cell(r, 12).value,
            "Allocated_QTE":   ws.cell(r, 13).value,
            "LME_Realloc":     ws.cell(r, 14).value,
            "Realloc_Value":   ws.cell(r, 15).value,
            "Last_QTY_T":      ws.cell(r, 16).value,
            "LME_Final":       ws.cell(r, 17).value,
            "Final_Value":     ws.cell(r, 18).value,
            "LME_Balance_Eur": ws.cell(r, 19).value,
        })
        if proj.upper() == "TOTAL":
            break
        r += 1

    if not rows:
        return None, f"❌ {fname}: no fixation rows found under the balance table."

    df_bal = pd.DataFrame(rows)
    num_cols = [c for c in df_bal.columns if c != "Fixation"]
    df_bal[num_cols] = df_bal[num_cols].apply(pd.to_numeric, errors="coerce")

    m = re.search(r'COF[\s_-]*([A-Za-z]+).*?(\d{1,2})[\s._-](\d{4})', fname)
    if m:
        entity = ENTITY_CODE_MAP.get(m.group(1).upper(), f"COFICAB {m.group(1).upper()}")
        month_num, year = m.group(2).zfill(2), m.group(3)
        month_label = f"{MONTH_NUM_MAP.get(month_num, month_num)} {year}"
        month_key = f"{year}-{month_num}"
    else:
        entity, month_label, month_key = "Unknown", fname, fname

    df_bal.insert(0, "Entity", entity)
    df_bal.insert(1, "Month", month_label)
    df_bal.insert(2, "MonthKey", month_key)
    df_bal.insert(3, "SourceFile", fname)
    return df_bal, None

@st.cache_data(ttl=3600)
def load_all_balance_files():
    """Auto-load every .xlsx bundled in the balance_files/ folder of the repo."""
    files = sorted(
        f for f in glob.glob(os.path.join(BALANCE_FOLDER, "*.xlsx"))
        if not os.path.basename(f).startswith("~$")
    )
    dfs, errs = [], []
    for path in files:
        df, err = parse_lme_balance_file(path)
        (errs if err else dfs).append(err if err else df)
    if not dfs:
        return pd.DataFrame(), errs
    return pd.concat(dfs, ignore_index=True), errs

def generate_balance_insights(view_fix, view_tot):
    insights = []
    if view_fix.empty or view_tot.empty:
        return ["Not enough data to generate an analysis for the current selection."]

    vf = view_fix.dropna(subset=["LME_Sales"])
    if not vf.empty:
        best = vf.loc[vf["LME_Sales"].idxmin()]
        worst = vf.loc[vf["LME_Sales"].idxmax()]
        if best["Fixation"] != worst["Fixation"]:
            insights.append(
                f"The most favorable sales fixation is **{best['Fixation']}** at "
                f"**{best['LME_Sales']:.4f} €/kg**, while **{worst['Fixation']}** carries the "
                f"highest sales price at **{worst['LME_Sales']:.4f} €/kg**."
            )

    vs = view_fix.dropna(subset=["Qty_Stock_T"])
    vs = vs[vs["Qty_Stock_T"] > 0]
    if not vs.empty:
        top_stock = vs.loc[vs["Qty_Stock_T"].idxmax()]
        insights.append(
            f"**{top_stock['Fixation']}** carries the largest stock position at "
            f"**{top_stock['Qty_Stock_T']:.1f} T**, making it the main buffer in the FIFO valuation chain."
        )

    ne = view_fix.dropna(subset=["Needs_Exceed_T"])
    deficits = ne[ne["Needs_Exceed_T"] > 0.5]
    surplus  = ne[ne["Needs_Exceed_T"] < -0.5]
    if not deficits.empty:
        d = deficits.iloc[0]
        if not surplus.empty:
            s = surplus.iloc[0]
            insights.append(
                f"Sold quantities on **{d['Fixation']}** exceeded available stock and purchases by "
                f"**{d['Needs_Exceed_T']:.1f} T**; under FIFO this shortfall is reallocated from the "
                f"surplus fixation **{s['Fixation']}**."
            )
        else:
            insights.append(
                f"Sold quantities on **{d['Fixation']}** exceeded available stock and purchases by "
                f"**{d['Needs_Exceed_T']:.1f} T**, requiring reallocation from the next fixation in the FIFO sequence."
            )

    tot_balance = view_tot["LME_Balance_Eur"].sum()
    direction = "favorable" if tot_balance >= 0 else "unfavorable"
    insights.append(
        f"The overall LME balance is **{direction}**, at **€{tot_balance:,.0f}** — sales were valued "
        f"{'above' if tot_balance >= 0 else 'below'} the FIFO cost of stock and purchases consumed."
    )

    vb = view_fix.dropna(subset=["LME_Balance_Eur"])
    if not vb.empty:
        top_c = vb.reindex(vb["LME_Balance_Eur"].abs().sort_values(ascending=False).index).iloc[0]
        sign = "gain" if top_c["LME_Balance_Eur"] >= 0 else "loss"
        insights.append(
            f"**{top_c['Fixation']}** is the largest single contributor to this result, driving a "
            f"**{sign} of €{abs(top_c['LME_Balance_Eur']):,.0f}**."
        )
    return insights

# ── PALETTE (consistent per-entity colors across all charts) ──
PALETTE = [COPPER, TEAL, GOLD, "#5b8dd9", "#b07d52", "#9b7bd9"]

def entity_color_map(entities):
    return {e: PALETTE[i % len(PALETTE)] for i, e in enumerate(entities)}

# ── SIDEBAR ──
with st.sidebar:
    st.markdown("### ⚖️ Balance LME")
    st.caption("COFICAB Kenitra · COFICAB Maroc")
    if st.button("🔄 Refresh Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.markdown("---")

# ── TITLE ──
st.markdown("""<div class="title-banner">
  <h1>Balance LME</h1>
  <p>FIFO Method · Sales vs Stock &amp; Purchase Valuation</p>
</div>""", unsafe_allow_html=True)

bal_all, errors = load_all_balance_files()
for e in errors:
    st.warning(e)

if bal_all.empty:
    st.info(
        f"No balance files found yet. Add your monthly `.xlsx` files "
        f"(e.g. `COF_KT_-_LME_balance_06_2026.xlsx`) to the **`{BALANCE_FOLDER}/`** "
        f"folder in this repo, commit, then click 🔄 Refresh Data in the sidebar."
    )
    st.stop()

bal_all["Group"] = bal_all["Entity"] + " · " + bal_all["Month"]
ENT_COLOR = entity_color_map(sorted(bal_all["Entity"].unique()))

# ── FILTERS (sidebar) ──
with st.sidebar:
    st.markdown('<p class="filter-label">Entity</p>', unsafe_allow_html=True)
    ent_opts = sorted(bal_all["Entity"].unique())
    sel_e = st.multiselect("", ent_opts, default=ent_opts, key="bal_ent", label_visibility="collapsed")

    st.markdown('<p class="filter-label">Month</p>', unsafe_allow_html=True)
    month_map = bal_all[["MonthKey","Month"]].drop_duplicates().sort_values("MonthKey")
    sel_m = st.multiselect("", month_map["Month"].tolist(), default=month_map["Month"].tolist(),
                            key="bal_month", label_visibility="collapsed")

    st.markdown("---")
    st.caption(f"📁 {bal_all['SourceFile'].nunique()} file(s) loaded")
    st.caption(f"🏭 {bal_all['Entity'].nunique()} entit{'y' if bal_all['Entity'].nunique()==1 else 'ies'} · "
               f"{bal_all['MonthKey'].nunique()} month(s)")

view = bal_all[bal_all["Entity"].isin(sel_e) & bal_all["Month"].isin(sel_m)].copy()
view_fix = view[view["Fixation"].str.upper() != "TOTAL"].copy()
view_tot = view[view["Fixation"].str.upper() == "TOTAL"].copy()

if view.empty:
    st.warning("⚠️ No data matches the selected Entity / Month filters.")
    st.stop()

groups = sorted(view["Group"].unique())
month_order = month_map[month_map["Month"].isin(sel_m)]["Month"].tolist()

# ══════════════════════ KPI ROW ══════════════════════
tot_sales   = view_tot["Sales_Value"].sum()
tot_final   = view_tot["Final_Value"].sum()
tot_balance = view_tot["LME_Balance_Eur"].sum()
tot_qty     = view_tot["Qty_Sold_T"].sum()
bal_per_t   = tot_balance / tot_qty if tot_qty else 0
bal_color   = TEAL if tot_balance >= 0 else "#f43f5e"
n_fav       = (view_tot["LME_Balance_Eur"] >= 0).sum()
n_tot       = len(view_tot)

sec("📊","Key Indicators")
k1,k2,k3,k4,k5 = st.columns(5)
kpi(k1,"Sales Valuation (€)",            f"{tot_sales:,.0f}",  COPPER)
kpi(k2,"Stock + Purchase Valuation (€)", f"{tot_final:,.0f}",  NAVY_LT)
kpi(k3,"Net LME Balance (€)",            f"{tot_balance:,.0f}",bal_color)
kpi(k4,"Balance per Ton (€/T)",          f"{bal_per_t:,.1f}",  GOLD)
kpi(k5,"Favorable Periods",              f"{n_fav} / {n_tot}", TEAL if n_fav==n_tot else "#f43f5e")

st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

# ══════════════════════ ROW A — TREND + SPLIT ══════════════════════
rowA1, rowA2 = st.columns([2,1])

with rowA1:
    with st.container(border=True):
        sec("📅","Net LME Balance — Trend")
        trend = (view_tot.groupby(["Entity","MonthKey","Month"])["LME_Balance_Eur"]
                 .sum().reset_index().sort_values("MonthKey"))
        if trend["MonthKey"].nunique() > 1:
            figA1 = px.line(trend, x="Month", y="LME_Balance_Eur", color="Entity",
                             markers=True, category_orders={"Month": month_order},
                             color_discrete_map=ENT_COLOR)
            figA1.update_traces(line=dict(width=3), marker=dict(size=9))
            figA1.add_hline(y=0, line_dash="dot", line_color="rgba(255,255,255,0.2)")
            alay(figA1, showlegend=len(sel_e) > 1,
                 yaxis=dict(title="LME Balance (€)", gridcolor="rgba(255,255,255,0.04)", tickfont=dict(color=SLATE)),
                 xaxis=dict(title="", tickfont=dict(color=SLATE)))
            st.plotly_chart(figA1, use_container_width=True)
        else:
            figA1 = px.bar(view_tot, x="Entity", y="LME_Balance_Eur", color="Entity",
                            color_discrete_map=ENT_COLOR, text_auto=",.0f")
            figA1.add_hline(y=0, line_dash="dot", line_color="rgba(255,255,255,0.2)")
            alay(figA1, showlegend=False,
                 yaxis=dict(title="LME Balance (€)", gridcolor="rgba(255,255,255,0.04)", tickfont=dict(color=SLATE)),
                 xaxis=dict(title="", tickfont=dict(color=SLATE)))
            st.plotly_chart(figA1, use_container_width=True)
            st.caption("Add more monthly files to unlock the trend view.")

with rowA2:
    with st.container(border=True):
        sec("🥯","Valuation Split")
        donut_df = pd.DataFrame({
            "Component": ["Sales", "Stock + Purchase"],
            "Value": [tot_sales, tot_final]
        })
        figA2 = go.Figure(go.Pie(
            labels=donut_df["Component"], values=donut_df["Value"], hole=0.62,
            marker=dict(colors=[COPPER, NAVY_LT], line=dict(color=NAVY, width=2)),
            textinfo="percent", textfont=dict(color=WHITE, size=12)
        ))
        alay(figA2, showlegend=True,
             legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5,
                         font=dict(color=ICE, size=11)),
             annotations=[dict(text=f"€{tot_sales - tot_final:+,.0f}", x=0.5, y=0.5,
                                font=dict(size=15, color=bal_color, family="Inter"), showarrow=False)])
        st.plotly_chart(figA2, use_container_width=True)

# ══════════════════════ ROW B — FIXATION BREAKDOWN ══════════════════════
rowB1, rowB2 = st.columns(2)

with rowB1:
    with st.container(border=True):
        sec("📈","LME Balance by Fixation")
        figB1 = px.bar(view_fix, x="Fixation", y="LME_Balance_Eur", color="Group",
                        barmode="group", text_auto=",.0f",
                        color_discrete_sequence=PALETTE)
        figB1.add_hline(y=0, line_color="rgba(255,255,255,0.15)")
        figB1.update_traces(textfont=dict(size=10, color=ICE), textposition="outside")
        alay(figB1, showlegend=len(groups) > 1,
             yaxis=dict(title="LME Balance (€)", gridcolor="rgba(255,255,255,0.04)", tickfont=dict(color=SLATE)),
             xaxis=dict(title="", tickfont=dict(color=SLATE)),
             legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                         font=dict(color=ICE, size=10)))
        st.plotly_chart(figB1, use_container_width=True)

with rowB2:
    with st.container(border=True):
        sec("💰","Sales vs Valuation by Fixation")
        melted = view_fix.groupby("Fixation")[["Sales_Value","Final_Value"]].sum().reset_index()
        figB2 = go.Figure()
        figB2.add_trace(go.Bar(name="Sales", y=melted["Fixation"], x=melted["Sales_Value"],
                                orientation="h", marker_color=COPPER))
        figB2.add_trace(go.Bar(name="Stock+Purchase", y=melted["Fixation"], x=melted["Final_Value"],
                                orientation="h", marker_color=NAVY_LT))
        alay(figB2, barmode="group",
             xaxis=dict(title="Value (€)", gridcolor="rgba(255,255,255,0.04)", tickfont=dict(color=SLATE)),
             yaxis=dict(title="", tickfont=dict(color=SLATE)),
             legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                         font=dict(color=ICE, size=10)))
        st.plotly_chart(figB2, use_container_width=True)

# ══════════════════════ ROW C — ENTITY x MONTH COMPARISON ══════════════════════
if len(sel_e) > 1 or len(sel_m) > 1:
    with st.container(border=True):
        sec("🌍","Balance by Entity & Month")
        figC = px.bar(view_tot, x="Month", y="LME_Balance_Eur", color="Entity",
                      barmode="group", text_auto=",.0f",
                      category_orders={"Month": month_order}, color_discrete_map=ENT_COLOR)
        figC.add_hline(y=0, line_color="rgba(255,255,255,0.15)")
        figC.update_traces(textfont=dict(size=10, color=ICE), textposition="outside")
        alay(figC, showlegend=True,
             yaxis=dict(title="LME Balance (€)", gridcolor="rgba(255,255,255,0.04)", tickfont=dict(color=SLATE)),
             xaxis=dict(title="", tickfont=dict(color=SLATE)),
             legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                         font=dict(color=ICE, size=10)))
        st.plotly_chart(figC, use_container_width=True)

# ══════════════════════ INSIGHTS ══════════════════════
with st.container(border=True):
    sec("🧠","Result Interpretation")
    if len(groups) == 1:
        for line in generate_balance_insights(view_fix, view_tot):
            st.markdown(f"- {line}")
    else:
        for g in groups:
            with st.expander(f"📌 {g}", expanded=False):
                sub_fix = view_fix[view_fix["Group"] == g]
                sub_tot = view_tot[view_tot["Group"] == g]
                for line in generate_balance_insights(sub_fix, sub_tot):
                    st.markdown(f"- {line}")

# ══════════════════════ FULL DATA TABLE ══════════════════════
with st.expander("📋 Full LME Balance Table", expanded=False):
    disp_cols = ["Entity","Month","Fixation","Qty_Sold_T","LME_Sales","Sales_Value",
                 "Qty_Stock_T","LME_Stock","Stock_Value","Qty_Purchase_T","LME_Purchase",
                 "Purchase_Value","Needs_Exceed_T","Last_QTY_T","LME_Final","Final_Value",
                 "LME_Balance_Eur"]
    disp = view[disp_cols].reset_index(drop=True)
    disp.columns = ["Entity","Month","Fixation","Qty Sold (T)","LME Sales (€/kg)","Sales Value (€)",
                    "Qty Stock (T)","LME Stock (€/kg)","Stock Value (€)","Qty Purchase (T)",
                    "LME Purchase (€/kg)","Purchase Value (€)","Needs(+)/Exceed(-) (T)",
                    "Final Qty (T)","LME Final (€/kg)","Final Value (€)","LME Balance (€)"]

    qty_cols = ["Qty Sold (T)","Qty Stock (T)","Qty Purchase (T)","Needs(+)/Exceed(-) (T)","Final Qty (T)"]
    lme_cols = ["LME Sales (€/kg)","LME Stock (€/kg)","LME Purchase (€/kg)","LME Final (€/kg)"]
    eur_cols = ["Sales Value (€)","Stock Value (€)","Purchase Value (€)","Final Value (€)","LME Balance (€)"]
    fmt = {c:"{:,.2f}" for c in qty_cols}
    fmt.update({c:"{:.4f}" for c in lme_cols})
    fmt.update({c:"€{:,.0f}" for c in eur_cols})

    st.dataframe(
        disp.style.format(fmt)
            .set_properties(**{"background-color":NAVY_MD,"color":ICE})
            .map(lambda v:"color:#10b981;font-weight:700" if isinstance(v,(int,float)) and v>0
                 else ("color:#f43f5e;font-weight:700" if isinstance(v,(int,float)) and v<0 else ""),
                 subset=["LME Balance (€)"])
            .map(lambda v:"font-weight:700;color:#d4a97a" if str(v).strip().upper()=="TOTAL" else "",
                 subset=["Fixation"]),
        use_container_width=True, hide_index=True, height=380
    )

    st.download_button(
        "⬇️ Download consolidated balance (CSV)",
        data=view[disp_cols].to_csv(index=False).encode("utf-8"),
        file_name="lme_balance_consolidated.csv", mime="text/csv"
    )

st.markdown(f"""<div style="text-align:center;color:#1a2e4a;font-size:0.72rem;
  margin-top:48px;padding:16px;border-top:1px solid rgba(139,94,60,0.12);">
  Balance LME &nbsp;·&nbsp; COFICAB Kenitra &amp; COFICAB Maroc &nbsp;·&nbsp;
</div>""", unsafe_allow_html=True)
