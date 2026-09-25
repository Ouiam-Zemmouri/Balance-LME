import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import glob
import os
import re
import openpyxl

# ===== PASSWORD PROTECTION =====
PASSWORD = "Finance26"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.set_page_config(page_title="Balance LME", page_icon="⚖️", layout="centered")
    st.markdown("""
    <style>
    :root, .stApp{
      --background-color:#f4f6fb !important;
      --secondary-background-color:#ffffff !important;
      --text-color:#16264a !important;
      --primary-color:#1e3a6d !important;
    }
    .main,[data-testid="stAppViewContainer"]{background:#f4f6fb;}
    input{color:#16264a !important;background:#ffffff !important;}
    </style>
    """, unsafe_allow_html=True)
    st.markdown("<br><br>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1,1.2,1])
    with c2:
        st.markdown("""<div style="text-align:left;margin-bottom:10px;">
          <div style="font-size:1.2rem;font-weight:800;color:#16264a;">⚖️ Balance LME</div>
          <div style="font-size:0.8rem;color:#5b6478;margin-top:2px;">COFICAB Kenitra · COFICAB Maroc</div>
        </div>""", unsafe_allow_html=True)
        password = st.text_input("Password", type="password")
        if st.button("Login", use_container_width=True):
            if password == PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password")
    st.stop()

st.set_page_config(page_title="Balance LME", page_icon="⚖️",
                   layout="wide", initial_sidebar_state="expanded")

# ═══════════════════ LIGHT CORPORATE THEME ═══════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

/* Force light theme at the root — fixes native widgets (multiselect, inputs, sidebar)
   that otherwise inherit Streamlit's dark default regardless of our custom CSS below */
:root, .stApp{
  --background-color:#f4f6fb !important;
  --secondary-background-color:#ffffff !important;
  --text-color:#16264a !important;
  --primary-color:#1e3a6d !important;
}
html,body,[class*="css"]{font-family:'Inter',sans-serif;color:#16264a;}
.main,[data-testid="stAppViewContainer"]{background:#f4f6fb;}
[data-testid="stHeader"]{background:transparent;}
.block-container{padding-top:1.4rem;padding-bottom:2rem;max-width:1400px;}

/* Sidebar — force white background + dark text on every element inside it */
[data-testid="stSidebar"]{background:#ffffff !important;border-right:1px solid #e6e9f2;}
[data-testid="stSidebar"] *{color:#16264a !important;}
[data-testid="stSidebar"] hr{border-color:#e6e9f2 !important;}

/* Safety net: force dark text everywhere in the main content area — inline styles
   (KPI values, badges, etc.) still win since they carry higher CSS specificity */
[data-testid="stAppViewContainer"] *{color:#16264a;}

/* Expanders (native widget headers) were rendering white-on-white — force dark text */
[data-testid="stExpander"] summary,
[data-testid="stExpander"] summary *,
[data-testid="stExpander"] p{color:#16264a !important;}
[data-testid="stExpander"]{background:#ffffff !important;border:1px solid #e9edf5 !important;}
[data-testid="stExpander"] summary,
[data-testid="stExpander"] details,
[data-testid="stExpander"] details[open],
[data-testid="stExpander"] details[open] summary,
[data-testid="stExpander"] summary:hover,
[data-testid="stExpander"] summary:focus,
[data-testid="stExpander"] summary:active{
  background:#ffffff !important;
  background-color:#ffffff !important;
}
[data-testid="stExpanderDetails"]{background:#ffffff !important;}

/* Multiselect / select widgets — force white control + dropdown backgrounds everywhere */
div[data-baseweb="select"] > div{background:#ffffff !important;border-color:#c7d7f2 !important;}
div[data-baseweb="popover"]{background:#ffffff !important;}
ul[role="listbox"]{background:#ffffff !important;}
li[role="option"]{background:#ffffff !important;color:#16264a !important;}
li[role="option"]:hover{background:#eaf0fb !important;}
input{color:#16264a !important;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
.main,[data-testid="stAppViewContainer"]{background:#f4f6fb;}
[data-testid="stHeader"]{background:transparent;}
.block-container{padding-top:1.4rem;padding-bottom:2rem;max-width:1400px;}

/* Sidebar */
[data-testid="stSidebar"]{background:#ffffff;border-right:1px solid #e6e9f2;}
[data-testid="stSidebar"] .stCaption{color:#8993a8;}

/* Header banner */
.title-banner{
  background:linear-gradient(120deg,#16264a 0%,#1e3a6d 100%);
  border-radius:16px;padding:26px 36px;margin-bottom:22px;
  display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;
  box-shadow:0 8px 28px rgba(20,35,70,0.18);
}
.title-banner h1, .title-banner .banner-title{
  font-size:1.65rem;font-weight:800;letter-spacing:1px;margin:0;color:#ffffff;
}
/* The broad safety-net rule above repaints every descendant dark. Streamlit wraps heading
   text in a child element, so the banner needs its colors restated on descendants too. */
.title-banner .banner-title, .title-banner .banner-title *,
.title-banner h1, .title-banner h1 *{color:#ffffff !important;}
.title-banner .banner-sub, .title-banner .banner-sub *,
.title-banner p, .title-banner p *{color:#a8c0e8 !important;}
.title-banner .banner-sub{font-size:0.82rem;margin:4px 0 0 0;letter-spacing:0.5px;}
.title-banner .badge, .title-banner .badge *{color:#e8eefc !important;}
.title-banner p{color:#a8c0e8;font-size:0.82rem;margin:4px 0 0 0;letter-spacing:0.5px;}
.badge-strip{display:flex;gap:10px;flex-wrap:wrap;}
.badge{
  background:rgba(255,255,255,0.1);border:1px solid rgba(255,255,255,0.18);
  border-radius:20px;padding:6px 16px;color:#e8eefc;font-size:0.76rem;font-weight:600;
}

/* Copper impact headline */
.impact-banner{
  border-radius:16px;padding:22px 30px;margin-bottom:18px;
  display:flex;align-items:center;gap:22px;flex-wrap:wrap;
  box-shadow:0 6px 22px rgba(20,35,70,0.10);
}
.impact-banner .impact-icon{font-size:2.4rem;line-height:1;}
.impact-banner .impact-text{flex:1;min-width:260px;}
.impact-banner .impact-headline{font-size:1.15rem;font-weight:800;line-height:1.35;}
.impact-banner .impact-headline .impact-amount{font-size:1.4rem;}
.impact-banner .impact-sub{font-size:0.82rem;margin-top:4px;opacity:0.85;}

/* KPI cards */
.kpi-card{
  background:#ffffff;border-radius:14px;padding:16px 18px;position:relative;overflow:hidden;
  border:1px solid #e9edf5;
  box-shadow:0 2px 10px rgba(20,35,70,0.05);
  height:112px;display:flex;flex-direction:column;justify-content:flex-start;
  transition:all 0.15s ease;
}
.kpi-card:hover{box-shadow:0 6px 20px rgba(20,35,70,0.10);transform:translateY(-1px);}
.kpi-top{display:flex;align-items:center;gap:8px;margin-bottom:10px;}
.kpi-icon{
  width:26px;height:26px;border-radius:8px;display:flex;align-items:center;justify-content:center;
  font-size:0.82rem;flex-shrink:0;
}
.kpi-label{color:#8993a8;font-size:0.66rem;font-weight:700;text-transform:uppercase;letter-spacing:1.1px;}
.kpi-value{color:#16264a;font-size:1.42rem;font-weight:800;line-height:1.1;}
.kpi-sub{color:#a3abbd;font-size:0.7rem;margin-top:3px;font-weight:500;}
.kpi-spark{position:absolute;right:0;bottom:0;width:96px;height:34px;opacity:0.9;}

/* Section headers */
.section-header{
  color:#16264a;font-size:0.92rem;font-weight:700;
  margin:0 0 14px 0;display:flex;align-items:center;gap:8px;
}
.section-sub{color:#8993a8;font-size:0.76rem;margin:-10px 0 14px 0;}

/* Chart / content cards */
div[data-testid="stVerticalBlockBorderWrapper"]{
  background:#ffffff;border-radius:14px;border:1px solid #e9edf5 !important;
  box-shadow:0 2px 10px rgba(20,35,70,0.04);
}

/* Equal-height cards: when two bordered cards sit side by side in a row of
   columns, stretch both to the height of the tallest one so their bottom
   edges line up instead of one card trailing off shorter than the other. */
div[data-testid="stHorizontalBlock"]{align-items:stretch;}
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]{display:flex;}
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] > div{
  display:flex;flex-direction:column;width:100%;
}
div[data-testid="stHorizontalBlock"] div[data-testid="stVerticalBlockBorderWrapper"]{
  flex:1;display:flex;flex-direction:column;
}
div[data-testid="stHorizontalBlock"] div[data-testid="stVerticalBlockBorderWrapper"] > div{
  flex:1;display:flex;flex-direction:column;
}

/* Filter pills (multiselect tags) — force brand colors regardless of Streamlit's internal DOM */
span[data-baseweb="tag"], div[data-baseweb="tag"],
[data-baseweb="tag"], [data-baseweb="tag"] *,
[data-testid="stMultiSelect"] [data-baseweb="tag"]{
  background-color:#eaf0fb !important;background:#eaf0fb !important;
  border:1px solid #c7d7f2 !important;color:#1e3a6d !important;
  border-radius:8px !important;font-weight:600 !important;
}
[data-baseweb="tag"] svg{fill:#1e3a6d !important;}
.filter-label{color:#5b6478;font-size:0.68rem;font-weight:700;text-transform:uppercase;
  letter-spacing:1.5px;margin:14px 0 4px 0;}

/* Month dropdown (popover) — light styling for the trigger and its floating panel */
[data-testid="stPopover"] > button, [data-testid="stPopover"] > div > button{
  background:#ffffff !important;border:1px solid #c7d7f2 !important;border-radius:9px !important;
}
[data-testid="stPopover"] button, [data-testid="stPopover"] button *{color:#16264a !important;font-weight:600;}
[data-testid="stPopoverBody"]{background:#ffffff !important;border:1px solid #e9edf5 !important;}
[data-testid="stPopoverBody"] label, [data-testid="stPopoverBody"] p,
[data-testid="stPopoverBody"] span{color:#16264a !important;}
[data-testid="stPopoverBody"] .stButton>button, [data-testid="stPopoverBody"] .stButton>button *{color:#ffffff !important;}

/* Sidebar logo header */
.sidebar-header{padding:4px 0 16px 0;}
.sidebar-header-sub{color:#5b6478 !important;font-size:0.76rem;margin-top:8px;}

/* Sidebar coverage gauge */
.coverage-card{
  background:linear-gradient(145deg,#f8f9fc 0%,#eef1f8 100%);
  border:1px solid #e9edf5;border-radius:14px;padding:18px 14px;text-align:center;
}
.coverage-ring{position:relative;width:110px;height:110px;margin:0 auto;}
.coverage-ring svg{transform:rotate(-90deg);}
.coverage-ring .ring-bg{fill:none;stroke:#e9edf5;stroke-width:9;}
.coverage-ring .ring-fill{
  fill:none;stroke:url(#coverageGradient);stroke-width:9;stroke-linecap:round;
  stroke-dasharray:301.6;
  animation:ringFill 1.4s cubic-bezier(0.65,0,0.35,1) forwards;
}
@keyframes ringFill{from{stroke-dashoffset:301.6;}to{stroke-dashoffset:var(--offset);}}
.coverage-center{
  position:absolute;top:0;left:0;width:100%;height:100%;
  display:flex;flex-direction:column;align-items:center;justify-content:center;
}
.coverage-value{font-size:1.3rem;font-weight:800;color:#16264a !important;line-height:1;}
.coverage-label{font-size:0.62rem;color:#8993a8 !important;text-transform:uppercase;letter-spacing:1px;margin-top:2px;}
.coverage-title{font-size:0.78rem;font-weight:700;color:#16264a !important;margin-bottom:10px;}
.coverage-stats{display:flex;justify-content:center;gap:18px;margin-top:12px;}
.coverage-stat{text-align:center;}
.coverage-stat-num{font-size:1rem;font-weight:800;color:#16264a !important;}
.coverage-stat-lbl{font-size:0.62rem;color:#8993a8 !important;text-transform:uppercase;letter-spacing:0.5px;}

/* Buttons */
.stButton>button, .stButton>button *{
  background:#1e3a6d;color:#ffffff !important;border-radius:9px;border:none;font-weight:600;
}
.stButton>button:hover{background:#16264a;color:#ffffff !important;}
[data-testid="stSidebar"] .stButton>button,
[data-testid="stSidebar"] .stButton>button *{color:#ffffff !important;}

/* Expander */
.streamlit-expanderHeader{background:#ffffff;border-radius:10px;font-weight:600;color:#16264a;}

/* Dataframe */
.stDataFrame{border-radius:12px;overflow:hidden;border:1px solid #e9edf5;}

::-webkit-scrollbar{width:6px;height:6px;}
::-webkit-scrollbar-track{background:#f4f6fb;}
::-webkit-scrollbar-thumb{background:#c7d1e3;border-radius:6px;}
</style>
""", unsafe_allow_html=True)

# ── PALETTE (light theme, corporate) ──
NAVY    = "#16264a"   # deep navy — headers, primary text
NAVY_MD = "#1e3a6d"   # medium navy — primary accent / bars
NAVY_LT = "#3d6fc4"   # lighter blue — secondary series
COPPER  = "#c2703d"   # brand copper — COFICAB accent
GOLD    = "#c9932e"   # amber accent
TEAL    = "#0d9488"   # favorable / positive
ROSE    = "#e11d48"   # unfavorable / negative
SLATE   = "#5b6478"   # muted axis / caption text
INK     = "#16264a"   # dark text on white
WHITE   = "#ffffff"

PALETTE = [NAVY_MD, COPPER, TEAL, GOLD, NAVY_LT, "#7c5cbf"]

LAY = dict(
    template="plotly_white",
    paper_bgcolor="#ffffff", plot_bgcolor="#ffffff",
    font=dict(color=SLATE, family="Inter", size=12),
    margin=dict(t=30, b=40, l=55, r=20),
    legend=dict(bgcolor="rgba(255,255,255,0.95)", bordercolor="#e9edf5",
                borderwidth=1, font=dict(color=INK, size=11)),
    xaxis=dict(gridcolor="#f0f2f8", zeroline=False, tickfont=dict(color=SLATE, size=11)),
    yaxis=dict(gridcolor="#f0f2f8", zeroline=False, tickfont=dict(color=SLATE, size=11)),
    hoverlabel=dict(bgcolor="#ffffff", bordercolor="#e9edf5", font=dict(color=INK, size=12)),
)

import io
import base64
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

def sparkline_b64(values, color):
    """Tiny transparent area-sparkline as a base64 PNG, for embedding inside a KPI card."""
    if values is None or len(values) < 2:
        return None
    fig, ax = plt.subplots(figsize=(2.0, 0.5), dpi=150)
    ax.plot(values, color=color, linewidth=2)
    ax.fill_between(range(len(values)), values, min(values), color=color, alpha=0.12)
    ax.axis("off")
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", transparent=True)
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode()

def kpi(col, icon, label, val, color=NAVY_MD, sub=None, spark=None):
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    spark_html = ""
    if spark:
        b64 = sparkline_b64(spark, color)
        if b64:
            spark_html = f'<img class="kpi-spark" src="data:image/png;base64,{b64}"/>'
    html = (
        f'<div class="kpi-card">'
        f'<div class="kpi-top">'
        f'<div class="kpi-icon" style="background:{color}1a;color:{color};">{icon}</div>'
        f'<div class="kpi-label">{label}</div>'
        f'</div>'
        f'<div class="kpi-value">{val}</div>'
        f'{sub_html}{spark_html}'
        f'</div>'
    )
    col.markdown(html, unsafe_allow_html=True)

def fmt_compact(value):
    """Format a euro amount compactly, e.g. 1990000 -> '1.99 M', 374400 -> '374.4 K'."""
    sign = "-" if value < 0 else ""
    v = abs(value)
    if v >= 1_000_000:
        return f"{sign}{v/1_000_000:.2f} M"
    if v >= 1_000:
        return f"{sign}{v/1_000:.1f} K"
    return f"{sign}{v:,.0f}"

def sec(icon, title, sub=None):
    st.markdown(f'<div class="section-header">{icon}&nbsp; {title}</div>', unsafe_allow_html=True)
    if sub:
        st.markdown(f'<div class="section-sub">{sub}</div>', unsafe_allow_html=True)

def alay(fig, **kw):
    fig.update_layout(**{**LAY, **kw}); return fig

def entity_color_map(entities):
    return {e: PALETTE[i % len(PALETTE)] for i, e in enumerate(entities)}

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

def find_realloc_sources(view_fix):
    """For each row that needed reallocation (Allocated_QTE > 0), identify which other
    fixation (same entity + month) actually supplied it, by matching the reallocation
    price (LME_Realloc) against that fixation's own Stock/Purchase LME price."""
    sources = {}  # deficit fixation -> set of source fixation names
    tol = 1e-3
    for (entity, month_key), g in view_fix.groupby(["Entity", "MonthKey"]):
        for _, row in g.iterrows():
            realloc_price = row.get("LME_Realloc")
            alloc_qty = row.get("Allocated_QTE")
            if pd.isna(realloc_price) or not alloc_qty or alloc_qty <= 0:
                continue
            for _, cand in g[g["Fixation"] != row["Fixation"]].iterrows():
                for price_col in ["LME_Stock", "LME_Purchase"]:
                    cp = cand.get(price_col)
                    if pd.notna(cp) and abs(cp - realloc_price) < tol:
                        sources.setdefault(row["Fixation"], set()).add(cand["Fixation"])
                        break
    return sources


# ── SIDEBAR ──
LOGO_CANDIDATES = [
    "coficab_logo.png", "coficab_logo.PNG", "Coficab_logo.png",
    "COFICAB.png", "Coficab.png", "coficab.png", "COFICAB.PNG",
    "logo.png", "Logo.png", "LOGO.png",
    "coficab_logo.jpg", "COFICAB.jpg", "coficab.jpg", "logo.jpg",
]
LOGO_PATH = next((p for p in LOGO_CANDIDATES if os.path.exists(p)), None)
if LOGO_PATH is None:
    # fall back to a case-insensitive scan of the repo root for anything containing "coficab" or "logo"
    for f in glob.glob("*"):
        if os.path.isfile(f) and f.lower().endswith((".png", ".jpg", ".jpeg")) and \
           ("coficab" in f.lower() or "logo" in f.lower()):
            LOGO_PATH = f
            break

with st.sidebar:
    st.markdown('<div class="sidebar-header">', unsafe_allow_html=True)
    if LOGO_PATH:
        st.image(LOGO_PATH, width=180)
    else:
        st.markdown('<div style="font-size:1.05rem;font-weight:800;color:#16264a;">⚖️ COFICAB</div>',
                     unsafe_allow_html=True)
    st.markdown('<div class="sidebar-header-sub">Balance LME · Kenitra &amp; Maroc</div></div>',
                 unsafe_allow_html=True)
    if st.button("🔄 Refresh Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.markdown("---")

bal_all, errors = load_all_balance_files()
for e in errors:
    st.warning(e)

if bal_all.empty:
    st.markdown("""<div class="title-banner">
      <div><div class="banner-title">⚖️ Balance LME</div><div class="banner-sub">FIFO Method · Sales vs Stock &amp; Purchase Valuation</div></div>
    </div>""", unsafe_allow_html=True)
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
    st.markdown('<p class="filter-label">🏭 Entity</p>', unsafe_allow_html=True)
    ent_opts = sorted(bal_all["Entity"].unique())
    sel_e = st.pills("", ent_opts, selection_mode="multi", default=ent_opts,
                      key="bal_ent", label_visibility="collapsed")

    st.markdown('<p class="filter-label">📅 Month</p>', unsafe_allow_html=True)
    month_map = bal_all[["MonthKey","Month"]].drop_duplicates().sort_values("MonthKey")
    month_opts = month_map["Month"].tolist()
    # Compact dropdown (popover with checkboxes) instead of always-visible pills
    for m in month_opts:
        st.session_state.setdefault(f"mchk_{m}", True)

    def _set_all_months(value):
        for m in month_opts:
            st.session_state[f"mchk_{m}"] = value

    sel_m = [m for m in month_opts if st.session_state.get(f"mchk_{m}", True)]
    if len(sel_m) == len(month_opts):
        month_label_btn = f"All months ({len(month_opts)})"
    elif len(sel_m) == 0:
        month_label_btn = "No month selected"
    elif len(sel_m) <= 2:
        month_label_btn = ", ".join(sel_m)
    else:
        month_label_btn = f"{len(sel_m)} of {len(month_opts)} months"

    with st.popover(month_label_btn, use_container_width=True):
        b1, b2 = st.columns(2)
        b1.button("All", key="mon_all", use_container_width=True,
                  on_click=_set_all_months, args=(True,))
        b2.button("None", key="mon_none", use_container_width=True,
                  on_click=_set_all_months, args=(False,))
        for m in month_opts:
            st.checkbox(m, key=f"mchk_{m}")

    st.markdown("---")
    n_files = bal_all['SourceFile'].nunique()
    n_ent   = bal_all['Entity'].nunique()
    n_mon   = bal_all['MonthKey'].nunique()
    latest_month = bal_all.loc[bal_all["MonthKey"].idxmax(), "Month"]
    cov_pct = min(n_mon / 12 * 100, 100)
    circumference = 301.6
    cov_offset = circumference * (1 - cov_pct / 100)
    st.markdown(f"""<div class="coverage-card">
      <div class="coverage-title">📊 Data Coverage</div>
      <div class="coverage-ring">
        <svg width="110" height="110" viewBox="0 0 110 110">
          <defs>
            <linearGradient id="coverageGradient" x1="0%" y1="0%" x2="100%" y2="100%">
              <stop offset="0%" stop-color="#c2703d"/>
              <stop offset="100%" stop-color="#1e3a6d"/>
            </linearGradient>
          </defs>
          <circle class="ring-bg" cx="55" cy="55" r="48"/>
          <circle class="ring-fill" cx="55" cy="55" r="48" style="--offset:{cov_offset}px;"/>
        </svg>
        <div class="coverage-center">
          <div class="coverage-value">{n_mon}/12</div>
          <div class="coverage-label">months</div>
        </div>
      </div>
      <div class="coverage-stats">
        <div class="coverage-stat"><div class="coverage-stat-num">{latest_month}</div><div class="coverage-stat-lbl">Latest Month</div></div>
        <div class="coverage-stat"><div class="coverage-stat-num">{n_ent}</div><div class="coverage-stat-lbl">Entities</div></div>
      </div>
    </div>""", unsafe_allow_html=True)

view = bal_all[bal_all["Entity"].isin(sel_e) & bal_all["Month"].isin(sel_m)].copy()
view_fix = view[view["Fixation"].str.upper() != "TOTAL"].copy()
view_tot = view[view["Fixation"].str.upper() == "TOTAL"].copy()

# ── HEADER BANNER ──
st.markdown(f"""<div class="title-banner">
  <div>
    <div class="banner-title">⚖️ Balance LME</div>
    <div class="banner-sub">FIFO Method · Sales vs Stock &amp; Purchase Valuation</div>
  </div>
  <div class="badge-strip">
    <div class="badge">🏭 {', '.join(sel_e) if len(sel_e)<=2 else f'{len(sel_e)} entities'}</div>
    <div class="badge">📅 {', '.join(sel_m) if len(sel_m)<=3 else f'{len(sel_m)} months'}</div>
  </div>
</div>""", unsafe_allow_html=True)

if view.empty:
    st.warning("⚠️ No data matches the selected Entity / Month filters.")
    st.stop()

groups = sorted(view["Group"].unique())
month_order = month_map[month_map["Month"].isin(sel_m)]["Month"].tolist()

# ══════════════════════ METHODOLOGY ══════════════════════
with st.expander("📖 Methodology — What the LME Balance Is & How It's Calculated", expanded=True):
    st.markdown(f"""
<div style="color:{INK};font-size:0.92rem;line-height:1.65;">

<p><strong>What is the LME Balance?</strong><br>
COFICAB sells cable and wire products whose copper content is invoiced against a specific
London Metal Exchange (LME) fixing. The physical copper actually consumed to manufacture what
was sold, however, was purchased and stocked earlier — often referencing a <em>different</em>
LME fixing. The <strong>LME Balance</strong> measures the euro gain or loss created by this
timing mismatch between the price used to sell and the price of the copper actually consumed,
computed separately for each fixation category (3M-1, 3M-2, M-1…), since each one carries its
own cost basis.</p>

<p><strong>The FIFO matching logic — how a sale is costed, step by step:</strong><br>
For each fixation, the quantity sold during the month must be matched, in strict
First-In-First-Out order, against the following sources — in this sequence:</p>

<ol style="margin-top:6px;">
<li><strong>Stock carried over from the previous month</strong>, on the <em>same</em> fixation,
is used first.</li>
<li>If that stock does not cover the full quantity sold, the shortfall is covered by
<strong>purchases consumed during the current month</strong>, still on the same fixation.</li>
<li>If a shortfall still remains after using the fixation's own stock and purchases, the
remaining quantity is <strong>reallocated from the next fixation in the FIFO sequence</strong> —
drawing from whatever surplus stock or purchase quantity that other fixation still has available.</li>
</ol>

<p>Each portion of the final quantity sold is therefore valorized at the LME fixing price of its
own source (own stock, own purchases, or reallocated stock/purchases from another fixation),
giving a single blended cost basis for the month.</p>

</div>
""", unsafe_allow_html=True)

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

    # ── FIFO cascade diagram ──
    st.markdown(f"""
<svg viewBox="0 0 900 660" xmlns="http://www.w3.org/2000/svg" style="width:100%;height:auto;">
  <defs>
    <marker id="arrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
      <path d="M0,0 L10,5 L0,10 z" fill="{SLATE}"/>
    </marker>
  </defs>

  <!-- Step 1 -->
  <rect x="130" y="10" width="420" height="66" rx="12" fill="{NAVY_MD}"/>
  <text x="340" y="38" text-anchor="middle" fill="#ffffff" font-size="15" font-weight="700" font-family="Inter">① Monthly Sales — Fixation F</text>
  <text x="340" y="60" text-anchor="middle" fill="#c9d4ea" font-size="12" font-family="Inter">Quantity to cover: Qty Sold (T)</text>

  <line x1="340" y1="76" x2="340" y2="118" stroke="{SLATE}" stroke-width="2" marker-end="url(#arrow)"/>
  <text x="352" y="102" fill="{SLATE}" font-size="11" font-family="Inter">covered first by ↓</text>

  <!-- Step 2 -->
  <rect x="130" y="120" width="420" height="66" rx="12" fill="{NAVY_LT}"/>
  <text x="340" y="148" text-anchor="middle" fill="#ffffff" font-size="15" font-weight="700" font-family="Inter">② Stock — Previous Month, Same Fixation</text>
  <text x="340" y="170" text-anchor="middle" fill="#e3ebfb" font-size="12" font-family="Inter">Available: Qty Stock (T)</text>

  <line x1="550" y1="153" x2="670" y2="153" stroke="{TEAL}" stroke-width="2" marker-end="url(#arrow)"/>
  <text x="608" y="145" text-anchor="middle" fill="{TEAL}" font-size="11" font-family="Inter">if sufficient</text>
  <rect x="672" y="120" width="180" height="66" rx="12" fill="#e6f7f4" stroke="{TEAL}" stroke-width="1.5"/>
  <text x="762" y="148" text-anchor="middle" fill="{TEAL}" font-size="12" font-weight="700" font-family="Inter">✅ Fully covered</text>
  <text x="762" y="166" text-anchor="middle" fill="{TEAL}" font-size="10.5" font-family="Inter">valued at Stock LME price</text>

  <line x1="340" y1="186" x2="340" y2="228" stroke="{SLATE}" stroke-width="2" marker-end="url(#arrow)"/>
  <text x="352" y="212" fill="{SLATE}" font-size="11" font-family="Inter">if shortfall remains ↓</text>

  <!-- Step 3 -->
  <rect x="130" y="230" width="420" height="66" rx="12" fill="{NAVY_LT}"/>
  <text x="340" y="258" text-anchor="middle" fill="#ffffff" font-size="15" font-weight="700" font-family="Inter">③ Purchases Consumed This Month, Same Fixation</text>
  <text x="340" y="280" text-anchor="middle" fill="#e3ebfb" font-size="12" font-family="Inter">Available: Qty Purchase (T)</text>

  <line x1="550" y1="263" x2="670" y2="263" stroke="{TEAL}" stroke-width="2" marker-end="url(#arrow)"/>
  <text x="608" y="255" text-anchor="middle" fill="{TEAL}" font-size="11" font-family="Inter">if sufficient</text>
  <rect x="672" y="230" width="180" height="66" rx="12" fill="#e6f7f4" stroke="{TEAL}" stroke-width="1.5"/>
  <text x="762" y="256" text-anchor="middle" fill="{TEAL}" font-size="12" font-weight="700" font-family="Inter">✅ Covered by</text>
  <text x="762" y="272" text-anchor="middle" fill="{TEAL}" font-size="10.5" font-family="Inter">Stock + Purchases</text>

  <line x1="340" y1="296" x2="340" y2="338" stroke="{SLATE}" stroke-width="2" marker-end="url(#arrow)"/>
  <text x="352" y="322" fill="{SLATE}" font-size="11" font-family="Inter">if still short ↓</text>

  <!-- Step 4 — Reallocation (highlighted) -->
  <rect x="110" y="340" width="460" height="86" rx="12" fill="{GOLD}"/>
  <text x="340" y="370" text-anchor="middle" fill="#ffffff" font-size="15" font-weight="800" font-family="Inter">④ Reallocation — Next Fixation (FIFO order)</text>
  <text x="340" y="392" text-anchor="middle" fill="#fff3e0" font-size="12" font-family="Inter">Remaining quantity is pulled from the surplus</text>
  <text x="340" y="408" text-anchor="middle" fill="#fff3e0" font-size="12" font-family="Inter">stock/purchase of the next fixation in sequence</text>

  <line x1="340" y1="426" x2="340" y2="468" stroke="{SLATE}" stroke-width="2" marker-end="url(#arrow)"/>

  <!-- Step 5 — Final -->
  <rect x="90" y="470" width="500" height="86" rx="12" fill="{COPPER}"/>
  <text x="340" y="500" text-anchor="middle" fill="#ffffff" font-size="15" font-weight="800" font-family="Inter">⑤ Final Valorized Quantity &amp; LME Balance</text>
  <text x="340" y="522" text-anchor="middle" fill="#fdeee0" font-size="12" font-family="Inter">LME Balance (€) = Sales Value (€) − Valorized Cost</text>
  <text x="340" y="538" text-anchor="middle" fill="#fdeee0" font-size="12" font-family="Inter">(Stock + Purchases + Reallocated Quantity)</text>

  <!-- Legend -->
  <rect x="90" y="580" width="16" height="16" rx="4" fill="{TEAL}"/>
  <text x="114" y="593" fill="{INK}" font-size="12" font-family="Inter">Positive balance — favorable to COFICAB (sold above FIFO cost)</text>
  <rect x="90" y="608" width="16" height="16" rx="4" fill="{ROSE}"/>
  <text x="114" y="621" fill="{INK}" font-size="12" font-family="Inter">Negative balance — unfavorable (sold below FIFO cost)</text>
</svg>
""", unsafe_allow_html=True)

    st.markdown(f"""
<div style="color:{SLATE};font-size:0.85rem;margin-top:10px;padding-top:10px;border-top:1px solid #e9edf5;">
💡 <strong>Reading the "Needs(+)/Exceed(-)" column</strong>: for a given fixation, a
<strong>positive</strong> value means its own stock + purchases were insufficient — the
remainder was reallocated from the next fixation (step ④ above). A <strong>negative</strong>
value means that fixation had a surplus, which may itself be used to cover another
fixation's shortfall.
</div>
""", unsafe_allow_html=True)

# ══════════════════════ KPI ROW ══════════════════════
# Aggregates are summed from the per-fixation rows (view_fix) rather than the TOTAL row,
# so a file with a blank or shifted TOTAL line cannot silently under-count the figures.
tot_sales   = view_fix["Sales_Value"].sum()
tot_balance = view_fix["LME_Balance_Eur"].sum()
tot_qty     = view_fix["Qty_Sold_T"].sum()
bal_per_t   = tot_balance / tot_qty if tot_qty else 0
bal_pct     = (tot_balance / tot_sales * 100) if tot_sales else 0
bal_color   = TEAL if tot_balance >= 0 else ROSE

grp_bal     = view_fix.groupby("Group")["LME_Balance_Eur"].sum()
n_fav       = int((grp_bal >= 0).sum())
n_tot       = int(len(grp_bal))

monthly = view_fix.groupby("MonthKey")[["Sales_Value","LME_Balance_Eur","Qty_Sold_T"]].sum().sort_index()
spark_balance = monthly["LME_Balance_Eur"].tolist() if len(monthly) > 1 else None
spark_qty     = monthly["Qty_Sold_T"].tolist()      if len(monthly) > 1 else None
spark_pct     = ((monthly["LME_Balance_Eur"] / monthly["Sales_Value"].replace(0, pd.NA) * 100)
                 .fillna(0).tolist()) if len(monthly) > 1 else None

sub_balance = "Favorable" if tot_balance >= 0 else "Unfavorable"
sub_qty     = f"over {n_tot} entity × month"
if len(monthly) > 1:
    d_bal = monthly["LME_Balance_Eur"].iloc[-1] - monthly["LME_Balance_Eur"].iloc[-2]
    d_qty = monthly["Qty_Sold_T"].iloc[-1] - monthly["Qty_Sold_T"].iloc[-2]
    sub_balance = f"{'▲' if d_bal>=0 else '▼'} €{fmt_compact(abs(d_bal))} vs last month"
    sub_qty     = f"{'▲' if d_qty>=0 else '▼'} {abs(d_qty):,.0f} T vs last month"

# ── Sales headline (total sales amount — the balance itself is shown in the KPI cards) ──
impact_bg   = "linear-gradient(120deg,#eaf0fb 0%,#dbe6f8 100%)"
impact_txt  = NAVY_MD
impact_icon = "💶"
period_lbl  = ", ".join(sel_m) if len(sel_m) <= 3 else f"the {len(sel_m)} selected months"
top_fix_row = view_fix.groupby("Fixation")["Sales_Value"].sum().reset_index()
top_fix_row = top_fix_row.sort_values("Sales_Value", ascending=False)
top_fix = top_fix_row.iloc[0]["Fixation"] if not top_fix_row.empty else "—"
top_ent_row = view_fix.groupby("Entity")["Sales_Value"].sum().reset_index()
top_ent_row = top_ent_row.sort_values("Sales_Value", ascending=False)
top_ent = top_ent_row.iloc[0]["Entity"] if not top_ent_row.empty else "—"

st.markdown(f"""<div class="impact-banner" style="background:{impact_bg};">
  <div class="impact-icon">{impact_icon}</div>
  <div class="impact-text">
    <div class="impact-headline" style="color:{impact_txt};">
      Total sales of <span class="impact-amount">€{tot_sales:,.0f}</span> over {period_lbl}
    </div>
    <div class="impact-sub" style="color:{impact_txt};">
      Largest share from <strong>{top_fix}</strong> at <strong>{top_ent}</strong>
    </div>
  </div>
</div>""", unsafe_allow_html=True)

k1,k2,k3 = st.columns(3)
kpi(k1,"⚖️","Net LME Balance",   f"€{tot_balance:,.0f}", bal_color, sub_balance, spark_balance)
kpi(k2,"📦","Total Qty Sold",    f"{tot_qty:,.0f} T",    NAVY_MD, sub_qty, spark_qty)
kpi(k3,"％","Balance % of Sales", f"{bal_pct:+.2f}%",     bal_color,
    f"on €{fmt_compact(tot_sales)} of sales", spark_pct)



st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

# ══════════════════════ MAIN TABS ══════════════════════
tab_overview, tab_stock, tab_sales, tab_supply, tab_insights, tab_data = st.tabs(
    ["📊 Overview", "📦 Stock Analysis", "💰 Sales Analysis", "⚖️ Supply vs Sales", "🧠 Insights", "📋 Data"]
)

# ─────────────────────────── TAB: OVERVIEW ───────────────────────────
with tab_overview:
    with st.container(border=True):
        sec("🔶","Copper Price vs. LME Balance", "How movements in the copper price line up with the resulting gain/loss")
        cp_monthly = view_fix.groupby(["MonthKey","Month"]).apply(
            lambda g: pd.Series({
                "Avg_LME_Price": (g["Qty_Sold_T"] * g["LME_Sales"]).sum() / g["Qty_Sold_T"].sum() if g["Qty_Sold_T"].sum() else 0,
                "Avg_Purchase_Price": (g["Qty_Purchase_T"] * g["LME_Purchase"]).sum() / g["Qty_Purchase_T"].sum() if g["Qty_Purchase_T"].sum() else None,
                "LME_Balance_Eur": g["LME_Balance_Eur"].sum(),
            })
        ).reset_index().sort_values("MonthKey")

        if len(cp_monthly) > 1:
            figCP = make_subplots(specs=[[{"secondary_y": True}]])
            figCP.add_trace(go.Bar(
                x=cp_monthly["Month"], y=cp_monthly["LME_Balance_Eur"], name="LME Balance (€)",
                marker_color=[TEAL if v >= 0 else ROSE for v in cp_monthly["LME_Balance_Eur"]],
                opacity=0.75, hovertemplate="€%{y:,.0f}<extra></extra>"
            ), secondary_y=False)
            figCP.add_trace(go.Scatter(
                x=cp_monthly["Month"], y=cp_monthly["Avg_LME_Price"], name="Avg Sales Price (€/kg)",
                mode="lines+markers", line=dict(color=COPPER, width=3), marker=dict(size=9),
                hovertemplate="%{y:.4f} €/kg<extra></extra>"
            ), secondary_y=True)
            pur = cp_monthly.dropna(subset=["Avg_Purchase_Price"])
            if not pur.empty:
                figCP.add_trace(go.Scatter(
                    x=pur["Month"], y=pur["Avg_Purchase_Price"], name="Avg Purchase Price (€/kg)",
                    mode="lines+markers", line=dict(color=NAVY_LT, width=3, dash="dot"), marker=dict(size=9, symbol="diamond"),
                    hovertemplate="%{y:.4f} €/kg<extra></extra>"
                ), secondary_y=True)
            figCP.update_layout(**LAY)
            figCP.update_yaxes(title_text="LME Balance (€)", secondary_y=False, gridcolor="#f0f2f8")
            figCP.update_yaxes(title_text="Avg Copper Price (€/kg)", secondary_y=True, showgrid=False)
            figCP.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(figCP, use_container_width=True, theme=None)
            st.caption("Bars = net LME Balance (€, left axis) · Solid line = average sales price · Dotted line = average purchase price (€/kg, right axis)")
        else:
            cp_row = cp_monthly.iloc[0] if not cp_monthly.empty else None
            if cp_row is not None:
                st.metric("Avg Copper Price this period (€/kg)", f"{cp_row['Avg_LME_Price']:.4f}")
            st.caption("Add more monthly files to see how the copper price and the balance move together over time.")

    rowA1, rowA2 = st.columns([5,3])

    with rowA1:
        with st.container(border=True):
            sec("📅","Net LME Balance — Trend", "Monthly evolution by entity")
            trend = (view_tot.groupby(["Entity","MonthKey","Month"])["LME_Balance_Eur"]
                     .sum().reset_index().sort_values("MonthKey"))
            if trend["MonthKey"].nunique() > 1:
                figA1 = px.line(trend, x="Month", y="LME_Balance_Eur", color="Entity",
                                 markers=True, category_orders={"Month": month_order},
                                 color_discrete_map=ENT_COLOR)
                figA1.update_traces(line=dict(width=3), marker=dict(size=9))
                figA1.add_hline(y=0, line_dash="dot", line_color="#dde3f0")
                alay(figA1, showlegend=len(sel_e) > 1, height=480,
                     legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, title=None),
                     yaxis=dict(title="LME Balance (€)"), xaxis=dict(title=""))
                st.plotly_chart(figA1, use_container_width=True, theme=None)
            else:
                figA1 = px.bar(view_tot, x="Entity", y="LME_Balance_Eur", color="Entity",
                                color_discrete_map=ENT_COLOR, text_auto=",.0f")
                figA1.add_hline(y=0, line_dash="dot", line_color="#dde3f0")
                alay(figA1, showlegend=False, height=480, yaxis=dict(title="LME Balance (€)"), xaxis=dict(title=""))
                st.plotly_chart(figA1, use_container_width=True, theme=None)
                st.caption("Add more monthly files to unlock the trend view.")

    with rowA2:
        with st.container(border=True):
            sec("🥯","How Many Periods Were Favorable?", "Click a button below to see which entity × month")
            n_unfav = n_tot - n_fav
            figA2 = go.Figure(go.Pie(
                labels=["Favorable", "Unfavorable"], values=[n_fav, n_unfav], hole=0.62,
                marker=dict(colors=[TEAL, ROSE], line=dict(color="#ffffff", width=3)),
                textinfo="value", textfont=dict(color="#ffffff", size=13), sort=False
            ))
            alay(figA2, showlegend=True, height=480,
                 legend=dict(orientation="h", yanchor="bottom", y=-0.15, xanchor="center", x=0.5),
                 annotations=[dict(text=f"{n_fav}/{n_tot}", x=0.5, y=0.5,
                                    font=dict(size=17, color=bal_color, family="Inter"), showarrow=False)])
            st.plotly_chart(figA2, use_container_width=True, theme=None)

            bc1, bc2 = st.columns(2)
            if bc1.button(f"✅ Favorable ({n_fav})", use_container_width=True, key="btn_fav"):
                st.session_state["fav_filter"] = None if st.session_state.get("fav_filter") == "Favorable" else "Favorable"
            if bc2.button(f"⚠️ Unfavorable ({n_unfav})", use_container_width=True, key="btn_unfav"):
                st.session_state["fav_filter"] = None if st.session_state.get("fav_filter") == "Unfavorable" else "Unfavorable"

    sel_label = st.session_state.get("fav_filter")
    if sel_label in ("Favorable", "Unfavorable"):
        want_fav = (sel_label == "Favorable")
        matches = grp_bal[grp_bal >= 0] if want_fav else grp_bal[grp_bal < 0]
        matches = matches.sort_values(ascending=not want_fav)
        with st.container(border=True):
            icon = "✅" if want_fav else "⚠️"
            sec(icon, f"{sel_label} periods — {len(matches)} entity × month", "Click the same button again to hide this")
            detail = matches.reset_index()
            detail.columns = ["Entity × Month", "LME Balance (€)"]
            st.dataframe(
                detail.style.format({"LME Balance (€)":"€{:,.0f}"})
                    .set_properties(**{"background-color":"#ffffff","color":INK})
                    .map(lambda v:"color:#0d9488;font-weight:700" if isinstance(v,(int,float)) and v>=0
                         else "color:#e11d48;font-weight:700", subset=["LME Balance (€)"]),
                use_container_width=True, hide_index=True, height=min(38*len(detail)+40, 300)
            )

    with st.container(border=True):
        sec("🏆","Which Fixation Wins or Loses the Most?", "Ranked by net LME Balance across the current selection")
        rank_df = view_fix.groupby("Fixation")["LME_Balance_Eur"].sum().reset_index().sort_values("LME_Balance_Eur")
        rank_df["Rank"] = range(len(rank_df), 0, -1)
        figRank = go.Figure(go.Bar(
            x=rank_df["LME_Balance_Eur"], y=rank_df["Fixation"], orientation="h",
            marker_color=[TEAL if v >= 0 else ROSE for v in rank_df["LME_Balance_Eur"]],
            text=[f"€{v:,.0f}" for v in rank_df["LME_Balance_Eur"]], textposition="outside"
        ))
        figRank.add_vline(x=0, line_color="#dde3f0")
        alay(figRank, xaxis=dict(title="LME Balance (€)"), yaxis=dict(title=""))
        st.plotly_chart(figRank, use_container_width=True, theme=None)
        best_fix = rank_df.iloc[-1]
        worst_fix = rank_df.iloc[0]
        st.caption(f"🥇 **{best_fix['Fixation']}** contributes the most (€{best_fix['LME_Balance_Eur']:,.0f}) · "
                   f"📉 **{worst_fix['Fixation']}** drags the result down the most (€{worst_fix['LME_Balance_Eur']:,.0f})")

    if len(sel_e) > 1:
        with st.container(border=True):
            sec("🌍","Global View (YTD)", "Kenitra vs Maroc, fixation by fixation — who performs better on which fixation?")
            figEF = px.bar(view_fix.groupby(["Entity","Fixation"])["LME_Balance_Eur"].sum().reset_index(),
                           x="Fixation", y="LME_Balance_Eur", color="Entity",
                           barmode="group", text_auto=",.0f", color_discrete_map=ENT_COLOR)
            figEF.add_hline(y=0, line_color="#dde3f0")
            figEF.update_traces(textfont=dict(size=10, color=INK), textposition="outside")
            alay(figEF, yaxis=dict(title="LME Balance (€)"), xaxis=dict(title=""),
                 legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(figEF, use_container_width=True, theme=None)

    if len(sel_e) > 1 or len(sel_m) > 1:
        with st.container(border=True):
            sec("🌍","Monthly View", "Balance by entity & month — side-by-side comparison")
            figC = px.bar(view_tot, x="Month", y="LME_Balance_Eur", color="Entity",
                          barmode="group", text_auto=",.0f",
                          category_orders={"Month": month_order}, color_discrete_map=ENT_COLOR)
            figC.add_hline(y=0, line_color="#dde3f0")
            figC.update_traces(textfont=dict(size=10, color=INK), textposition="outside")
            alay(figC, showlegend=True, yaxis=dict(title="LME Balance (€)"), xaxis=dict(title=""),
                 legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(figC, use_container_width=True, theme=None)

# ─────────────────────────── TABS: STOCK ANALYSIS & SALES ANALYSIS ───────────────────────────
def flow_agg(df, qty_col, val_col):
    """Aggregate a flow (stock / purchase / sales) by fixation: total qty (T), total value (€)
    and quantity-weighted average LME price (€/kg = value / (qty in T × 1000))."""
    g = df.groupby("Fixation")[[qty_col, val_col]].sum().reset_index()
    g.columns = ["Fixation", "Qty", "Value"]
    g["LME"] = g["Value"] / (g["Qty"].where(g["Qty"] > 0) * 1000)
    return g

def wavg(df, qty_col, val_col):
    q, v = df[qty_col].sum(), df[val_col].sum()
    return q, v, (v / (q * 1000) if q > 0 else None)

def combo_fixation_chart(agg, qty_name, bar_color):
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=agg["Fixation"], y=agg["Qty"], name=qty_name, marker_color=bar_color, opacity=0.85,
        text=[f"{v:,.1f}" for v in agg["Qty"]], textposition="outside"), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=agg["Fixation"], y=agg["LME"], name="Avg LME (€/kg)", mode="lines+markers",
        line=dict(color=COPPER, width=3), marker=dict(size=10)), secondary_y=True)
    alay(fig, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    fig.update_yaxes(title_text="Quantity (T)", secondary_y=False, gridcolor="#f0f2f8")
    fig.update_yaxes(title_text="LME (€/kg)", secondary_y=True, showgrid=False)
    return fig

def analysis_section(icon, title, sub, df, qty_col, val_col, qty_name, color, empty_msg):
    with st.container(border=True):
        sec(icon, title, sub)
        q, v, p = wavg(df, qty_col, val_col)
        if q <= 0:
            st.info(empty_msg)
            return
        c1, c2, c3 = st.columns(3)
        kpi(c1, "📦", qty_name, f"{q:,.1f} T", color)
        kpi(c2, "🔶", "Weighted Avg LME", f"{p:.4f} €/kg", COPPER)
        kpi(c3, "💶", "Total Value", f"€{fmt_compact(v)}", NAVY_MD)
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
        agg = flow_agg(df, qty_col, val_col)
        agg = agg[agg["Qty"] > 0]
        st.plotly_chart(combo_fixation_chart(agg, qty_name, color),
                        use_container_width=True, theme=None)
        top = agg.loc[agg["Qty"].idxmax()]
        lo, hi = agg.loc[agg["LME"].idxmin()], agg.loc[agg["LME"].idxmax()]
        note = (f"**{top['Fixation']}** carries the largest quantity "
                f"({top['Qty']:,.1f} T at {top['LME']:.4f} €/kg).")
        if lo["Fixation"] != hi["Fixation"]:
            note += (f" Lowest LME: **{lo['Fixation']}** ({lo['LME']:.4f} €/kg) · "
                     f"highest LME: **{hi['Fixation']}** ({hi['LME']:.4f} €/kg).")
        st.caption(note)

def flow_evolution_chart(df, qty_col, val_col, qty_name, color):
    rows = []
    for mk, g in df.groupby("MonthKey"):
        q, v, p = wavg(g, qty_col, val_col)
        rows.append({"MonthKey": mk, "Month": g["Month"].iloc[0], "Qty": q, "LME": p})
    ev = pd.DataFrame(rows).sort_values("MonthKey")
    ev = ev[ev["Qty"] > 0]
    if ev.empty:
        return None
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=ev["Month"], y=ev["Qty"], name=qty_name, marker_color=color, opacity=0.85,
        text=[f"{v:,.0f}" for v in ev["Qty"]], textposition="outside"), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=ev["Month"], y=ev["LME"], name="Avg LME (€/kg)", mode="lines+markers",
        line=dict(color=COPPER, width=3), marker=dict(size=10)), secondary_y=True)
    alay(fig, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    fig.update_yaxes(title_text="Quantity (T)", secondary_y=False, gridcolor="#f0f2f8")
    fig.update_yaxes(title_text="LME (€/kg)", secondary_y=True, showgrid=False)
    return fig

def render_flow_tab(icon, title, sub, qty_col, val_col, qty_name, color, empty_msg, key):
    analysis_section(icon, title, sub, view_fix, qty_col, val_col, qty_name, color, empty_msg)
    if view_fix[qty_col].sum() <= 0:
        return

    if view_fix["MonthKey"].nunique() > 1:
        with st.container(border=True):
            sec("📈", f"{qty_name} & LME Price by Month", "Monthly quantity (bars) and weighted average LME (line)")
            figEV = flow_evolution_chart(view_fix, qty_col, val_col, qty_name, color)
            if figEV is not None:
                st.plotly_chart(figEV, use_container_width=True, theme=None)

    with st.container(border=True):
        sec("📋", "Detail by Fixation", "Aggregated across the current selection")
        d = flow_agg(view_fix, qty_col, val_col)
        d["Share"] = d["Qty"] / d["Qty"].sum() * 100
        d = d.sort_values("Fixation")[["Fixation", "Qty", "LME", "Value", "Share"]]
        d.columns = ["Fixation", f"{qty_name} (T)", "LME (€/kg)", "Value (€)", "Share of Qty (%)"]
        d_fmt = {f"{qty_name} (T)": "{:,.1f}", "LME (€/kg)": "{:.4f}",
                 "Value (€)": "€{:,.0f}", "Share of Qty (%)": "{:.1f}%"}
        st.dataframe(
            d.style.format(d_fmt, na_rep="—")
               .set_properties(**{"background-color": "#ffffff", "color": INK}),
            use_container_width=True, hide_index=True, height=38 * len(d) + 40)

with tab_stock:
    render_flow_tab("📦", "Stock Analysis", "Opening stock carried over from the previous month, by fixation",
                    "Qty_Stock_T", "Stock_Value", "Qty Stock", NAVY_LT,
                    "No stock recorded for the current selection.", "stock")

with tab_sales:
    render_flow_tab("💰", "Sales Analysis", "Quantities sold and LME sales price, by fixation",
                    "Qty_Sold_T", "Sales_Value", "Qty Sold", COPPER,
                    "No sales recorded for the current selection.", "sales")

def _pc(a, b):
    return (a / b * 100) if b else 0.0

def _rgba(hexc, alpha):
    h = hexc.lstrip("#")
    return f"rgba({int(h[0:2],16)},{int(h[2:4],16)},{int(h[4:6],16)},{alpha})"

def _stack_bar(parts, height=34, show_text=True):
    """100% stacked horizontal bar (pure HTML). parts = [(label, value, color), ...]"""
    total = sum(v for _, v, _ in parts if v > 0)
    if total <= 0:
        return ""
    segs = ""
    for label, v, c in parts:
        if v <= 0:
            continue
        p = v / total * 100
        txt = f"{label} {p:.0f}%" if (show_text and p >= 9) else ""
        segs += (f'<div style="width:{p:.2f}%;background:{c};color:#ffffff;display:flex;align-items:center;'
                 f'justify-content:center;font-size:0.8rem;font-weight:700;white-space:nowrap;">{txt}</div>')
    return (f'<div style="display:flex;height:{height}px;border-radius:{height//2}px;overflow:hidden;'
            f'background:#e9edf5;">{segs}</div>')

with tab_supply:
    sup = view_fix.copy()
    for c in ["Qty_Stock_T","Qty_Purchase_T","Qty_Sold_T","Stock_Value","Purchase_Value","Sales_Value"]:
        sup[c] = pd.to_numeric(sup[c], errors="coerce").fillna(0)
    sup["Supply_T"] = sup["Qty_Stock_T"] + sup["Qty_Purchase_T"]
    sup["Supply_Value"] = sup["Stock_Value"] + sup["Purchase_Value"]

    g = sup.groupby("Fixation")[["Supply_T","Supply_Value","Qty_Sold_T","Sales_Value"]].sum().reset_index()
    g["Supply_LME"] = g["Supply_Value"] / (g["Supply_T"].where(g["Supply_T"] > 0) * 1000)
    g["Sales_LME"]  = g["Sales_Value"]  / (g["Qty_Sold_T"].where(g["Qty_Sold_T"] > 0) * 1000)
    g["Gap_T"] = g["Supply_T"] - g["Qty_Sold_T"]
    FIX_ORDER2 = {"M-1": 0, "3M-1": 1, "3M-2": 2}
    g["_ord"] = g["Fixation"].map(FIX_ORDER2).fillna(99)
    g = g.sort_values("_ord").drop(columns="_ord").reset_index(drop=True)

    T_supply, T_sold2 = g["Supply_T"].sum(), g["Qty_Sold_T"].sum()
    T_supply_val, T_sales_val = g["Supply_Value"].sum(), g["Sales_Value"].sum()
    coverage = _pc(T_supply, T_sold2) if T_sold2 else 0
    gap = T_supply - T_sold2

    if T_sold2 <= 0 and T_supply <= 0:
        st.info("Not enough data to compare supply and sales for the current selection.")
    else:
        # ── Headline banner (isolated iframe, same technique as the Insights hero) ──
        cover_word = "fully covered" if coverage >= 100 else "short"
        cover_color = TEAL if coverage >= 100 else ROSE
        headline = (f'Stock &amp; purchases supplied <b>{T_supply:,.0f} T</b> against '
                    f'<b>{T_sold2:,.0f} T</b> sold — a coverage of <b>{coverage:.0f}%</b>, '
                    f'{cover_word} by {abs(gap):,.0f} T {"of surplus" if gap >= 0 else "made up through reallocation"}.')
        sub = (f'Own supply was valued at <b>{(T_supply_val/(T_supply*1000)) if T_supply else 0:.4f} €/kg</b> on average, '
               f'sold at <b>{(T_sales_val/(T_sold2*1000)) if T_sold2 else 0:.4f} €/kg</b> — '
               f'a spread of <b>{((T_sales_val/(T_sold2*1000)) - (T_supply_val/(T_supply*1000))) if T_supply and T_sold2 else 0:+.4f} €/kg</b>.')
        supply_hero = f'''<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
  html,body{{margin:0;padding:0;background:transparent;font-family:'Inter','Segoe UI',Arial,sans-serif;}}
  .card{{background:linear-gradient(120deg,{NAVY} 0%,{NAVY_MD} 100%);border-radius:18px;
    padding:26px 30px;box-shadow:0 8px 24px rgba(22,38,74,0.18);box-sizing:border-box;}}
  .line{{font-size:1.12rem;line-height:1.6;color:#ffffff;margin:0 0 8px 0;}}
  .soft{{font-size:1.0rem;line-height:1.6;color:#dbe6f8;}}
  b{{color:inherit;font-weight:800;}}
  .badge{{display:inline-block;margin-top:14px;padding:5px 14px;border-radius:999px;
    font-size:0.78rem;font-weight:700;background:{cover_color}33;color:{cover_color if cover_color!=TEAL else '#9ff0d6'};}}
</style></head>
<body><div class="card">
  <div class="line">{headline}</div>
  <div class="soft">{sub}</div>
  <div class="badge">{"✅ Self-sufficient" if coverage >= 100 else "🔁 Needs reallocation"}</div>
</div></body></html>'''
        components.html(supply_hero, height=190, scrolling=False)

        c1, c2, c3, c4 = st.columns(4)
        kpi(c1, "📥", "Total Supply", f"{T_supply:,.0f} T", NAVY_LT, f"Stock + purchases · €{fmt_compact(T_supply_val)}")
        kpi(c2, "📤", "Total Sold", f"{T_sold2:,.0f} T", COPPER, f"€{fmt_compact(T_sales_val)}")
        kpi(c3, "🎯", "Coverage", f"{coverage:.0f}%", cover_color, "Supply vs sales, by tonnage")
        kpi(c4, "⚖️", "Net Gap", f"{gap:+,.0f} T", TEAL if gap >= 0 else ROSE,
            "Surplus left unsold" if gap >= 0 else "Covered by reallocation")
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

        gv = g[(g["Supply_T"] > 0) | (g["Qty_Sold_T"] > 0)].reset_index(drop=True)

        with st.container(border=True):
            sec("🦋", "Supply vs Sales, by Fixation", "Tornado view — supply on the left, sales on the right; the wider side wins")
            figSV = go.Figure()
            figSV.add_trace(go.Bar(
                y=gv["Fixation"], x=-gv["Supply_T"], name="Supply (Stock+Purchase)", orientation="h",
                marker_color=NAVY_LT, text=[f"{v:,.0f} T" for v in gv["Supply_T"]], textposition="outside",
                hovertemplate="%{customdata:,.1f} T<extra></extra>", customdata=gv["Supply_T"]))
            figSV.add_trace(go.Bar(
                y=gv["Fixation"], x=gv["Qty_Sold_T"], name="Sold", orientation="h",
                marker_color=COPPER, text=[f"{v:,.0f} T" for v in gv["Qty_Sold_T"]], textposition="outside",
                hovertemplate="%{x:,.1f} T<extra></extra>"))
            figSV.add_vline(x=0, line_color="#c9d4ea", line_width=1.5)
            m = max(gv["Supply_T"].max(), gv["Qty_Sold_T"].max()) * 1.35 if len(gv) else 1
            alay(figSV, barmode="overlay", height=120 + 90 * len(gv),
                 legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                 xaxis=dict(title="Quantity (T)  ←  Supply   |   Sold  →", range=[-m, m],
                             tickvals=[-m*0.66, -m*0.33, 0, m*0.33, m*0.66],
                             ticktext=[f"{m*0.66:,.0f}", f"{m*0.33:,.0f}", "0", f"{m*0.33:,.0f}", f"{m*0.66:,.0f}"]),
                 yaxis=dict(title=""))
            st.plotly_chart(figSV, use_container_width=True, theme=None)

        with st.container(border=True):
            sec("🎯", "Price: Supply Cost vs Sales Price", "Dumbbell view — each line is the margin captured on that fixation (€/kg)")
            figP = go.Figure()
            for _, r in gv.iterrows():
                if pd.isna(r["Supply_LME"]) or pd.isna(r["Sales_LME"]):
                    continue
                up = r["Sales_LME"] >= r["Supply_LME"]
                figP.add_trace(go.Scatter(
                    x=[r["Supply_LME"], r["Sales_LME"]], y=[r["Fixation"]] * 2, mode="lines",
                    line=dict(color=TEAL if up else ROSE, width=4), showlegend=False, hoverinfo="skip"))
            figP.add_trace(go.Scatter(
                x=gv["Supply_LME"], y=gv["Fixation"], mode="markers", name="Supply LME (€/kg)",
                marker=dict(size=16, color=NAVY_LT, line=dict(color="#ffffff", width=2)),
                hovertemplate="%{x:.4f} €/kg<extra></extra>"))
            figP.add_trace(go.Scatter(
                x=gv["Sales_LME"], y=gv["Fixation"], mode="markers", name="Sales LME (€/kg)",
                marker=dict(size=16, color=COPPER, line=dict(color="#ffffff", width=2)),
                hovertemplate="%{x:.4f} €/kg<extra></extra>"))
            alay(figP, height=120 + 90 * len(gv),
                 legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                 xaxis=dict(title="€/kg"), yaxis=dict(title=""))
            st.plotly_chart(figP, use_container_width=True, theme=None)

        with st.container(border=True):
            sec("📋", "Detail by Fixation", "Supply, sales, prices and coverage — aggregated across the current selection")
            det = gv[["Fixation","Supply_T","Supply_LME","Qty_Sold_T","Sales_LME","Gap_T"]].copy()
            det["Coverage %"] = det.apply(lambda r: _pc(r["Supply_T"], r["Qty_Sold_T"]) if r["Qty_Sold_T"] else None, axis=1)
            det.columns = ["Fixation","Supply (T)","Supply LME (€/kg)","Sold (T)","Sales LME (€/kg)","Gap (T)","Coverage %"]
            det_fmt = {"Supply (T)":"{:,.1f}","Sold (T)":"{:,.1f}","Gap (T)":"{:+,.1f}",
                       "Supply LME (€/kg)":"{:.4f}","Sales LME (€/kg)":"{:.4f}","Coverage %":"{:.0f}%"}
            st.dataframe(
                det.style.format(det_fmt, na_rep="—")
                   .set_properties(**{"background-color": "#ffffff", "color": INK}),
                use_container_width=True, hide_index=True, height=38*len(det)+40)

# ─────────────────────────── TAB: INSIGHTS ───────────────────────────
with tab_insights:
    fx = view_fix.copy()
    num_cols = ["Qty_Sold_T", "Qty_Stock_T", "Qty_Purchase_T", "Allocated_QTE",
                "Sales_Value", "Final_Value", "LME_Balance_Eur"]
    for c in num_cols:
        fx[c] = pd.to_numeric(fx[c], errors="coerce").fillna(0)
    FIX_ORDER = {"M-1": 0, "3M-1": 1, "3M-2": 2}
    ins = fx.groupby("Fixation")[num_cols].sum().reset_index()
    ins["_ord"] = ins["Fixation"].map(FIX_ORDER).fillna(99)
    ins = ins.sort_values("_ord").drop(columns="_ord").reset_index(drop=True)
    ins.columns = ["Fixation", "Sold", "Stock", "Purch", "Realloc", "Sales", "Cost", "Bal"]
    ins["Src"] = ins["Stock"] + ins["Purch"] + ins["Realloc"]

    T_sold, T_stock, T_purch, T_realloc = ins["Sold"].sum(), ins["Stock"].sum(), ins["Purch"].sum(), ins["Realloc"].sum()
    T_src, T_sales, T_bal = ins["Src"].sum(), ins["Sales"].sum(), ins["Bal"].sum()

    if ins.empty or T_sold <= 0 or T_src <= 0:
        st.info("Not enough data to build the insights for the current selection.")
    else:
        spread   = T_bal / (T_sold * 1000)
        res_pct  = _pc(T_bal, T_sales)
        res_col  = TEAL if T_bal >= 0 else ROSE
        pos, neg = ins[ins["Bal"] > 0], ins[ins["Bal"] < 0]

        # ── Hero: the whole operation in three sentences + source-mix bar ──
        def _b(txt, color="#ffffff"):
            return f'<span style="color:{color} !important;">{txt}</span>'

        l1 = (f'Out of every {_b("100 T")} sold, {_b(f"{_pc(T_stock,T_src):.0f} T")} came from stock, '
              f'{_b(f"{_pc(T_purch,T_src):.0f} T")} from purchases and '
              f'{_b(f"{_pc(T_realloc,T_src):.0f} T")} were pulled from another fixation.')
        l2 = (f'Net result: {_b(f"€{T_bal:,.0f}")} — sales were valued {_b(f"{abs(res_pct):.2f}% " + ("above" if T_bal >= 0 else "below"))} '
              f'the FIFO cost, i.e. {_b(f"{spread:+.4f} €/kg")} sold.')
        bits = []
        if not pos.empty:
            top = pos.loc[pos["Bal"].idxmax()]
            bits.append(f'🏆 {_b(top["Fixation"])} is the engine ({_pc(top["Bal"], pos["Bal"].sum()):.0f}% of all gains)')
        if not neg.empty:
            worst = neg.loc[neg["Bal"].idxmin()]
            bits.append(f'📉 {_b(worst["Fixation"])} is the main drag ({_pc(-worst["Bal"], -neg["Bal"].sum()):.0f}% of all losses)')
        l3 = " &nbsp;·&nbsp; ".join(bits)
        mix_bar = _stack_bar([("Stock", T_stock, NAVY_LT), ("Purchases", T_purch, TEAL),
                              ("Reallocation", T_realloc, GOLD)], 40)
        # Rendered in a fully isolated iframe (components.html) so no CSS rule from the
        # rest of the app can ever override the white text here.
        hero_html = f'''<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
  html,body{{margin:0;padding:0;background:transparent;
    font-family:'Inter','Segoe UI',Arial,sans-serif;}}
  .card{{background:linear-gradient(120deg,{NAVY} 0%,{NAVY_MD} 100%);border-radius:18px;
    padding:26px 30px;box-shadow:0 8px 24px rgba(22,38,74,0.18);box-sizing:border-box;}}
  .line{{font-size:1.12rem;line-height:1.6;color:#ffffff;margin:0 0 4px 0;}}
  .soft{{font-size:1.0rem;line-height:1.6;color:#dbe6f8;margin-bottom:16px;}}
  .note{{font-size:0.75rem;color:#9fb4dc;margin-top:8px;}}
  b, strong, span{{color:inherit;font-weight:inherit;}}
</style></head>
<body>
  <div class="card">
    <div class="line">{l1}</div>
    <div class="line" style="margin-bottom:4px;">{l2}</div>
    <div class="soft">{l3}</div>
    {mix_bar}
    <div class="note">Origin of the copper valued against the tonnage sold
      (FIFO order: own stock &rarr; own purchases &rarr; reallocation from another fixation)</div>
  </div>
</body></html>'''
        components.html(hero_html, height=300, scrolling=False)


        # ── 4 headline tiles ──
        t1, t2, t3, t4 = st.columns(4)
        routes = find_realloc_sources(view_fix)
        routes_txt = " · ".join(f"{' + '.join(sorted(s))} → {d}" for d, s in sorted(routes.items())) \
                     if routes else "No cross-fixation borrowing"
        kpi(t1, "🎯", "Own Coverage", f"{_pc(T_stock+T_purch, T_src):.0f}%", NAVY_LT,
            "of tonnage covered by own stock & purchases")
        kpi(t2, "🔁", "Reallocation Reliance", f"{_pc(T_realloc, T_src):.0f}%", GOLD, routes_txt)
        kpi(t3, "📈", "Result vs Sales", f"{res_pct:+.2f}%", res_col, f"{spread:+.4f} €/kg sold")
        kpi(t4, "✅", "Favorable Periods", f"{_pc(n_fav, n_tot):.0f}%", TEAL, f"{n_fav} of {n_tot} entity × month")
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

        # ── Flow (Sankey) ──
        fix_names = list(ins["Fixation"])

        with st.container(border=True):
            sec("🌊", "Where Each Fixation's Copper Comes From", "Tonnes flowing from each source into each fixation (% of total)")
            src_names = ["Stock", "Purchases", "Reallocation"]
            src_cols  = [NAVY_LT, TEAL, GOLD]
            src_vals  = [T_stock, T_purch, T_realloc]
            labels = ([f"{n} · {_pc(v, T_src):.0f}%" for n, v in zip(src_names, src_vals)] +
                      [f"{f} · {_pc(s, T_sold):.0f}%" for f, s in zip(fix_names, ins["Sold"])])
            S, Tg, V, C = [], [], [], []
            for j, (_, r) in enumerate(ins.iterrows()):
                for i, key in enumerate(["Stock", "Purch", "Realloc"]):
                    if r[key] > 0:
                        S.append(i); Tg.append(3 + j); V.append(r[key]); C.append(_rgba(src_cols[i], 0.38))
            figS = go.Figure(go.Sankey(
                node=dict(label=labels, color=src_cols + [NAVY_MD] * len(fix_names),
                          pad=24, thickness=22, line=dict(width=0)),
                link=dict(source=S, target=Tg, value=V, color=C,
                          hovertemplate="%{source.label} → %{target.label}<br>%{value:,.1f} T<extra></extra>")))
            alay(figS, height=420)
            figS.update_layout(font=dict(size=13, color=INK))
            st.plotly_chart(figS, use_container_width=True, theme=None)

        # ── Fixation identity cards ──
        with st.container(border=True):
            sec("🪪", "Fixation Identity Cards", "What each fixation represents — its weight, its result and where its copper came from")
            best_name  = pos.loc[pos["Bal"].idxmax(), "Fixation"] if not pos.empty else None
            worst_name = neg.loc[neg["Bal"].idxmin(), "Fixation"] if not neg.empty else None
            rows = list(ins.iterrows())
            for i in range(0, len(rows), 3):
                cols = st.columns(3)
                for j, (_, r) in enumerate(rows[i:i + 3]):
                    sold = r["Sold"]
                    bal  = r["Bal"]
                    col  = TEAL if bal >= 0 else ROSE
                    sale_p = r["Sales"] / (sold * 1000) if sold else 0
                    cost_p = r["Cost"] / (sold * 1000) if sold else 0
                    sp     = bal / (sold * 1000) if sold else 0
                    vshare = _pc(sold, T_sold)
                    if r["Fixation"] == best_name:
                        tag = "🏆 Top contributor"
                    elif r["Fixation"] == worst_name:
                        tag = "📉 Main drag"
                    else:
                        tag = "▲ Gain" if bal >= 0 else "▼ Loss"
                    src_bar = _stack_bar([("Stock", r["Stock"], NAVY_LT), ("Purchases", r["Purch"], TEAL),
                                          ("Realloc.", r["Realloc"], GOLD)], 10, False)
                    src_leg = (f'Stock {_pc(r["Stock"], r["Src"]):.0f}% · Purchases {_pc(r["Purch"], r["Src"]):.0f}% · '
                               f'Realloc. {_pc(r["Realloc"], r["Src"]):.0f}%')
                    cols[j].markdown(
                        f'<div style="background:#ffffff;border:1px solid #e9edf5;border-top:4px solid {col};'
                        f'border-radius:14px;padding:16px 18px;box-shadow:0 2px 8px rgba(22,38,74,0.05);">'
                        f'<div style="display:flex;justify-content:space-between;align-items:center;">'
                        f'<div style="font-size:1.15rem;font-weight:800;color:{NAVY};">{r["Fixation"]}</div>'
                        f'<div style="font-size:0.72rem;font-weight:700;color:{col};background:{col}1a;'
                        f'padding:3px 10px;border-radius:999px;">{tag}</div></div>'
                        f'<div style="font-size:1.7rem;font-weight:800;color:{col};margin-top:8px;">€{bal:,.0f}</div>'
                        f'<div style="font-size:0.78rem;color:#6b7896;margin-bottom:12px;">'
                        f'{_pc(bal, r["Sales"]):+.2f}% of its sales · {sp:+.4f} €/kg</div>'
                        f'<div style="font-size:0.75rem;color:#6b7896;display:flex;justify-content:space-between;">'
                        f'<span>Weight in volume sold</span><b style="color:{NAVY};">{vshare:.0f}% · {sold:,.0f} T</b></div>'
                        f'<div style="height:8px;border-radius:4px;background:#e9edf5;margin:4px 0 12px 0;">'
                        f'<div style="width:{vshare:.1f}%;height:8px;border-radius:4px;background:{NAVY_MD};"></div></div>'
                        f'<div style="font-size:0.75rem;color:#6b7896;display:flex;justify-content:space-between;margin-bottom:12px;">'
                        f'<span>Sold at → FIFO cost</span><b style="color:{NAVY};">{sale_p:.3f} → {cost_p:.3f} €/kg</b></div>'
                        f'<div style="font-size:0.75rem;color:#6b7896;margin-bottom:4px;">Where the copper came from</div>'
                        f'{src_bar}'
                        f'<div style="font-size:0.7rem;color:#6b7896;margin-top:4px;">{src_leg}</div>'
                        f'</div>', unsafe_allow_html=True)
                st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

# ─────────────────────────── TAB: DATA ───────────────────────────
with tab_data:
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
            .set_properties(**{"background-color":"#ffffff","color":INK})
            .map(lambda v:"color:#0d9488;font-weight:700" if isinstance(v,(int,float)) and v>0
                 else ("color:#e11d48;font-weight:700" if isinstance(v,(int,float)) and v<0 else ""),
                 subset=["LME Balance (€)"])
            .map(lambda v:"font-weight:700;color:#c2703d" if str(v).strip().upper()=="TOTAL" else "",
                 subset=["Fixation"]),
        use_container_width=True, hide_index=True, height=440
    )

    st.download_button(
        "⬇️ Download consolidated balance (CSV)",
        data=view[disp_cols].to_csv(index=False).encode("utf-8"),
        file_name="lme_balance_consolidated.csv", mime="text/csv"
    )

st.markdown(f"""<div style="text-align:center;color:#a3abbd;font-size:0.72rem;
  margin-top:40px;padding:16px;border-top:1px solid #e9edf5;">
  Balance LME &nbsp;·&nbsp; COFICAB Kenitra &amp; COFICAB Maroc
</div>""", unsafe_allow_html=True)
