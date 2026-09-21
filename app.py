"""Beyond the Report: ESG Decision Intelligence Dashboard (APFA802 Major Project 2026).

Run with:  streamlit run app.py
"""
from __future__ import annotations

import base64
import io
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import esg

st.set_page_config(page_title="ESG Decision Dashboard", layout="wide", initial_sidebar_state="collapsed")

DEFAULT_FILE = Path(__file__).parent / "data" / "APFA802_ESG_Workbook.xlsx"
INK, INK2, MUTED = "#14213D", "#4A5775", "#8A94AD"
CYAN, CYAN_SOFT = "#4DD0E1", "#C9F1F6"
PINKS = ["#F47C7C", "#F69494", "#F8ACAC", "#FAC3C3", "#FCD7D7", "#FDE7E7"]
PLOT_CFG = {"displayModeBar": False}

st.markdown(
    """<style>
.stApp {background: linear-gradient(180deg, #E9F0FA 0%, #F2F6FC 40%, #EEF3FA 100%);}
header[data-testid="stHeader"] {background: transparent; height: 0;}
.block-container {padding-top: 0 !important; padding-bottom: 2rem; max-width: 1320px;}
/* header banner */
.topbar {display:flex; align-items:stretch; justify-content:space-between; margin: 0 -1rem 22px -1rem; min-height: 64px;}
.banner {flex: 0 1 68%; display:flex; align-items:center; gap:14px; padding: 12px 28px;
  background: linear-gradient(100deg, #0A3480 0%, #134BA6 60%, #1A5CC2 100%);
  border-bottom-right-radius: 120px 100%; box-shadow: 0 6px 18px rgba(10,52,128,.18);}
.logo {width:40px; height:40px; background:#fff; border-radius:9px; display:flex; align-items:center; justify-content:center; flex:none;}
.banner h1 {color:#fff !important; font-size: 25px !important; font-weight:700 !important; margin:0 !important; padding:0 !important; line-height:1.1;}
.banner small {display:block; color:#B8CCF0; font-size:9.5px; letter-spacing:.14em; font-weight:600; margin-top:2px;}
.actions {display:flex; align-items:center; gap:12px; padding: 10px 6px;}
.btn {text-decoration:none !important; font-weight:600; font-size:14.5px; padding:10px 20px; border-radius:7px; white-space:nowrap;}
.btn.light {background:#DDE3F6; color:#1E2A5A !important;}
.btn.dark {background:#2A1E6E; color:#fff !important;}
/* filters */
div[data-testid="stSelectbox"] > div > div {background:#fff; border:1px solid #B7C3DA; border-radius:6px; min-height:46px;}
div[data-testid="stSelectbox"] label p {font-size:12px; color:${MUTED}; font-weight:600; letter-spacing:.04em; text-transform:uppercase;}
/* KPI cards */
.kpis {display:grid; grid-template-columns: repeat(4, 1fr); gap:16px; margin: 10px 0 18px 0;}
.kpi {border-radius:10px; padding:16px 18px 12px 18px; color:#fff; min-height:112px; display:flex; flex-direction:column; box-shadow:0 4px 14px rgba(20,33,61,.12);}
.kpi .l {font-size:15px; font-weight:500; opacity:.95;}
.kpi .v {font-size:32px; font-weight:600; line-height:1.15; margin-top:2px; font-variant-numeric: tabular-nums;}
.kpi .f {display:flex; justify-content:space-between; align-items:flex-end; margin-top:auto; font-size:13px;}
.kpi .f i {opacity:.9;}
.k1 {background: linear-gradient(120deg, #0A4AA8, #1766C9);}
.k2 {background: linear-gradient(120deg, #0B86C8, #19A3DD);}
.k3 {background: linear-gradient(120deg, #13B983, #22CC8A);}
.k4 {background: linear-gradient(120deg, #6C3FB0, #8C55C9);}
/* panels */
div[class*="st-key-panel"] {background:#fff; border:1px solid #E1E8F3 !important; border-radius:10px; box-shadow: 0 2px 10px rgba(20,33,61,.05); padding:14px 16px;}
.ptitle {font-size:17px; font-weight:700; color:#14213D; margin: 2px 0 6px 0;}
.ptitle span {font-size:13.5px; font-weight:400; font-style:italic; color:#8A94AD; margin-left:14px;}
.stitle {font-size:17px; font-weight:700; color:#14213D; margin: 18px 0 6px 2px;}
.stitle span {font-size:13.5px; font-weight:400; font-style:italic; color:#8A94AD; margin-left:14px;}
/* tables */
table.t {width:100%; border-collapse:collapse; font-size:14px; color:#14213D; font-variant-numeric: tabular-nums;}
table.t th {background:#E8EFFA; font-weight:700; text-align:right; padding:7px 10px; border-bottom:1px solid #D5DFEE; white-space:nowrap;}
table.t th:first-child, table.t td:first-child {text-align:left;}
table.t td {padding:7px 10px; text-align:right; border-bottom:1px solid #EEF2F8; white-space:nowrap;}
table.t td.tx, table.t th.tx {text-align:left; white-space:normal;}
.tscroll {overflow-x:auto;}
.chip {white-space:normal;}
table.t tr:hover td {filter: brightness(.97);}
.chip {display:inline-block; padding:2px 9px; border-radius:20px; font-size:12.5px; font-weight:600;}
.chip.good {background:#DDF4E4; color:#1D6B36;} .chip.warn {background:#FFF1D1; color:#7A5300;} .chip.bad {background:#FBE0DC; color:#8F2A1D;}
.note {color:#6B7690; font-size:12.5px; margin-top:6px;}
.foot {color:#8A94AD; font-size:12px; margin-top:24px;}
@media (max-width: 900px) {.kpis {grid-template-columns: repeat(2, 1fr);} .banner {flex-basis:100%; border-bottom-right-radius:40px 100%;} .actions {display:none;}}
</style>""".replace("${MUTED}", MUTED),
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------- helpers
@st.cache_data(show_spinner=False)
def load(file_bytes: bytes | None):
    return esg.load(DEFAULT_FILE if file_bytes is None else io.BytesIO(file_bytes))


def num(v, dec=0, suffix="") -> str:
    if v is None or pd.isna(v):
        return "–"
    return f"{v:,.{dec}f}{suffix}"


def mix(c1: str, c2: str, t: float) -> str:
    """Blend hex c1 -> c2 by t (0..1)."""
    t = max(0.0, min(1.0, float(t)))
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(a, b))


def delta(cur, prev, fy) -> str:
    if prev is None or pd.isna(prev) or pd.isna(cur) or prev == 0:
        return ""
    ch = (cur - prev) / abs(prev) * 100
    arrow = "▲" if ch >= 0 else "▼"
    return f"{arrow} {abs(ch):.1f}% vs FY{fy - 1}"


def panel_title(title: str, sub: str = "") -> None:
    st.markdown(f'<div class="ptitle">{title}{f"<span>{sub}</span>" if sub else ""}</div>', unsafe_allow_html=True)


def base_layout(fig: go.Figure, h: int, **kw) -> go.Figure:
    fig.update_layout(height=h, margin=dict(l=8, r=16, t=8, b=8), plot_bgcolor="#fff", paper_bgcolor="#fff",
                      font=dict(family="Barlow, sans-serif", size=13, color=INK2),
                      hoverlabel=dict(bgcolor="#fff", bordercolor="#D5DFEE", font=dict(color=INK, family="Barlow, sans-serif")), **kw)
    return fig


def reset_weights() -> None:
    for k, w in esg.DEFAULT_WEIGHTS.items():
        st.session_state[f"w_{k}"] = w


for k, w in esg.DEFAULT_WEIGHTS.items():
    st.session_state.setdefault(f"w_{k}", w)

# ---------------------------------------------------------------- data
upload = st.session_state.get("upload")
up_bytes = upload.getvalue() if upload is not None else None
try:
    D = load(up_bytes)
except Exception as e:  # a clear message rather than a stack trace
    st.error(f"Could not read the workbook: {e}")
    st.stop()
weights = {k: st.session_state[f"w_{k}"] for k in esg.DEFAULT_WEIGHTS}
df = esg.esg_index(D["raw"], weights)
companies = list(dict.fromkeys(df.name))
years = sorted(df.fy.unique())

# ---------------------------------------------------------------- header
csv_b64 = base64.b64encode(df.drop(columns=[c for c in df.columns if c.startswith("n_")]).to_csv(index=False).encode()).decode()
logo = ('<svg width="24" height="24" viewBox="0 0 24 24" fill="none"><path d="M4 20c0-9 6-15 16-16-1 10-7 16-16 16Z" fill="#1552B0"/>'
        '<path d="M4 20 13 11" stroke="#fff" stroke-width="2" stroke-linecap="round"/></svg>')
st.markdown(
    f'<div class="topbar"><div class="banner"><div class="logo">{logo}</div><div><h1>ESG decision dashboard</h1>'
    f'<small>APFA802 · BEYOND THE REPORT</small></div></div><div class="actions">'
    f'<a class="btn light" href="#data-and-sources" target="_self">Explore sources</a>'
    f'<a class="btn dark" download="apfa802_esg_dataset.csv" href="data:text/csv;base64,{csv_b64}">Download dataset</a></div></div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------- filters
f1, f2, _ = st.columns([1.35, 1.35, 3])
fy = f1.selectbox("Fiscal year", years[::-1], format_func=lambda y: f"FY{y}", key="fy")
co = f2.selectbox("Company", ["All companies"] + companies, key="co")
scope = companies if co == "All companies" else [co]
cur = df[(df.fy == fy) & df.name.isin(scope)]
prev = df[(df.fy == fy - 1) & df.name.isin(scope)]
scope_label = "all 5 companies" if co == "All companies" else co

# ---------------------------------------------------------------- KPI cards
rev, rev_p = cur.revenue.sum(), prev.revenue.sum() if len(prev) else None
ghg, ghg_p = cur.ghg.sum(min_count=1), prev.ghg.sum(min_count=1) if len(prev) else None
esg_v, esg_p = cur.esg.mean(), prev.esg.mean() if len(prev) else None
fat = cur.fatalities.sum(min_count=1)
fat_gap = cur[cur.fatalities.isna()].name.tolist()
if co == "All companies":
    esg_foot = f"mean of 5 · {delta(esg_v, esg_p, fy)}"
else:
    esg_foot = f"rank {int(cur['rank'].iloc[0])} of {len(companies)} in FY{fy}"
ghg_txt = f"{ghg / 1e6:,.2f} Mt" if ghg and ghg >= 1e6 else num(ghg)
cards = [
    ("k1", "Revenue", f"R{rev:,.0f}m", delta(rev, rev_p, fy), "in R million"),
    ("k2", "Scope 1+2 emissions", ghg_txt, delta(ghg, ghg_p, fy), "tCO2e"),
    ("k3", "ESG index score", num(esg_v, 1), esg_foot, "0–100 scale"),
    ("k4", "Workplace fatalities", num(fat), ("not disclosed: " + ", ".join(fat_gap)) if fat_gap else "all disclosed", f"FY{fy}"),
]
st.markdown('<div class="kpis">' + "".join(
    f'<div class="kpi {c}"><div class="l">{l}</div><div class="v">{v}</div><div class="f"><span>{d}</span><i>{u}</i></div></div>'
    for c, l, v, d, u in cards) + "</div>", unsafe_allow_html=True)

# ---------------------------------------------------------------- row: bar + donut
yr = df[df.fy == fy].sort_values("esg")
c1, c2 = st.columns([1.45, 1])
with c1, st.container(key="panel_1"):
    panel_title(f"ESG index score by company", f"FY{fy} · 0–100, higher is better")
    colors = [CYAN if (co == "All companies" or n == co) else CYAN_SOFT for n in yr.name]
    fig = go.Figure(go.Bar(
        y=yr.name, x=yr.esg, orientation="h", marker=dict(color=colors, line=dict(width=0)),
        text=[f"{v:.1f}" for v in yr.esg], textposition="outside", cliponaxis=False, textfont=dict(color=INK, size=13),
        customdata=yr[["rank", "sector"]].values,
        hovertemplate="<b>%{y}</b><br>ESG index %{x:.1f} · rank %{customdata[0]}<br>%{customdata[1]}<extra></extra>",
    ))
    base_layout(fig, 300, bargap=0.38, barcornerradius=4,
                xaxis=dict(range=[0, 100], dtick=20, gridcolor="#EEF2F8", zeroline=False, tickfont=dict(color=MUTED)),
                yaxis=dict(tickfont=dict(color=INK, size=13.5)))
    st.plotly_chart(fig, width="stretch", config=PLOT_CFG)

with c2, st.container(key="panel_2"):
    if co == "All companies":
        parts = pd.DataFrame([esg.contributions(r, weights) for _, r in df[df.fy == fy].iterrows()]).mean()
        sub = f"FY{fy} · peer average"
    else:
        parts = pd.Series(esg.contributions(cur.iloc[0], weights))
        sub = f"FY{fy} · {co}"
    panel_title("What makes up the score", sub)
    fig = go.Figure(go.Pie(
        labels=parts.index, values=parts.values, hole=0.52, sort=False, direction="clockwise",
        marker=dict(colors=PINKS[:len(parts)], line=dict(color="#fff", width=2)),
        texttemplate="%{percent:.1%}", textfont=dict(color=INK, size=12.5), insidetextorientation="horizontal",
        hovertemplate="<b>%{label}</b><br>%{value:.1f} points of the score (%{percent})<extra></extra>",
    ))
    base_layout(fig, 300, showlegend=True, legend=dict(x=1.0, y=0.5, yanchor="middle", font=dict(color=INK, size=13)))
    st.plotly_chart(fig, width="stretch", config=PLOT_CFG)

# ---------------------------------------------------------------- row: two tables
t1, t2 = st.columns([1, 1])
with t1, st.container(key="panel_3"):
    panel_title("ESG index by year", "0–100 · change FY2021 to selected year")
    piv = df.pivot_table(index="name", columns="fy", values="esg").reindex(scope)
    shown = [y for y in years if y <= fy]
    head = "<tr><th>Company</th>" + "".join(f"<th>FY{str(y)[2:]}</th>" for y in shown) + "<th>Change</th></tr>"
    body = ""
    for n, r in piv.iterrows():
        cells = "".join(f'<td style="background:{mix("#FFFFFF", "#8EDBE6", (r[y] - 20) / 70)}">{r[y]:.1f}</td>' for y in shown)
        ch = r[fy] - r[shown[0]]
        chc = mix("#FFFFFF", "#9FE0B6" if ch >= 0 else "#F6B3A6", min(abs(ch) / 20, 1))
        body += f'<tr><td>{n}</td>{cells}<td style="background:{chc};font-weight:600">{ch:+.1f}</td></tr>'
    st.markdown(f'<table class="t">{head}{body}</table>', unsafe_allow_html=True)

with t2, st.container(key="panel_4"):
    panel_title("Financial snapshot", f"FY{fy} · R million")
    snap = cur.sort_values("revenue", ascending=False)
    rmax = df[df.fy == fy].revenue.max()
    head = "<tr><th>Company</th><th>Op. margin</th><th>Net margin</th><th>Debt / equity</th><th>ROE</th><th>Revenue</th></tr>"
    body = ""
    for _, r in snap.iterrows():
        body += (f"<tr><td>{r['name']}</td><td>{num(r.op_margin, 1, '%')}</td><td>{num(r.net_margin, 1, '%')}</td>"
                 f"<td>{num(r.de, 2)}</td><td>{num(r.roe, 1, '%')}</td>"
                 f'<td style="background:{mix("#FFF4EF", "#F9A98E", (r.revenue / rmax) ** 0.35)};font-weight:600">{num(r.revenue)}</td></tr>')
    st.markdown(f'<table class="t">{head}{body}</table>', unsafe_allow_html=True)
    st.markdown('<div class="note">“–” means not disclosed in that year’s integrated report (left blank, not zero). '
                "ArcelorMittal SA’s FY2025 equity is negative, which is why its ROE and debt/equity look extreme.</div>",
                unsafe_allow_html=True)

# ---------------------------------------------------------------- trends
METRICS = {
    "ESG index score": ("esg", "0–100"), "Revenue": ("revenue", "R million"), "Operating margin": ("op_margin", "%"),
    "Net margin": ("net_margin", "%"), "Return on assets": ("roa", "%"), "Debt to equity": ("de", "ratio"),
    "Scope 1+2 emissions": ("ghg", "tCO2e"), "GHG intensity": ("ghg_int", "tCO2e per R million revenue"),
    "Water intensity": ("water_int", "m³ per R million revenue"), "Safety rate": ("safety", "company's own metric"),
    "Women on board": ("board_women", "%"), "B-BBEE level": ("bbee", "1 is best, 8 is weakest"),
    "CSI spend": ("csi", "R million"), "CSI intensity": ("csi_int", "% of revenue"), "Employees": ("employees", "headcount"),
}
with st.container(key="panel_5"):
    h1, h2, h3 = st.columns([2.2, 1.2, 1])
    with h1:
        panel_title("Five-year trend", "FY2021 – FY2025")
    metric = h2.selectbox("Measure", list(METRICS), key="metric", label_visibility="collapsed")
    indexed = h3.toggle("Index FY2021 = 100", key="indexed")
    col, unit = METRICS[metric]
    fig = go.Figure()
    for n in companies:
        s = df[df.name == n].set_index("fy")[col].dropna()
        if s.empty:
            continue
        y = s / abs(s.iloc[0]) * 100 if indexed and s.iloc[0] else s
        on = co == "All companies" or n == co
        fig.add_trace(go.Scatter(
            x=[f"FY{v}" for v in s.index], y=y, name=n, mode="lines+markers",
            line=dict(color=esg.COMPANY_COLOUR[n], width=2.5 if on else 1.5), marker=dict(size=8, line=dict(color="#fff", width=2)),
            opacity=1 if on else 0.3, hovertemplate=f"{n}: %{{y:,.2f}}<extra></extra>",
        ))
    base_layout(fig, 360, hovermode="x unified",
                legend=dict(orientation="h", y=1.08, x=0, font=dict(color=INK, size=13)),
                yaxis=dict(title=dict(text="Index (FY2021 = 100)" if indexed else unit, font=dict(color=MUTED)),
                           gridcolor="#EEF2F8", zerolinecolor="#C9D3E3", tickfont=dict(color=MUTED)),
                xaxis=dict(tickfont=dict(color=MUTED), categoryorder="array", categoryarray=[f"FY{y}" for y in years]))
    if metric == "Safety rate":
        st.markdown('<div class="note">Afrimat, ArcelorMittal SA and Omnia (from FY2023) report LTIFR; AECI reports TRIR; Sasol and Omnia '
                    "(before FY2023) report RCR. Compare each company's own trend, not levels across companies.</div>", unsafe_allow_html=True)
    st.plotly_chart(fig, width="stretch", config=PLOT_CFG)

# ---------------------------------------------------------------- targets vs actual
st.markdown('<div class="stitle">Stated GHG targets vs observed emissions<span>Scope 1+2, FY2021 to latest year</span></div>', unsafe_allow_html=True)
with st.container(key="panel_6"):
    tv = D["targets"]
    tv = tv[tv.Company.isin(scope)]
    head = "<tr><th>Company</th><th>FY2021 (tCO2e)</th><th>Latest (tCO2e)</th><th>Observed change</th><th>Stated target</th><th class=\"tx\">Credibility flag</th><th class=\"tx\">Why</th></tr>"
    body = ""
    for _, r in tv.iterrows():
        flag = str(r["Credibility Flag"])
        cls = "bad" if flag.startswith("Inconsistent") else "warn" if "behind" in flag else "good"
        icon = {"bad": "✕", "warn": "!", "good": "✓"}[cls]
        ch = r["Observed % Change"]
        target = r["Stated Target (%, by year)"]
        chc = mix("#FFFFFF", "#9FE0B6" if ch < 0 else "#F6B3A6", min(abs(ch) / 50, 1))
        body += (f"<tr><td>{r.Company}</td><td>{num(r['FY2021 GHG (tCO2e)'])}</td><td>{num(r['Latest-year GHG (tCO2e)'])}</td>"
                 f'<td style="background:{chc};font-weight:600">{ch:+.1f}%</td><td>{"none stated" if not target else f"−{target:.0f}%"}</td>'
                 f'<td class="tx" style="width:210px"><span class="chip {cls}">{icon} {flag}</span></td><td class="tx" style="font-size:12.5px;color:#4A5775;min-width:260px">{r.Notes}</td></tr>')
    st.markdown(f'<div class="tscroll"><table class="t">{head}{body}</table></div>', unsafe_allow_html=True)
    st.markdown('<div class="note">Targets use each company’s own baseline year (Sasol 2017, ArcelorMittal SA 2018, Omnia FY2024), so this checks '
                "direction of travel, not literal target attainment. A gap is a signal to investigate, not proof of greenwashing.</div>", unsafe_allow_html=True)

# ---------------------------------------------------------------- correlation + data quality
q1, q2 = st.columns([1.3, 1])
with q1, st.container(key="panel_7"):
    panel_title("How ESG and financial measures move together", "Pearson r, 25 company-years")
    cm = {"Op. margin": "op_margin", "Net margin": "net_margin", "ROA": "roa", "Debt/equity": "de", "GHG intensity": "ghg_int",
          "Water intensity": "water_int", "CSI intensity": "csi_int", "Women on board": "board_women", "ESG index": "esg"}
    corr = df[list(cm.values())].corr().values
    labels = list(cm)
    fig = go.Figure(go.Heatmap(
        z=corr, x=labels, y=labels, zmin=-1, zmax=1, xgap=2, ygap=2,
        colorscale=[[0, "#E34948"], [0.5, "#F0EFEC"], [1, "#2A78D6"]],
        text=[[f"{v:.2f}" for v in row] for row in corr], texttemplate="%{text}", textfont=dict(size=11.5, color=INK),
        hovertemplate="%{y} × %{x}: r = %{z:.2f}<extra></extra>", colorbar=dict(thickness=10, outlinewidth=0, tickfont=dict(color=MUTED)),
    ))
    base_layout(fig, 400, xaxis=dict(tickangle=-35, tickfont=dict(color=INK2)), yaxis=dict(autorange="reversed", tickfont=dict(color=INK2)))
    st.plotly_chart(fig, width="stretch", config=PLOT_CFG)
    st.markdown('<div class="note">Blue = move together, red = move in opposite directions. Correlation across five companies is not causation.</div>',
                unsafe_allow_html=True)

with q2, st.container(key="panel_8"):
    cit = D["citations"]
    panel_title("Source verification", f"{len(cit)} citation rows")
    status = cit.Status.str.extract(r"^(CONFIRMED|DERIVED|RESTATED|CORRECTED|GENUINELY ABSENT)")[0].fillna("OTHER")
    vc = status.value_counts().reindex(["CONFIRMED", "DERIVED", "RESTATED", "CORRECTED", "GENUINELY ABSENT"]).fillna(0)
    fig = go.Figure(go.Bar(
        y=[s.title() for s in vc.index][::-1], x=vc.values[::-1], orientation="h", marker=dict(color=CYAN),
        text=[f"{int(v)}" for v in vc.values[::-1]], textposition="outside", cliponaxis=False, textfont=dict(color=INK),
        hovertemplate="%{y}: %{x} rows<extra></extra>",
    ))
    base_layout(fig, 250, bargap=0.4, barcornerradius=4, xaxis=dict(showgrid=False, showticklabels=False), yaxis=dict(tickfont=dict(color=INK, size=13)))
    st.plotly_chart(fig, width="stretch", config=PLOT_CFG)
    confirmed = status.isin(["CONFIRMED", "DERIVED", "RESTATED", "CORRECTED"]).mean()
    st.markdown(f'<div class="note"><b>{confirmed:.0%}</b> of citations trace a value to a specific page in the company’s own integrated report. '
                "A second-pass re-check corrected 8 transcription errors; “genuinely absent” fields are left blank, never estimated.</div>",
                unsafe_allow_html=True)

# ---------------------------------------------------------------- data and sources
st.markdown('<div id="data-and-sources" class="stitle">Data and sources<span>every value traceable to a report page</span></div>', unsafe_allow_html=True)
with st.container(key="panel_9"):
    tab1, tab2, tab3 = st.tabs(["Source citations", "Raw data", "Index weights"])
    with tab1:
        a, b, c = st.columns([1, 1, 2])
        cc = a.selectbox("Company", ["All"] + companies, key="c_co")
        cs = b.selectbox("Status", ["All"] + sorted(cit.Status.dropna().unique()), key="c_st")
        cq = c.text_input("Search variable, section or note", key="c_q")
        v = cit.copy()
        if cc != "All":
            v = v[v.Company == cc]
        if cs != "All":
            v = v[v.Status == cs]
        if cq:
            v = v[(v.Variable.astype(str) + " " + v["IAR Section"].astype(str) + " " + v.Note.fillna("").astype(str)).str.contains(cq, case=False, regex=False)]
        st.caption(f"{len(v)} of {len(cit)} rows")
        st.dataframe(v, hide_index=True, width="stretch", height=420)
    with tab2:
        show = df[["name", "ticker", "sector", "fy", "revenue", "op_profit", "net_income", "assets", "equity", "debt", "eps", "scope1", "scope2",
                   "ghg", "water", "employees", "safety_metric", "safety", "fatalities", "bbee", "board_women", "ghg_target", "ghg_target_year",
                   "csi", "esg", "rank"]]
        st.dataframe(show, hide_index=True, width="stretch", height=420, column_config={"esg": st.column_config.NumberColumn("ESG index", format="%.1f")})
        st.download_button("Download as CSV", show.to_csv(index=False).encode("utf-8"), "apfa802_esg_dataset.csv", "text/csv")
        st.file_uploader("Replace the workbook (optional)", type=["xlsx"], key="upload",
                         help="Upload an updated workbook with the same 'Raw Data', 'Source Citations' and 'Target vs Actual Check' sheets.")
    with tab3:
        st.caption("Each sub-indicator is min-max scaled to 0–100 across all 25 company-years (worst = 0, best = 100). GHG intensity, water "
                   "intensity, safety rate and B-BBEE level are inverted because lower is better. Missing values are skipped and the other "
                   "weights rescaled. Defaults match the workbook.")
        cols = st.columns(3)
        for i, (k, label, *_) in enumerate(esg.INDEX_PARTS):
            cols[i % 3].slider(f"{label} (%)", 0, 50, step=5, key=f"w_{k}")
        tot = sum(st.session_state[f"w_{k}"] for k in esg.DEFAULT_WEIGHTS)
        st.caption(f"Weights add up to {tot}%. Scores are rescaled by the total, so they stay on 0–100.")
        st.button("Reset to workbook weights", on_click=reset_weights)

st.markdown('<div class="foot">APFA802 Major Project 2026 · Afrimat, AECI, ArcelorMittal SA, Omnia and Sasol integrated annual reports FY2021–FY2025. '
            "All values in South African Rand. Omnia’s fiscal year ends 31 March, up to nine months before the others.</div>", unsafe_allow_html=True)
