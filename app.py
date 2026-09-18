"""Beyond the Report: ESG Decision Intelligence Dashboard (APFA802 Major Project 2026).

Run with:  streamlit run app.py
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import scoring as sc

st.set_page_config(page_title="Beyond the Report: ESG Dashboard", layout="wide")

DEFAULT_FILE = Path(__file__).parent / "data" / "APFA802_Dataset.xlsx"

st.markdown(
    """
    <style>
    h1, h2, h3 {font-family: Georgia, "Times New Roman", serif !important; letter-spacing: -0.01em;}
    [data-testid="stMetricValue"] {font-family: Georgia, serif;}
    .flagbox {border-left: 3px solid #A63F2B; background: #E4EAE6; padding: 10px 14px; border-radius: 0 6px 6px 0; margin-top: 8px;}
    .muted {color: #56675F; font-size: 0.9rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------- helpers
@st.cache_data(show_spinner=False)
def load(file_bytes: bytes | None):
    src = DEFAULT_FILE if file_bytes is None else __import__("io").BytesIO(file_bytes)
    return sc.load_dataset(src)


def fmt(v) -> str:
    if v is None or pd.isna(v):
        return "–"
    a = abs(v)
    if a >= 1000:
        return f"{v:,.0f}".replace(",", " ")
    if a >= 100:
        return f"{v:.0f}"
    if a >= 10:
        return f"{v:.1f}".rstrip("0").rstrip(".")
    return f"{v:.2f}".rstrip("0").rstrip(".")


def score_colour(s):
    if s is None:
        return "#8A8C7E"
    return "#23735F" if s >= 3.5 else "#A63F2B" if s <= 2.5 else "#8A8C7E"


def reset_model():
    for k, v in sc.DEFAULTS.items():
        st.session_state[k] = v


for k, v in sc.DEFAULTS.items():
    st.session_state.setdefault(k, v)

# ---------------------------------------------------------------- sidebar
with st.sidebar:
    st.header("Data")
    up = st.file_uploader("Replace the dataset (optional)", type=["xlsx"],
                          help="Upload an updated APFA802_Dataset.xlsx. It must contain a sheet named 'Dataset' with the same columns.")
    st.header("Scoring model")
    st.radio("Weighting", ["Group's weights", "By pillar"], key="mode",
             help="Group's weights uses the indicator weights in the Excel scorecard. By pillar splits each pillar's weight equally across a company's indicators in that pillar.")
    if st.session_state.mode == "By pillar":
        st.slider("Financial", 0, 100, key="wF")
        st.slider("Environmental", 0, 100, key="wE")
        st.slider("Social", 0, 100, key="wS")
        tot = st.session_state.wF + st.session_state.wE + st.session_state.wS or 1
        st.caption(f"Normalised: Financial {st.session_state.wF / tot:.0%}, Environmental {st.session_state.wE / tot:.0%}, Social {st.session_state.wS / tot:.0%}. "
                   "A company with no indicators in a pillar is scored on the pillars it does report.")
    st.slider("Change needed for a score of 1 or 5 (±%)", 10, 100, step=5, key="thr")
    st.checkbox("Leave out indicators with no baseline year instead of scoring them a neutral 3", key="exnb")
    st.button("Reset to the group's model", on_click=reset_model)

try:
    df, load_warnings = load(up.getvalue() if up else None)
except Exception as e:  # show a clear message rather than a stack trace
    st.error(f"Could not read the dataset: {e}")
    st.stop()
for w in load_warnings:
    st.warning(w)

tot = st.session_state.wF + st.session_state.wE + st.session_state.wS or 1
pillar_w = {"F": st.session_state.wF / tot * 100, "E": st.session_state.wE / tot * 100, "S": st.session_state.wS / tot * 100}
R = sc.compute(df, st.session_state.mode, pillar_w, st.session_state.thr, st.session_state.exnb)

# ---------------------------------------------------------------- header
st.title("Beyond the Report")
st.markdown(
    '<p class="muted">An ESG-adjusted financial risk scorecard for five JSE-listed chemicals and basic-materials companies, '
    "built from their integrated reports for FY2021 to FY2025. Every figure links back to its source page.</p>",
    unsafe_allow_html=True,
)
m1, m2, m3, m4 = st.columns(4)
m1.metric("Companies", df.Company.nunique())
m2.metric("Period", f"FY{int(df.Year.min()) % 100}–{int(df.Year.max()) % 100}")
m3.metric("Observations", len(df))
m4.metric("Page-traced", f"{df.Page.notna().mean():.0%}")

tabs = st.tabs(["Risk scorecard", "Company explorer", "Cross-company comparison", "Disclosure coverage", "Data and sources", "Method and data checks"])

# ================================================================ 1. SCORECARD
with tabs[0]:
    st.subheader("Which company carries the strongest ESG-adjusted position?")
    st.caption("Composite score from 1 (weakest) to 5 (strongest), based on how each indicator moved between its baseline and latest year.")

    rank = sorted(sc.COMPANIES, key=lambda c: R[c].composite or 0)  # ascending for a horizontal bar chart
    fig = go.Figure(go.Bar(
        y=rank, x=[(R[c].composite or 1) - 1 for c in rank], base=1, orientation="h",
        marker_color=[score_colour(R[c].composite) for c in rank],
        text=[f"{R[c].composite:.2f}" for c in rank], textposition="outside", cliponaxis=False,
        customdata=[[R[c].pillars.get(p) for p in "FES"] for c in rank],
        hovertemplate="<b>%{y}</b><br>Composite %{text}<extra></extra>",
    ))
    fig.add_vline(x=3, line_dash="dash", line_color="#56675F", annotation_text="neutral", annotation_position="top")
    fig.update_layout(height=340, margin=dict(l=10, r=40, t=30, b=30), xaxis=dict(range=[1, 5.3], dtick=1, title="Composite score"),
                      plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", font=dict(size=14), yaxis=dict(tickfont=dict(size=15)))
    c1, c2 = st.columns([3, 2])
    with c1:
        st.plotly_chart(fig, width="stretch")
    with c2:
        ptab = pd.DataFrame([{
            "Company": c,
            "Composite": R[c].composite,
            "Fin.": R[c].pillars.get("F"),
            "Env.": R[c].pillars.get("E"),
            "Soc.": R[c].pillars.get("S"),
            "Checks": len(R[c].flags),
        } for c in rank[::-1]])
        st.markdown("**Pillar sub-scores**")
        st.dataframe(ptab, hide_index=True, width="stretch", column_config={
            k: st.column_config.NumberColumn(k, format="%.2f") for k in ["Composite", "Fin.", "Env.", "Soc."]
        } | {"Checks": st.column_config.NumberColumn("Checks", help="Issues to read before relying on the score")})
        st.caption("An empty pillar score means the company has no scored indicator in that pillar.")

    st.divider()
    sel = st.selectbox("See what drives a company's score", rank[::-1], key="sel")
    r = R[sel]
    st.subheader(f"What drives {sel}'s score")
    st.caption(f"Composite {r.composite:.2f} of 5")
    bd = r.items.copy()
    bd["Pillar"] = bd["Pillar"].map(sc.PILLAR_NAME)
    bd["Baseline"] = bd.apply(lambda x: f"{fmt(x.Baseline)} (FY{int(x['Baseline FY'])})" if pd.notna(x["Baseline FY"]) and pd.notna(x.Baseline) else "none", axis=1)
    bd["Latest"] = bd.apply(lambda x: f"{fmt(x.Latest)} (FY{int(x['Latest FY'])})" if pd.notna(x.Latest) else "–", axis=1)
    bd["Change"] = bd["Change"] * 100
    bd["Weight share"] = bd["Weight share"] * 100
    bd["Note"] = bd.apply(lambda x: "No baseline: neutral 3" + (" (left out)" if not x.Included else "") if x.Neutral else "", axis=1)
    st.dataframe(
        bd[["Indicator", "Pillar", "Direction", "Unit", "Baseline", "Latest", "Change", "Score", "Weight share", "Contribution", "Source", "Note"]],
        hide_index=True, width="stretch",
        column_config={
            "Change": st.column_config.NumberColumn("Change", format="%+.1f%%"),
            "Score": st.column_config.ProgressColumn("Score (1–5)", min_value=1, max_value=5, format="%.2f"),
            "Weight share": st.column_config.NumberColumn("Weight", format="%.1f%%"),
            "Contribution": st.column_config.NumberColumn("Contribution", format="%.2f", help="Score × weight share. Contributions add up to the composite."),
        },
    )
    if r.flags:
        st.markdown('<div class="flagbox"><b>Read this score with care</b><ul>' + "".join(f"<li>{f}</li>" for f in r.flags) + "</ul></div>",
                    unsafe_allow_html=True)

# ================================================================ 2. EXPLORER
with tabs[1]:
    st.subheader("Company explorer")
    st.caption("Every indicator extracted for a company, with its trend across the years disclosed.")
    co = st.radio("Company", sc.COMPANIES, horizontal=True, key="ex_co")
    g = df[df.Company == co].sort_values(["Category", "Indicator", "Year"])
    summ = []
    for ind, x in g.groupby("Indicator", sort=False):
        x = x.sort_values("Year")
        last = x.iloc[-1]
        summ.append({"Indicator": ind, "Category": last.Category, "Capital": last.Capital, "Latest": last.Value, "Unit": last.Unit,
                     "Latest FY": f"FY{int(last.Year)}", "Years": len(x), "Trend": x.Value.tolist()})
    st.dataframe(pd.DataFrame(summ), hide_index=True, width="stretch", column_config={
        "Latest": st.column_config.NumberColumn("Latest", format="%.2f"),
        "Trend": st.column_config.LineChartColumn("Trend", width="medium"),
    })
    ind = st.selectbox("Indicator detail", [s["Indicator"] for s in summ], key="ex_ind")
    x = g[g.Indicator == ind].sort_values("Year")
    d1, d2 = st.columns([3, 2])
    with d1:
        f2 = go.Figure(go.Scatter(x=[f"FY{y}" for y in x.Year], y=x.Value, mode="lines+markers",
                                  line=dict(color=sc.COMPANY_COLOUR[co], width=3), marker=dict(size=10),
                                  customdata=x.SourceShort, hovertemplate="%{x}: %{y:,}<br>%{customdata}<extra></extra>"))
        f2.update_layout(height=340, margin=dict(l=10, r=10, t=30, b=30), yaxis_title=x.Unit.iloc[0], title=f"{ind} ({co})",
                         plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(f2, width="stretch")
    with d2:
        st.dataframe(x.assign(FY=x.Year.map(lambda y: f"FY{y}"))[["FY", "Value", "Unit", "SourceShort", "Notes"]]
                     .rename(columns={"SourceShort": "Source"}), hide_index=True, width="stretch")
        if x.Page.isna().any():
            st.warning("Some values here have no page reference.")

# ================================================================ 3. COMPARE
with tabs[2]:
    st.subheader("Cross-company comparison")
    st.caption("Indicators that measure the same thing, brought onto a common unit.")
    cc1, cc2 = st.columns([2, 3])
    meas = cc1.selectbox("Measure", list(sc.COMPARE), key="cmp")
    idx_on = cc2.checkbox("Index each company's first disclosed year to 100", key="cmp_idx")
    unit, mapping, caveat = sc.COMPARE[meas]
    look = {(r_.Company, r_.Indicator, int(r_.Year)): r_ for r_ in df.itertuples(index=False)}
    series = {}
    for c, lst in mapping.items():
        pts = {}
        for ind_, mult in lst:
            for y in sc.YEARS:
                if y in pts:
                    continue
                if ind_ == "__S12":
                    a, b = look.get((c, "Scope 1 GHG emissions", y)), look.get((c, "Scope 2 GHG emissions", y))
                    if a and b:
                        pts[y] = (a.Value + b.Value, a.SourceShort + " (Scope 1 + 2)")
                else:
                    rr = look.get((c, ind_, y))
                    if rr:
                        pts[y] = (rr.Value * mult, rr.SourceShort)
        if pts:
            series[c] = dict(sorted(pts.items()))
    f3 = go.Figure()
    for c, pts in series.items():
        ys = [v for v, _ in pts.values()]
        if idx_on and ys[0]:
            ys = [v / abs(ys[0]) * 100 for v in ys]
        f3.add_trace(go.Scatter(x=[f"FY{y}" for y in pts], y=ys, name=c, mode="lines+markers",
                                line=dict(color=sc.COMPANY_COLOUR[c], width=3), marker=dict(size=9),
                                customdata=[s for _, s in pts.values()], hovertemplate=f"<b>{c}</b> %{{x}}: %{{y:,.1f}}<br>%{{customdata}}<extra></extra>"))
    f3.update_layout(height=420, margin=dict(l=10, r=10, t=20, b=30), yaxis_title="Index" if idx_on else unit,
                     xaxis=dict(categoryorder="array", categoryarray=[f"FY{y}" for y in sc.YEARS]),
                     legend=dict(orientation="h", y=-0.15), plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(f3, width="stretch")
    tbl = pd.DataFrame({c: {f"FY{y}": pts.get(y, (None,))[0] for y in sc.YEARS} for c, pts in series.items()}).T
    st.dataframe(tbl, width="stretch", column_config={f"FY{y}": st.column_config.NumberColumn(format="%.2f") for y in sc.YEARS})
    missing = [c for c in sc.COMPANIES if c not in series]
    st.info(caveat + (f"\n\nNot in dataset: {', '.join(missing)}." if missing else ""))

# ================================================================ 4. COVERAGE
with tabs[3]:
    st.subheader("How comparable is what these companies disclose?")
    st.caption("Years of data captured for each common measure, out of five (FY2021 to FY2025).")
    rows = []
    for p, name, m in sc.CONCEPTS:
        row = {"Pillar": sc.PILLAR_NAME[p], "Measure": name}
        for c in sc.COMPANIES:
            row[c] = df[(df.Company == c) & df.Indicator.isin(m.get(c, []))].Year.nunique()
        row["Companies"] = sum(1 for c in sc.COMPANIES if row[c])
        rows.append(row)
    mx = pd.DataFrame(rows)
    pillar_bg = {v: sc.PILLAR_COLOUR[k] for k, v in sc.PILLAR_NAME.items()}

    def shade(col):
        out = []
        for v, p in zip(col, mx["Pillar"]):
            if not v:
                out.append("color:#9BADA5")
            else:
                alpha = int(40 + 180 * min(v, 5) / 5)
                out.append(f"background-color:{pillar_bg[p]}{alpha:02x};color:{'white' if v > 2 else '#16302B'};font-weight:600;text-align:center")
        return out

    st.dataframe(mx.style.apply(shade, subset=sc.COMPANIES).format({c: (lambda v: v if v else "–") for c in sc.COMPANIES}, na_rep="–"),
                 hide_index=True, width="stretch", height=(len(mx) + 1) * 35 + 3)
    cov = {c: mx[c].clip(upper=5).sum() / (len(mx) * 5) for c in sc.COMPANIES}
    cols = st.columns(len(sc.COMPANIES))
    for col, c in zip(cols, sc.COMPANIES):
        col.metric(f"{c} coverage", f"{cov[c]:.0%}")

    st.markdown("#### Coverage of the six capitals")
    st.caption("Share of all extracted observations that relate to each capital of the International <IR> Framework.")
    caps = ["Financial", "Manufactured", "Intellectual", "Human", "Social and relationship", "Natural"]
    counts = df.Capital.value_counts()
    cols = st.columns(6)
    for col, k in zip(cols, caps):
        col.metric(k, f"{counts.get(k, 0) / len(df):.0%}", help=f"{counts.get(k, 0)} rows")

    st.markdown("#### Traceability")
    tr = df.groupby("Company").agg(Observations=("Value", "size"), With_page=("Page", lambda s: s.notna().sum()),
                                   First=("Year", "min"), Last=("Year", "max"),
                                   Sources=("Source", lambda s: "; ".join(dict.fromkeys(s)))).reset_index()
    tr["Page-traced"] = tr.With_page / tr.Observations * 100
    tr["Years"] = "FY" + tr.First.astype(str) + "–FY" + tr.Last.astype(str)
    st.dataframe(tr[["Company", "Observations", "With_page", "Page-traced", "Years", "Sources"]].rename(columns={"With_page": "With page number"}),
                 hide_index=True, width="stretch",
                 column_config={"Page-traced": st.column_config.ProgressColumn("Page-traced", min_value=0, max_value=100, format="%.0f%%")})

# ================================================================ 5. DATA
with tabs[4]:
    st.subheader("Data and sources")
    f1, f2_, f3_, f4 = st.columns([1, 1, 2, 1])
    fc = f1.selectbox("Company", ["All"] + sc.COMPANIES, key="dt_c")
    fk = f2_.selectbox("Category", ["All"] + sorted(df.Category.unique()), key="dt_k")
    q = f3_.text_input("Search indicators or notes", key="dt_q")
    np_only = f4.checkbox("Only rows missing a page", key="dt_np")
    v = df.copy()
    if fc != "All":
        v = v[v.Company == fc]
    if fk != "All":
        v = v[v.Category == fk]
    if q:
        v = v[(v.Indicator + " " + v.Notes).str.contains(q, case=False, regex=False)]
    if np_only:
        v = v[v.Page.isna()]
    st.caption(f"{len(v)} of {len(df)} rows")
    out = v[["Company", "FY", "Indicator", "Category", "Value", "Unit", "Source", "Page", "Capital", "Notes"]]
    st.dataframe(out, hide_index=True, width="stretch", height=520)
    st.download_button("Download these rows as CSV", out.to_csv(index=False).encode("utf-8"), "apfa802_filtered.csv", "text/csv")

# ================================================================ 6. METHOD
with tabs[5]:
    a, b = st.columns(2)
    with a:
        st.subheader("How the score works")
        st.markdown(f"""
For each indicator, the change between the baseline year and the latest year is

`change = (latest − baseline) ÷ |baseline|`

and is converted to a score around a neutral 3:

`score = 3 ± change × (2 ÷ threshold)`, capped between 1 and 5.

The sign depends on the preferred direction. Lower emissions, injury rates and debt score above 3 when they fall; revenue, earnings and returns score above 3 when they rise. With the current threshold of {st.session_state.thr}%, an improvement of {st.session_state.thr}% or more scores 5 and an equal deterioration scores 1.

The composite is the weighted average of indicator scores divided by the sum of weights used, so it stays on the 1 to 5 scale even when weights do not add up to 100%.

Indicators with no baseline year score a neutral 3 by default, matching the group's workbook. The sidebar lets you leave them out, which shows how much of a score rests on missing data.

The score measures direction of travel, not level. A company with high but falling emissions can score well on that indicator, so read it alongside the explorer and coverage tabs.
""")
        st.subheader("Six capitals mapping")
        st.markdown("Financial indicators map to financial capital. Emissions, energy, waste and environmental spend map to natural capital. "
                    "Safety, headcount, diversity, employment cost and skills map to human capital. B-BBEE, preferential procurement and "
                    "community investment map to social and relationship capital. Production volume maps to manufactured capital. "
                    "No extracted indicator maps to intellectual capital.")
    with b:
        st.subheader("Data checks to resolve before submission")
        st.caption("Found in the workbook. Each affects either a score or the traceability requirement.")
        for i, (t, d) in enumerate(sc.CHECKS, 1):
            with st.expander(f"{i}. {t}"):
                st.write(d)

st.divider()
st.caption("APFA802 Major Project 2026, Durban University of Technology. Data: company integrated reports FY2021 to FY2025 and Sasol SENS "
           "results announcements. Sappi values converted from US dollars at annual average USD/ZAR rates.")
