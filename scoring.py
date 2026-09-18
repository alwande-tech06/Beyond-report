"""ESG-adjusted financial risk scoring model and reference configuration.

Mirrors the group's 'ESG-Financial Risk Scorecard' sheet in APFA802_Dataset.xlsx:
    change = (latest - baseline) / |baseline|
    score  = clamp(3 +/- change * (2 / threshold), 1, 5)
    composite = sum(score * weight) / sum(weight)
"""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

COMPANIES = ["AECI", "ArcelorMittal SA", "Omnia", "Sappi", "Sasol"]
YEARS = [2021, 2022, 2023, 2024, 2025]
PILLAR_OF_CATEGORY = {"Financial": "F", "Environmental": "E", "Social": "S", "Operational": "O"}
PILLAR_NAME = {"F": "Financial", "E": "Environmental", "S": "Social", "O": "Operational"}
PILLAR_COLOUR = {"F": "#2E4A7D", "E": "#3D7A48", "S": "#9C6A12", "O": "#5F6470"}
COMPANY_COLOUR = {
    "AECI": "#3566A8",
    "ArcelorMittal SA": "#A63F2B",
    "Omnia": "#3D8656",
    "Sappi": "#7F4FA0",
    "Sasol": "#B7801A",
}

# (company, display label, dataset indicator, baseline FY, latest FY, preferred direction, weight, pillar)
MODEL = [
    ("AECI", "Total carbon footprint (Scope 1+2)", "Total carbon footprint (Scope 1+2)", 2024, 2025, "down", 0.20, "E"),
    ("AECI", "TRIR (safety)", "TRIR (Total Recordable Injury Rate)", 2024, 2025, "down", 0.15, "S"),
    ("AECI", "EBITDA margin", "EBITDA margin", 2024, 2025, "up", 0.20, "F"),
    ("AECI", "ROIC", "ROIC", 2024, 2025, "up", 0.20, "F"),
    ("AECI", "Female representation, management", "Female representation, top/senior/middle mgmt", 2024, 2025, "up", 0.15, "S"),
    ("AECI", "Net debt / EBITDA", "Net debt/EBITDA", 2024, 2025, "down", 0.10, "F"),
    ("ArcelorMittal SA", "Fatalities", "Employee & contractor fatalities", 2021, 2025, "down", 0.20, "S"),
    ("ArcelorMittal SA", "LTIFR", "LTIFR", 2021, 2025, "down", 0.15, "S"),
    ("ArcelorMittal SA", "CO2 intensity", "CO2 emissions intensity", 2021, 2025, "down", 0.20, "E"),
    ("ArcelorMittal SA", "Return on capital employed", "Return on capital employed", 2021, 2025, "up", 0.25, "F"),
    ("ArcelorMittal SA", "B-BBEE compliance score", "B-BBEE compliance score", 2021, 2025, "up", 0.10, "S"),
    ("ArcelorMittal SA", "Net debt", "Net debt position", 2021, 2025, "down", 0.10, "F"),
    ("Omnia", "Scope 1 GHG emissions", "Scope 1 GHG emissions", 2021, 2025, "down", 0.20, "E"),
    ("Omnia", "Scope 2 GHG emissions", "Scope 2 GHG emissions", 2021, 2025, "down", 0.15, "E"),
    ("Omnia", "Revenue", "Revenue", 2021, 2025, "up", 0.20, "F"),
    ("Omnia", "Headline EPS", "Headline earnings per share", 2021, 2025, "up", 0.25, "F"),
    ("Omnia", "Recordable case rate", "Recordable case rate (RCR)", 2024, 2025, "down", 0.10, "S"),
    ("Omnia", "Renewable energy use", "Renewable energy use", 2024, 2025, "up", 0.10, "E"),
    ("Sappi", "Sales (converted from USD)", "Sales", 2021, 2025, "up", 0.15, "F"),
    ("Sappi", "Adjusted EBITDA (converted)", "Adjusted EBITDA", 2021, 2025, "up", 0.25, "F"),
    ("Sappi", "Headline EPS (converted)", "Headline earnings/(loss) per share", 2021, 2025, "up", 0.25, "F"),
    ("Sappi", "Net asset value per share (converted)", "Net asset value per share", 2021, 2025, "up", 0.15, "F"),
    ("Sappi", "Profit/(loss) for the year (converted)", "Profit/(loss) for the year", 2021, 2025, "up", 0.20, "F"),
    ("Sasol", "GHG reduction vs FY2017 (Intl Chemicals)", "GHG emission reduction vs FY2017 baseline (International Chemicals)", None, 2025, "up", 0.30, "E"),
    ("Sasol", "Fatalities", "Fatalities", 2022, 2024, "down", 0.30, "S"),
    ("Sasol", "Recordable case rate", "Recordable case rate", None, 2024, "down", 0.20, "S"),
    ("Sasol", "External turnover", "External turnover", 2021, 2025, "up", 0.15, "F"),
    ("Sasol", "Adjusted EBITDA", "Adjusted EBITDA", 2021, 2025, "up", 0.20, "F"),
]

DEFAULTS = {"mode": "Group's weights", "wF": 34, "wE": 33, "wS": 33, "thr": 40, "exnb": False}


# ---------------------------------------------------------------- data loading
def load_dataset(source) -> tuple[pd.DataFrame, list[str]]:
    """Read the 'Dataset' sheet and return a tidy frame plus any load warnings."""
    warnings: list[str] = []
    df = pd.read_excel(source, sheet_name="Dataset")
    df = df.rename(columns={
        "Financial Year": "FY", "Source Document": "Source", "Page Reference": "Page",
    })
    required = {"Company", "Sector", "FY", "Indicator", "Category", "Value", "Unit", "Source", "Page", "Notes"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"The Dataset sheet is missing these columns: {', '.join(sorted(missing))}")
    df = df[[c for c in ["Company", "Sector", "FY", "Indicator", "Category", "Value", "Unit", "Source", "Page", "Notes"]]].copy()
    df = df.dropna(subset=["Company", "Indicator", "FY"])
    df["Year"] = df["FY"].astype(str).str.extract(r"(\d{4})").astype(float).astype("Int64")
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce")

    # Formula cells (e.g. AECI Scope 1+2 total) have no cached value if the file was saved
    # outside Excel. Rebuild that total from its components rather than dropping it.
    mask = df["Value"].isna() & df["Indicator"].eq("Total carbon footprint (Scope 1+2)")
    for ix in df[mask].index:
        c, y = df.at[ix, "Company"], df.at[ix, "Year"]
        parts = df[(df.Company == c) & (df.Year == y) & df.Indicator.isin(["Scope 1 GHG emissions", "Scope 2 GHG emissions"])]["Value"]
        if len(parts) == 2 and parts.notna().all():
            df.at[ix, "Value"] = parts.sum()
    bad = df["Value"].isna().sum()
    if bad:
        warnings.append(f"{bad} row(s) have no numeric value and were left out. Open and re-save the file in Excel so formula results are stored.")
        df = df.dropna(subset=["Value"])

    df["Page"] = pd.to_numeric(df["Page"], errors="coerce").astype("Int64")
    df["Notes"] = df["Notes"].fillna("").astype(str).str.strip()
    df["Pillar"] = df["Category"].map(PILLAR_OF_CATEGORY).fillna("O")
    df["Capital"] = df.apply(capital_of, axis=1)
    df["SourceShort"] = df.apply(short_source, axis=1)
    return df.reset_index(drop=True), warnings


def capital_of(r) -> str:
    i = str(r["Indicator"]).lower()
    if r["Category"] == "Financial":
        return "Financial"
    if r["Category"] == "Operational":
        return "Manufactured"
    if r["Category"] == "Environmental":
        return "Natural"
    if any(k in i for k in ("b-bbee", "procurement", "social responsibility", "socio-economic")):
        return "Social and relationship"
    return "Human"


def short_source(r) -> str:
    s = str(r["Source"]).replace("Integrated Report", "IR").replace("Holdings ", "")
    if s.startswith("Sasol Limited audited results"):
        s = "Sasol SENS results"
    return s + (f", p.{int(r['Page'])}" if pd.notna(r["Page"]) else "")


# ---------------------------------------------------------------- scoring
@dataclass
class CompanyResult:
    company: str
    items: pd.DataFrame
    composite: float | None
    pillars: dict = field(default_factory=dict)
    weight_total: float = 0.0
    group_weight_sum: float = 0.0
    neutral_share: float = 0.0
    flags: list = field(default_factory=list)


def compute(df: pd.DataFrame, mode: str, pillar_w: dict, thr: float, exclude_no_baseline: bool) -> dict[str, CompanyResult]:
    look = {(r.Company, r.Indicator, int(r.Year)): r for r in df.itertuples(index=False)}
    k = 2 / (thr / 100)
    results = {}
    for c in COMPANIES:
        rows = []
        for (co, label, ind, b, l, d, w, p) in MODEL:
            if co != c:
                continue
            br = look.get((c, ind, b)) if b is not None else None
            lr = look.get((c, ind, l))
            ch = (lr.Value - br.Value) / abs(br.Value) if (br is not None and lr is not None and br.Value != 0) else None
            sgn = 1 if d == "up" else -1
            score = 3.0 if ch is None else min(5.0, max(1.0, 3 + sgn * ch * k))
            srcs = list(dict.fromkeys([x.SourceShort for x in (br, lr) if x is not None]))
            rows.append({
                "Indicator": label, "Pillar": p, "Direction": "Higher is better" if d == "up" else "Lower is better",
                "Unit": lr.Unit if lr is not None else "", "Baseline FY": b, "Baseline": br.Value if br is not None else None,
                "Latest FY": l, "Latest": lr.Value if lr is not None else None, "Change": ch, "Score": score,
                "GroupWeight": w, "Neutral": ch is None, "Source": "; ".join(srcs),
            })
        it = pd.DataFrame(rows)
        if it.empty:
            results[c] = CompanyResult(c, it, None)
            continue
        it["Included"] = ~(exclude_no_baseline & it["Neutral"])
        if mode == "Group's weights":
            it["Weight"] = it["GroupWeight"].where(it["Included"], 0.0)
        else:
            n = it[it.Included].groupby("Pillar").size()
            it["Weight"] = it.apply(lambda r: (pillar_w[r.Pillar] / 100) / n[r.Pillar] if r.Included and r.Pillar in n else 0.0, axis=1)
        W = it["Weight"].sum()
        comp = (it.Score * it.Weight).sum() / W if W > 0 else None
        it["Weight share"] = it["Weight"] / W if W > 0 else 0.0
        it["Contribution"] = it.Score * it["Weight share"]
        pil = {}
        for p in "FES":
            g = it[(it.Pillar == p) & (it.Weight > 0)]
            pil[p] = (g.Score * g.Weight).sum() / g.Weight.sum() if len(g) else None
        neutral_share = it.loc[it.Neutral & (it.Weight > 0), "Weight"].sum() / W if W > 0 else 0.0
        res = CompanyResult(c, it, comp, pil, W, it["GroupWeight"].sum(), neutral_share)
        res.flags = flags_for(res, mode)
        results[c] = res
    return results


def flags_for(r: CompanyResult, mode: str) -> list[str]:
    it, f = r.items, []
    has = lambda p: bool(((it.Pillar == p) & (it.Weight > 0)).any())
    if not has("E") and not has("S"):
        f.append("No environmental or social indicators in the model. The composite is financial-only, so it is not yet ESG-adjusted for this company.")
    else:
        if not has("E"):
            f.append("No environmental indicator is scored.")
        if not has("S"):
            f.append("No social indicator is scored.")
    nn = int((it.Neutral & (it.Weight > 0)).sum())
    if nn:
        f.append(f"{nn} indicator{'s have' if nn > 1 else ' has'} no baseline and default{'' if nn > 1 else 's'} to a neutral 3 "
                 f"({r.neutral_share:.0%} of this company's weight).")
    spans = sorted({int(l - b) for b, l in zip(it["Baseline FY"], it["Latest FY"]) if pd.notna(b)})
    if len(spans) > 1:
        f.append("Indicators are compared over different periods (" + " and ".join(f"{s}-year" for s in spans) + " changes), which weakens comparability.")
    elif spans == [1]:
        f.append("Every indicator is a one-year change (FY2024 to FY2025), while most other companies are measured over four years. "
                 "The four-year AECI series in the dataset could be used instead.")
    if mode == "Group's weights" and abs(r.group_weight_sum - 1) > 0.001:
        f.append(f"The group's weights for this company add up to {r.group_weight_sum:.0%}, not 100%. The composite is rescaled, but the weights should be fixed in the workbook.")
    if r.company == "ArcelorMittal SA":
        f.append("Check whether the B-BBEE figure is a contributor level (1 is best, 8 is lowest). If so, the move from 8 to 4 is an improvement and should score above 3, not 1.")
    if r.company == "Sasol":
        f.append("The Excel scorecard hard-codes Adjusted EBITDA as n/a with a neutral 3. The dataset has both years, so this dashboard scores the actual change.")
    return f


# ---------------------------------------------------------------- comparison and coverage config
# measure -> (unit, {company: [(indicator, multiplier), ...]}, caveat). "__S12" = Scope 1 + Scope 2.
COMPARE = {
    "Revenue / sales": ("Rm", {"AECI": [("Revenue", 1)], "Omnia": [("Revenue", 1)], "Sappi": [("Sales", 1)], "Sasol": [("External turnover", 1)]},
        "Sappi reports in US dollars; values were converted at each year's average USD/ZAR rate, so part of its Rand growth reflects currency weakness. Sasol figures come from SENS announcements, not its integrated reports."),
    "EBITDA": ("Rm", {"AECI": [("EBITDA", 1), ("EBITDA (core)", 1)], "Sappi": [("Adjusted EBITDA", 1)], "Sasol": [("Adjusted EBITDA", 1)]},
        "AECI's FY2022–FY2023 values are core EBITDA from the strategy KPI table; FY2024–FY2025 use reported EBITDA. Sappi and Sasol report adjusted EBITDA, each on its own definition. Sasol FY2023 is approximate, per the dataset note."),
    "Headline earnings per share": ("cents", {"AECI": [("HEPS", 1)], "Omnia": [("Headline earnings per share", 1)], "Sappi": [("Headline earnings/(loss) per share", 1)], "Sasol": [("Headline earnings per share", 100)]},
        "Sasol HEPS converted from Rand to cents. Per-share figures depend on share count, so compare direction rather than size across companies."),
    "Net debt": ("Rm", {"AECI": [("Net debt", 1000)], "ArcelorMittal SA": [("Net debt position", 1)], "Sasol": [("Net debt (excl. leases)", 1)]},
        "AECI converted from R billion. Sasol excludes leases; the others do not say, so definitions may differ."),
    "Return on capital": ("%", {"AECI": [("ROIC (strategy KPI)", 1)], "ArcelorMittal SA": [("Return on capital employed", 1)]},
        "AECI reports ROIC; ArcelorMittal SA reports ROCE. They use different numerators and capital bases."),
    "Scope 1 + 2 GHG emissions": ("tCO2e", {"AECI": [("Total carbon footprint (Scope 1+2)", 1)], "Omnia": [("__S12", 1)]},
        "Omnia's total is the sum of its reported Scope 1 and Scope 2. Only two companies disclose absolute emissions in the dataset, a major comparability gap."),
    "Scope 1 GHG emissions": ("tCO2e", {"AECI": [("Scope 1 GHG emissions", 1)], "Omnia": [("Scope 1 GHG emissions", 1)]},
        "Omnia reports tonnes CO2; AECI reports tCO2e. If Omnia excludes non-CO2 gases, it is understated relative to AECI."),
    "Renewable energy": ("MWh", {"AECI": [("Renewable electricity", 1)], "Omnia": [("Renewable energy use", 1)]},
        "AECI reports renewable electricity only; Omnia reports renewable energy use."),
    "Fatalities": ("number", {"AECI": [("Fatalities", 1)], "ArcelorMittal SA": [("Employee & contractor fatalities", 1)], "Sasol": [("Fatalities", 1)]},
        "Sasol's FY2024 figure includes one fatality shortly after year-end."),
    "Injury rate": ("rate", {"AECI": [("TRIR (Total Recordable Injury Rate)", 1)], "ArcelorMittal SA": [("TIFR", 1)], "Omnia": [("Recordable case rate (RCR)", 1)], "Sasol": [("Recordable case rate", 1)]},
        "Not directly comparable. Omnia uses 200,000 hours worked; the others do not state their base in the dataset. ArcelorMittal SA's TIFR runs far higher, which usually means a larger hours base (for example 1,000,000 hours)."),
    "Community investment": ("Rm", {"AECI": [("Social responsibility spend", 1)], "Sasol": [("Investment in socio-economic development", 1)]},
        "Sasol is far larger than AECI, so scale spend by revenue before comparing."),
}

CONCEPTS = [
    ("F", "Revenue / sales", {"AECI": ["Revenue"], "Omnia": ["Revenue"], "Sappi": ["Sales"], "Sasol": ["External turnover"]}),
    ("F", "EBITDA", {"AECI": ["EBITDA", "EBITDA (core)"], "ArcelorMittal SA": ["EBITDA per tonne sold"], "Sappi": ["Adjusted EBITDA"], "Sasol": ["Adjusted EBITDA"]}),
    ("F", "Operating profit", {"Omnia": ["Operating profit"], "Sappi": ["Operating profit excl. special items"]}),
    ("F", "Headline EPS", {"AECI": ["HEPS"], "Omnia": ["Headline earnings per share"], "Sappi": ["Headline earnings/(loss) per share"], "Sasol": ["Headline earnings per share"]}),
    ("F", "Net debt", {"AECI": ["Net debt"], "ArcelorMittal SA": ["Net debt position"], "Sasol": ["Net debt (excl. leases)"]}),
    ("F", "Return on capital", {"AECI": ["ROIC", "ROIC (strategy KPI)"], "ArcelorMittal SA": ["Return on capital employed"]}),
    ("E", "Scope 1 emissions", {"AECI": ["Scope 1 GHG emissions"], "Omnia": ["Scope 1 GHG emissions"]}),
    ("E", "Scope 2 emissions", {"AECI": ["Scope 2 GHG emissions"], "Omnia": ["Scope 2 GHG emissions"]}),
    ("E", "Scope 3 emissions", {"AECI": ["Scope 3 GHG emissions"]}),
    ("E", "Emissions intensity or reduction", {"ArcelorMittal SA": ["CO2 emissions intensity"], "Sasol": ["GHG emission reduction vs FY2017 baseline (Group)", "GHG emission reduction vs FY2017 baseline (International Chemicals)"]}),
    ("E", "Renewable energy", {"AECI": ["Renewable electricity"], "Omnia": ["Renewable energy use"]}),
    ("E", "Waste", {"AECI": ["Hazardous waste disposed"]}),
    ("E", "Environmental spend", {"ArcelorMittal SA": ["Environmental spend"]}),
    ("S", "Fatalities", {"AECI": ["Fatalities"], "ArcelorMittal SA": ["Employee & contractor fatalities"], "Sasol": ["Fatalities"]}),
    ("S", "Injury rate", {"AECI": ["TRIR (Total Recordable Injury Rate)"], "ArcelorMittal SA": ["LTIFR", "TIFR"], "Omnia": ["Recordable case rate (RCR)"], "Sasol": ["Recordable case rate"]}),
    ("S", "Women in management or workforce", {"AECI": ["Female representation, top/senior/middle mgmt"], "Sasol": ["Gender: female representation (South Africa)"]}),
    ("S", "Headcount", {"AECI": ["Total employees"], "Omnia": ["Total permanent employees (SA+Africa+Other)"]}),
    ("S", "Employment cost or skills", {"ArcelorMittal SA": ["Total cost of employment"], "Sasol": ["Investment in skills development"]}),
    ("S", "B-BBEE or preferential procurement", {"ArcelorMittal SA": ["B-BBEE compliance score", "Preferential procurement spend"]}),
    ("S", "Community investment", {"AECI": ["Social responsibility spend"], "Sasol": ["Investment in socio-economic development"]}),
]

CHECKS = [
    ("ArcelorMittal SA B-BBEE direction", "The scorecard treats a higher B-BBEE figure as better, so 8 to 4 scores the minimum of 1. If the report shows a contributor level, where Level 1 is best, this is an improvement and the direction should be flipped. Confirm on p.9 of the 2025 report."),
    ("Sasol Adjusted EBITDA is hard-coded", "The Excel scorecard enters n/a and a neutral 3 for this row, although FY2021 (R48,420m) and FY2025 (R51,764m) are both in the dataset. The actual change is +6.9%. Replace the hard-coded cells with the same formula used in other rows."),
    ("Sasol weights add up to 115%", "Weights of 30%, 30%, 20%, 15% and 20% sum to 1.15. The composite formula rescales them, so the score is still on 1 to 5, but the stated weights are misleading. Reduce them to 100%."),
    ("AECI baseline years do not match the notes", "The formulas compare FY2024 with FY2025 for EBITDA margin, ROIC and net debt/EBITDA, but the notes say an FY2022 baseline was used. The dataset has a four-year series (FY2022 to FY2025) for core EBITDA, ROIC and net debt/EBITDA from the strategy KPI table."),
    ("Column headings say FY2021 and FY2025", "Several rows use other years (AECI FY2024, Omnia FY2024 for safety and renewables, Sasol FY2022 to FY2024 for fatalities). Rename the headings to Baseline and Latest and add a column showing the years used."),
    ("Sappi has no ESG indicators", "All five Sappi indicators are financial, so its composite is not ESG-adjusted. Sappi publishes a sustainability report with emissions and safety data. Adding even Scope 1+2 and a safety rate would make the comparison fair."),
    ("Exchange rate source is inconsistent", "The data dictionary cites Nedbank Group Economic Unit rates, but the Sappi notes in the scorecard say OECD/FRED. Use one source and cite it the same way everywhere."),
    ("Sasol rows without a page reference", "The SENS results announcements have no page numbers. Add the announcement date and URL in a source column so the rows remain traceable, and explain in the methodology why data outside the integrated report was needed."),
    ("Injury rates use different bases", "TRIR, TIFR, LTIFR and RCR are measured differently across companies. The scorecard uses change over time within each company, which is sound, but the report should not compare levels across companies."),
]
