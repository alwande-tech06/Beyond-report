"""Data loading, ratios and the composite ESG-Financial Decision Index.

Mirrors the group's workbook (APFA802_ESG_Workbook.xlsx):
  * Ratios sheet: margins, returns, leverage and ESG intensities per company-year.
  * ESG Index sheet: six sub-indicators min-max normalised to 0-100 across all
    company-years (worst = 0, best = 100), "lower is better" ones inverted, then
    combined with the weights below. Missing sub-indicators are dropped and the
    remaining weights rescaled, which is what the workbook's formula does.
"""
from __future__ import annotations

import pandas as pd

YEARS = [2021, 2022, 2023, 2024, 2025]

# Workbook column -> short field name
RAW_COLS = {
    "Company": "company", "Ticker": "ticker", "Sector": "sector", "FY": "fy", "Currency": "currency",
    "Revenue (Rm)": "revenue", "Operating Profit (Rm)": "op_profit", "Net Income (Rm)": "net_income",
    "Total Assets (Rm)": "assets", "Total Equity (Rm)": "equity", "Total Debt (Rm)": "debt", "EPS (cents)": "eps",
    "Scope 1 (tCO2e)": "scope1", "Scope 2 (tCO2e)": "scope2", "Total GHG (tCO2e)": "ghg", "Water (m3)": "water",
    "Employees": "employees", "Safety Metric": "safety_metric", "Safety Value": "safety", "Fatalities": "fatalities",
    "B-BBEE Level": "bbee", "Board Women (%)": "board_women", "GHG Target (%)": "ghg_target",
    "GHG Target Year": "ghg_target_year", "CSI Spend (Rm)": "csi",
}

SHORT = {
    "Afrimat Limited": "Afrimat", "AECI Limited": "AECI", "ArcelorMittal South Africa": "ArcelorMittal SA",
    "Omnia Holdings Limited": "Omnia", "Sasol Limited": "Sasol",
}

# Fixed categorical order: colour follows the company, never its rank.
COMPANY_COLOUR = {
    "Afrimat": "#2A78D6",
    "AECI": "#EB6834",
    "ArcelorMittal SA": "#1BAF7A",
    "Omnia": "#EDA100",
    "Sasol": "#E87BA4",
}

# (key, label, source column, lower is better, default weight)
INDEX_PARTS = [
    ("ghg_int", "GHG intensity", "ghg_int", True, 20),
    ("water_int", "Water intensity", "water_int", True, 15),
    ("safety", "Safety rate", "safety", True, 20),
    ("bbee", "B-BBEE level", "bbee", True, 15),
    ("board_women", "Women on board", "board_women", False, 15),
    ("csi_int", "CSI intensity", "csi_int", False, 15),
]
DEFAULT_WEIGHTS = {k: w for k, _, _, _, w in INDEX_PARTS}


def load(source) -> dict[str, pd.DataFrame]:
    raw = pd.read_excel(source, sheet_name="Raw Data")
    missing = set(RAW_COLS) - set(raw.columns)
    if missing:
        raise ValueError(f"The 'Raw Data' sheet is missing: {', '.join(sorted(missing))}")
    raw = raw[list(RAW_COLS)].rename(columns=RAW_COLS)
    raw = raw[pd.to_numeric(raw.fy, errors="coerce").notna() & raw.company.notna()].copy()
    raw["fy"] = raw.fy.astype(int)
    for c in raw.columns:
        if c not in ("company", "ticker", "sector", "currency", "safety_metric", "fy"):
            raw[c] = pd.to_numeric(raw[c], errors="coerce")
    raw["name"] = raw.company.map(SHORT).fillna(raw.company)

    cit = pd.read_excel(source, sheet_name="Source Citations", header=3)
    cit = cit[pd.to_numeric(cit.FY, errors="coerce").notna()].copy()
    cit["FY"] = cit.FY.astype(int)
    cit["Company"] = cit.Company.map(SHORT).fillna(cit.Company)

    tva = pd.read_excel(source, sheet_name="Target vs Actual Check", header=3)
    tva = tva[tva.Company.notna()].copy()
    tva["Company"] = tva.Company.map(SHORT).fillna(tva.Company)

    return {"raw": add_ratios(raw).reset_index(drop=True), "citations": cit.reset_index(drop=True), "targets": tva.reset_index(drop=True)}


def _div(a, b, k=1.0):
    return (a / b * k).where(b.notna() & (b != 0))


def add_ratios(d: pd.DataFrame) -> pd.DataFrame:
    d = d.copy()
    d["op_margin"] = _div(d.op_profit, d.revenue, 100)
    d["net_margin"] = _div(d.net_income, d.revenue, 100)
    d["roa"] = _div(d.net_income, d.assets, 100)
    d["roe"] = _div(d.net_income, d.equity, 100)
    d["de"] = _div(d.debt, d.equity)
    d["ghg_int"] = _div(d.ghg, d.revenue)
    d["water_int"] = _div(d.water, d.revenue)
    d["ghg_per_emp"] = _div(d.ghg, d.employees)
    d["csi_int"] = _div(d.csi, d.revenue, 100)
    d["rev_per_emp"] = _div(d.revenue, d.employees)
    return d


def esg_index(d: pd.DataFrame, weights: dict[str, float]) -> pd.DataFrame:
    """Add a 0-100 sub-score per index part, the weighted index and the within-year rank."""
    d = d.copy()
    num = pd.Series(0.0, index=d.index)
    den = pd.Series(0.0, index=d.index)
    for key, _, col, lower_better, _ in INDEX_PARTS:
        s = d[col]
        lo, hi = s.min(), s.max()
        n = (s - lo) / (hi - lo) * 100 if hi > lo else s * 0 + 50
        if lower_better:
            n = 100 - n
        d[f"n_{key}"] = n
        w = weights.get(key, 0) / 100
        num += (n * w).fillna(0)
        den += w * n.notna()
    d["esg"] = (num / den).where(den > 0)
    d["rank"] = d.groupby("fy").esg.rank(ascending=False, method="min").astype("Int64")
    return d


def contributions(row: pd.Series, weights: dict[str, float]) -> dict[str, float]:
    """Points each sub-indicator adds to one company-year's index (they sum to the score)."""
    avail = {k: weights.get(k, 0) for k, *_ in INDEX_PARTS if pd.notna(row[f"n_{k}"]) and weights.get(k, 0) > 0}
    tot = sum(avail.values())
    return {lbl: row[f"n_{k}"] * avail[k] / tot for k, lbl, *_ in INDEX_PARTS if k in avail} if tot else {}
