# Beyond the Report: ESG Decision Intelligence Dashboard

Streamlit version of the APFA802 (2026) digital artefact. It turns the group's integrated-reporting workbook
into a single-page ESG decision dashboard for Afrimat, AECI, ArcelorMittal SA, Omnia and Sasol (FY2021 to FY2025).

## Run it

Requires Python 3.10 or newer.

```bash
cd esg_dashboard
python -m venv .venv
# Windows:      .venv\Scripts\activate
# macOS/Linux:  source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

The dashboard opens at http://localhost:8501.

## Files

| File | Purpose |
|---|---|
| `app.py` | The dashboard: KPI cards, ESG index ranking and breakdown, year-by-year and financial tables, five-year trends, target-vs-actual check, correlations, source citations |
| `esg.py` | Workbook loading, financial and ESG ratios, and the composite ESG index |
| `data/APFA802_ESG_Workbook.xlsx` | Default dataset (the group's workbook) |
| `requirements.txt` | Python dependencies |
| `.streamlit/config.toml` | Theme |

## Updating the data

Replace `data/APFA802_ESG_Workbook.xlsx`, or upload a workbook under *Data and sources → Raw data*. It needs the
sheets `Raw Data`, `Source Citations` and `Target vs Actual Check` in the same layout. Ratios and the ESG index
are recalculated from `Raw Data`, so only the raw figures have to be right.

## ESG index

Six sub-indicators (GHG intensity, water intensity, safety rate, B-BBEE level, women on board, CSI intensity) are
min-max scaled to 0–100 across all 25 company-years, with lower-is-better measures inverted, then weighted
20/15/20/15/15/15. A missing value is skipped and the remaining weights rescaled. This reproduces the workbook's
ESG Index sheet exactly. Weights can be changed under *Data and sources → Index weights*.

## Deploying online (optional)

Push the `esg_dashboard` folder to a GitHub repository and deploy it free on Streamlit Community Cloud
(share.streamlit.io), pointing it at `app.py`.
