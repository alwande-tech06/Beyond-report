# Beyond the Report: ESG Decision Intelligence Dashboard

Streamlit version of the APFA802 (2026) digital artefact. It turns the group's integrated-reporting dataset
into an interactive ESG-adjusted financial risk scorecard for AECI, ArcelorMittal SA, Omnia, Sappi and Sasol.

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
| `app.py` | The Streamlit dashboard (six tabs) |
| `scoring.py` | Scoring model, indicator configuration, comparison and coverage mappings, data checks |
| `data/APFA802_Dataset.xlsx` | Default dataset (the group's workbook) |
| `requirements.txt` | Python dependencies |
| `.streamlit/config.toml` | Theme |

## Updating the data

Either replace `data/APFA802_Dataset.xlsx`, or upload a new workbook from the sidebar. The file needs a
sheet named `Dataset` with the columns Company, Sector, Financial Year, Indicator, Category, Value, Unit,
Source Document, Page Reference and Notes. Save it in Excel so formula cells keep their calculated values.

To change which indicators are scored, their baseline and latest years, directions or weights, edit the
`MODEL` list at the top of `scoring.py`.

## Scoring method

For each indicator: `change = (latest − baseline) / |baseline|`, then
`score = clamp(3 ± change × (2 / threshold), 1, 5)`, with the sign set by the preferred direction.
The composite is `Σ(score × weight) / Σ(weight)`. With the default ±40% threshold this reproduces the
group's Excel scorecard exactly, except Sasol (2.95 here, 2.89 in Excel), because the Excel sheet
hard-codes Sasol's Adjusted EBITDA as n/a.

## Deploying online (optional)

Push the `esg_dashboard` folder to a GitHub repository and deploy it free on Streamlit Community Cloud
(share.streamlit.io), pointing it at `app.py`.
