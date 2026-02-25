# Consumption Forecast Tool

A lightweight, browser-based forecasting dashboard that turns historical consumption data from Excel into a clear 5‑year outlook.  
Upload a spreadsheet, explore trends per **Product** and **Subtype**, adjust a scaling factor for “what‑if” scenarios, and export consolidated results back to Excel.

---

## Why this tool

- **Fast, local workflow:** runs in the browser—no server setup required.
- **Structured forecasting:** built-in seasonal time-series forecasting (Holt‑Winters) with a safe fallback (linear trend).
- **Decision-ready output:** monthly, quarterly, and yearly rollups plus year‑over‑year deltas.
- **Presentation-friendly UI:** interactive charts and a clean dark/light theme.

---

## Key capabilities

- **Excel import** (`.xlsx`, `.xls`)
- **Automatic parsing** of Product, Subtype, and monthly columns (`MM/YYYY`)
- **Time-series forecasting**
  - Holt‑Winters triple exponential smoothing (seasonal)
  - Fallback to linear regression when history is insufficient
- **5-year forecast horizon**
  - Year 1: **monthly**
  - Years 2–3: **quarterly**
  - Years 4–5: **yearly**
- **Interactive visualization** with Chart.js
- **Adjustable scaling factor** for scenario modeling
- **Excel export** of aggregated results
- **Dark / Light mode**

---

## Quick start

1. Open `index.html` in your browser  
2. Upload your Excel file containing historical consumption values  
3. Select a **Product** and **Subtype**  
4. (Optional) Adjust the **Factor** to simulate higher/lower consumption  
5. Click **Calculate forecast** and review the tables and YoY changes  
6. Export the consolidated results to `prognose.xlsx`

> Tip: Because everything runs locally in your browser, your data stays on your machine.

---

## Expected Excel format

The first columns should represent **Product** and **Subtype** (the app also tries to detect them by header name).  
Monthly values must be provided as separate columns with headers in `MM/YYYY` format.

| Product | Subtype | 01/2024 | 02/2024 | ... |
|--------|---------|--------:|--------:|-----|
| Electricity | Kitchen | 150 | 145 | ... |

Notes:
- Month columns are automatically detected by the `MM/YYYY` pattern.
- Empty cells are ignored.
- Values are parsed as numbers.

---

## Forecast methodology

- **Primary model:** Holt‑Winters triple exponential smoothing (seasonal length = 12)
- **Fallback model:** linear regression trend forecast
- **Non-negative constraint:** forecasts are clipped to 0 to avoid negative consumption

The **Factor** is applied as a multiplicative adjustment on forecast values to support quick scenario analysis.

---

## Output & export

For each Product/Subtype, the app computes:
- Month-by-month forecast for Year 1
- Quarterly totals for Years 2 and 3
- Annual totals for Years 4 and 5
- Year-over-year percentage changes (relative comparisons across periods)

All selected results are stored in an internal results table and can be exported to Excel.
