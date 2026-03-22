# 🔵 Cisco Forecast League (CFL) — Phase 1
### Demand Forecasting for 30 Cisco Products | FY26 Q2

> **Target:** Forecast quarterly demand for 30 critical Cisco networking products  
> **Submission:** FY26 Q2 (November 2025 – January 2026)  
> **Tool:** R

---

## 📂 Repository Structure

```
cfl-demand-forecast/
│
├── CFL_Forecast_R_Script_v2.R     # Main R forecasting script
└── README.md                      # This file
```

---

## 🧠 Methodology

Aligned with **Cisco Forecasting 101** (March 2026) — Slides 5, 8, 9, 11, 12.

> *"Forecasting is both an ART and a SCIENCE"* — Slide 9

### Step 1 — Statistical Model (40% weight)

Each product is first classified into one of Cisco's 6 time-series types (Slide 11):

| Series Type | Classification Rule | Method Used |
|---|---|---|
| **Stable** | Low trend slope, low CV | 40% Seasonal Index + 35% 4-Qtr Avg + 25% Holt's ES |
| **Growing** | Relative slope > 4% | 40% Holt's ES + 35% Seasonal Index + 25% Linear Trend |
| **Declining** | Relative slope < −4% OR PLC=Decline | 55% Linear Trend + 45% SES (dampened) |
| **Seasonal** | High Q2 vs Q3 spread (>15%) | 50% Seasonal Index + 30% Holt's ES + 20% 4-Qtr Avg |
| **Volatile** | Coefficient of Variation > 0.35 | 50% SES (α=0.35) + 50% 4-Qtr Avg |
| **NPI** | PLC=NPI-Ramp or ≤3 quarters of history | 50% Holt's ES (aggressive) + 50% Linear Trend |

**Core techniques (Cisco Slide 5 & 12):**
- **4-Quarter Average** — Cisco baseline; avg of last 4 actuals
- **Seasonal Index** — OLS trend fitted through same-quarter values (FY23Q2 → FY24Q2 → FY25Q2), extrapolated to FY26Q2
- **Holt's Double ES** — regression-initialised trend; captures level + slope
- **Linear Extrapolation** — OLS for declining/NPI products
- **SES** — Simple exponential smoothing for volatile/noisy series

### Step 2 — Reference Ensemble (60% weight)

Three reference forecasts are provided by Cisco teams. Each is weighted by its **mean historical accuracy** over the last 3 quarters (FY25Q3, FY25Q4, FY26Q1):

| Team | Avg Accuracy | Weight |
|---|---|---|
| Demand Planner (DP) | ~83.9% | **35.7%** |
| Data Science (DS) | ~77.6% | **33.1%** |
| Marketing (MKT) | ~73.2% | **31.2%** |

```
Ref_Ensemble = w_DP × DP_Forecast + w_MKT × MKT_Forecast + w_DS × DS_Forecast
```

### Step 3 — Final Blend

```
My_Forecast = 0.40 × Stat_Model + 0.60 × Ref_Ensemble
```

The 60/40 lean towards the reference ensemble reflects Slide 9's "Art + Science" framing — the DS team incorporates macroeconomic features and the DP team has domain context unavailable in raw bookings data.

---

## 📊 Accuracy Metric

Per Cisco's evaluation criteria (Slide 8):

```
Bias     = (Forecast − Actual) / Actual
Accuracy = max(0, 1 − |Bias|)
```

---

## 🚀 How to Run

### Prerequisites
```r
install.packages(c("readxl", "dplyr", "tidyr", "ggplot2",
                   "scales", "patchwork", "openxlsx", "RColorBrewer"))
```

### Steps
1. Place `CFL_Forecast_R_Script_v2.R` and `CFL_External_Data_Pack_Phase1.xlsx` in the same folder
2. Open the R script and update `DATA_FILE` path if needed
3. Run the script — outputs generated:

| Output File | Description |
|---|---|
| `CFL_My_Forecasts_FY26Q2.xlsx` | Submission-ready Excel (3 sheets: forecast, full data, methodology) |
| `CFL_Top9_TimeSeries.png` | Time-series plots for top 9 products |
| `CFL_Forecast_Comparison.png` | All 30 products — 4-way bar chart comparison |
| `CFL_Accuracy_Heatmap.png` | Team accuracy heatmap (last 3 quarters) |
| `CFL_Portfolio_Overview.png` | Series-type taxonomy + top 10 by volume |

---

## 📦 Data Pack (not included)

The `CFL_External_Data_Pack_Phase1.xlsx` file is Cisco Confidential and not committed to this repo. It contains:
- **Data Pack - Actual Bookings** — 12 quarters of historical demand (FY23Q2–FY26Q1) for 30 products
- **Big Deal / SCMS / VMS** — supplementary segmentation data
- **Reference forecasts** — DP, Marketing, and Data Science team forecasts for FY26Q2

---

## 🏷️ Product Coverage

30 Cisco products spanning:
- Enterprise & Core Switching (48-port, fiber, PoE, UPOE variants)
- Wireless Access Points (WiFi6 / WiFi6E, indoor/outdoor, internal/external antenna)
- Routers (branch, edge, core)
- Security Firewalls (Next-Gen)
- IP Phones & Conference Phones
- Industrial Switches
- Data Center Switching

---

*Submitted as part of Cisco Forecast League (CFL) Phase 1 — March 2026*
