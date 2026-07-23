# Power BI Build Guide — Supply Chain Disruption Risk

## Step 1 — Generate the data workbook

```bash
python powerbi/export_for_powerbi.py
```

Produces `powerbi/SupplyChainRisk_PowerBI.xlsx` with seven sheets.

---

## Step 2 — Import into Power BI Desktop

1. **Get Data → Excel Workbook** → select `SupplyChainRisk_PowerBI.xlsx`
2. Tick **all seven sheets** in the Navigator → Load

---

## Step 3 — Set up relationships (Model view)

| From table | From column | To table | To column | Cardinality |
|---|---|---|---|---|
| Monthly_Trends | YearMonth | Shipments | Date *(date)* | One → Many |
| By_Transport_Mode | Transport_Mode | Shipments | Transport_Mode | One → Many |
| By_Product_Category | Product_Category | Shipments | Product_Category | One → Many |
| Route_Summary | Origin_Port | Shipments | Origin_Port | One → Many |

> The `Risk_Scatter` and `Model_Comparison` tables are standalone — no relationship needed.

---

## Step 4 — Add DAX measures

Open `measures.dax`. Add each block as a **New Measure** in the **Shipments** table
(except `SVM Recall`, `MLP Accuracy`, `Recall Lift vs MLP` — add those to `Model_Comparison`).

---

## Step 5 — Build the report pages

### Page 1 — Executive Summary
| Visual | Config |
|---|---|
| KPI card | `Total Shipments` |
| KPI card | `Total Disruptions` |
| KPI card | `Disruption Rate Label` |
| KPI card | `Avg Lead Time (All)` |
| KPI card | `Lead Time Delta` |
| Donut chart | Legend: `Disruption_Label`, Values: `Total Shipments` |

### Page 2 — Disruption by Transport Mode
- **Clustered bar chart**
  - Axis: `By_Transport_Mode[Transport_Mode]` (sort by `disruption_rate_pct` desc)
  - Values: `By_Transport_Mode[disruption_rate_pct]`
  - Title: *Disruption Rate by Transport Mode*
- **Clustered bar chart**
  - Axis: `By_Transport_Mode[Transport_Mode]`
  - Values: `By_Transport_Mode[avg_lead_time]`
  - Title: *Average Lead Time by Mode*

### Page 3 — Disruption by Product Category
- **Horizontal bar chart**
  - Axis: `By_Product_Category[Product_Category]`
  - Values: `By_Product_Category[disruption_rate_pct]`
  - Conditional formatting: gradient Red→Green
- **Table** — all columns from `By_Product_Category`

### Page 4 — Trends Over Time
- **Line chart**
  - X-axis: `Monthly_Trends[YearMonth]`
  - Line 1: `Monthly_Trends[disruption_rate_pct]`  (primary Y)
  - Line 2: `Monthly_Trends[avg_lead_time]`         (secondary Y)
  - Title: *Disruption Rate & Lead Time — Monthly*
- **Area chart**
  - X-axis: `Monthly_Trends[YearMonth]`
  - Values: `Monthly_Trends[total_shipments]`, `Monthly_Trends[disruptions]`

### Page 5 — Route Risk Map
- **Map visual** (requires Bing Maps enabled in Options)
  - Location: `Route_Summary[Origin_Port]`
  - Latitude: `Route_Summary[origin_lat]`
  - Longitude: `Route_Summary[origin_lon]`
  - Size / Colour saturation: `Route_Summary[disruption_rate_pct]`
- **Table** below showing top 10 routes by disruption rate

### Page 6 — Model Performance
- **Table visual** — all columns from `Model_Comparison`
  - Conditional format `Recall_Minority` column: red→green gradient
  - Conditional format `Recommended` column: Yes = green, No = red
- **Clustered bar chart**
  - Axis: `Model_Comparison[Model]`
  - Values: `Model_Comparison[Accuracy_%]`, `Model_Comparison[Recall_Minority]`
  - Title: *Accuracy vs Recall — Why the SVM Wins*
- **KPI card:** `Recall Lift vs MLP`

### Page 7 — Risk Scatter (Feature Relationships)
- **Scatter chart**
  - X-axis: `Risk_Scatter[GSCPI_Zscore]`
  - Y-axis: `Risk_Scatter[Lead_Time_Days_Zscore]`
  - Legend: `Risk_Scatter[Disruption_Label]`
  - Size: `Risk_Scatter[Distance_km]`
  - Title: *GSCPI vs Lead Time — Disrupted vs On-Time*
- **Scatter chart** (second)
  - X-axis: `Risk_Scatter[Gust_Zscore]`
  - Y-axis: `Risk_Scatter[Lead_Time_Days_Zscore]`
  - Legend: `Risk_Scatter[Disruption_Label]`
  - Title: *Weather Severity vs Lead Time*

---

## Step 6 — Slicers (add to every page)

| Slicer | Field |
|---|---|
| Transport Mode | `Shipments[Transport_Mode]` |
| Product Category | `Shipments[Product_Category]` |
| Disruption | `Shipments[Disruption_Label]` |
| Risk Tier | `Shipments[Risk_Tier]` |
| Date Range | `Shipments[Date]` |

---

## Step 7 — Save as .pbix

File → Save As → `SupplyChainRisk.pbix`

Add the `.pbix` file to this repo so reviewers can open it directly in Power BI Desktop.
