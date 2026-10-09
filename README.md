# 📊 PR.2 Analyzer — Sales Analytics Workbook

### *Conditional Formatting · What-If · Regression · Pivot Tables · Dashboard*

> *"Data is only as useful as the story you can tell with it."*

**Author:** Gosai Kishal  |  **Date:** 09 October, 2026

---

## 📋 Table of Contents

* [📌 Overview](#-overview)
* [🎯 Problem Statement](#-problem-statement)
* [✨ Key Features](#-key-features)
* [🏗️ Workbook Structure](#️-workbook-structure)
* [🔄 Project Workflow](#-project-workflow)
* [📥 Part A — Dataset Sheet](#-part-a--dataset-sheet)
* [📊 Part B — Pivot Table Sheet](#-part-b--pivot-table-sheet)
* [📈 Part C — Dashboard Sheet](#-part-c--dashboard-sheet)
* [🛠️ Tools & Excel Features Used](#️-tools--excel-features-used)
* [📈 Results & Insights](#-results--insights)
* [⚠️ Notes & Assumptions](#️-notes--assumptions)
* [🏆 Advantages](#-advantages)
* [👤 Author](#-author)

---

## 📌 Overview

**PR.2 Analyzer** is an Excel workbook that analyses a sales dataset of **200 orders** (10-Apr-2024 to 07-Apr-2025) covering 30 customers, 5 regions and 5 product categories. It completes every task listed on the *Project Instructions* sheet:

* Highlighting top customers with **conditional formatting**
* Running a **what-if analysis** on discount vs profit
* Performing **linear regression**, **descriptive statistics** and a **histogram** in Analysis ToolPak format
* Showing **▲ / ▼ arrows** on monthly sales growth with custom number formats
* Adding a **timestamp** column with `NOW()`
* Finding **high-value customers** with `INDEX()`, `MATCH()` and filters
* Building **PivotTables** and a **dashboard with bar, line and pie charts**

---

## 🎯 Problem Statement

> **Objective:** Analyse the sales dataset, uncover insights and present them in a clean dashboard.

| #  | Task (from the instruction sheet)                           | Where it is done     |
| -- | ----------------------------------------------------------- | -------------------- |
| 1  | Conditional formatting — top 10 customers by total purchase | Dataset, Pivot Table |
| 2  | What-if analysis — impact of discount on total profit       | Dataset              |
| 3  | Linear regression (Profit vs Sales)                         | Dataset              |
| 4  | Descriptive statistics                                      | Dataset              |
| 5  | Up/down arrows on monthly sales growth                      | Pivot Table          |
| 6  | Timestamp column using `NOW()`                              | Dataset              |
| 7  | High-value customers using `INDEX()`, `MATCH()` and filters | Pivot Table          |
| 8  | Pivot Table — total sales by region and product             | Pivot Table          |
| 9  | Bar, line and pie charts for KPIs                           | Dashboard            |
| 10 | Summary of insights in a dashboard                          | Dashboard            |

---

## ✨ Key Features

| Feature                     | Description                                                            |
| --------------------------- | ---------------------------------------------------------------------- |
| 🏅 **Top 10 Highlighting**  | Every order from the top 10 customers is shaded gold and bold          |
| 🎛️ **What-If Scenario**    | Type any discount and see sales, profit, margin and change update live |
| 🎯 **Goal Seek Cell**       | Shows the discount needed to reach a target profit                     |
| 📉 **Regression Output**    | Summary output, ANOVA, coefficients and residuals for Profit vs Sales  |
| 📋 **Descriptive Stats**    | 14 statistics for Sales, Quantity, Discount and Profit                 |
| 📊 **Histogram**            | Frequency table plus chart for order sales                             |
| ▲▼ **Growth Arrows**        | Custom number format shows ▲ or ▼ with green/red colour                |
| 🕒 **Timestamp**            | `NOW()` report timestamp on every row                                  |
| 👑 **High-Value Customers** | Ranked list built with `INDEX` and `MATCH`                             |
| 🧮 **Native PivotTables**   | Two real Excel PivotTable objects                                      |
| 🖥️ **Dashboard**           | 12 KPI cards, 6 charts and 7 auto-updating insight sentences           |

---

## 🏗️ Workbook Structure

```text
📦 PR2_Analyzer_Completed.xlsx
│
├── 📄 Project Instructions   ← Topics and the 10 main tasks (original sheet)
├── 📄 Dataset                ← Data + calculated columns + analysis panels
├── 📄 Pivot Table            ← Pivot summaries, customer analysis, native PivotTables
└── 📄 Dashboard              ← KPI cards, charts and written insights
```

**Colour code used in the sheets**

| Style               | Meaning                                                     |
| ------------------- | ----------------------------------------------------------- |
| Navy header         | Original dataset columns / section titles                   |
| Blue header         | Columns and tables added during analysis                    |
| Grey cells          | Formulas                                                    |
| Blue text on yellow | Inputs you can change (discount, target profit, bins, etc.) |

---

## 🔄 Project Workflow

```text
        Raw Dataset (200 orders)
                 │
                 ▼
   ┌──────────────────────────────┐
   │  Dataset sheet               │
   │  • Calculated columns J–R    │
   │  • What-If · Stats ·         │
   │    Regression · Histogram    │
   └──────────────┬───────────────┘
                  ▼
   ┌──────────────────────────────┐
   │  Pivot Table sheet           │
   │  • Region × Product          │
   │  • Monthly trend + arrows    │
   │  • Customer ranking / tiers  │
   │  • Native PivotTables        │
   └──────────────┬───────────────┘
                  ▼
   ┌──────────────────────────────┐
   │  Dashboard sheet             │
   │  • KPI cards                 │
   │  • Bar · Pie · Line charts   │
   │  • Written insights          │
   └──────────────────────────────┘
```

---

## 📥 Part A — Dataset Sheet

### 🗂️ 1. Source Columns (A–I)

`Customer_ID` · `Customer_Name` · `Region` · `Product_Category` · `Sales` · `Quantity` · `Discount` · `Order_Date` · `Profit`

### ➕ 2. Calculated Columns (J–R)

| Column | Name                       | Purpose                                         |
| ------ | -------------------------- | ----------------------------------------------- |
| J      | Customer Total Purchase    | `SUMIFS` of sales per customer                  |
| K      | Top 10 Customer            | Flags customers ranked 1–10 (`INDEX` / `MATCH`) |
| L      | Gross Sales (pre-discount) | `Sales / (1 − Discount)`                        |
| M      | Scenario Sales             | Gross sales after the scenario discount         |
| N      | Scenario Profit            | Scenario sales minus cost                       |
| O      | Profit Margin %            | Profit ÷ Sales (colour scale)                   |
| P      | Order Month                | First day of the order month                    |
| Q      | Customer Tier              | "High Value" or "Standard"                      |
| R      | Report Timestamp           | `=NOW()` formatted `yyyy-mm-dd hh:mm:ss`        |

Filters are switched on for the header row. A **JUMP TO** list at cell **AA1** links to each analysis block below.

### 🎛️ 3. What-If Analysis (cell T1)

> Change the scenario discount (default 10%) and every figure updates.

* **Current vs Scenario table:** average discount, total sales, cost (held constant), total profit and margin with ▲/▼ change
* **Goal-seek cell:** discount required for a target profit
* **Sensitivity table:** 0% to 25% discount in 5% steps
* **Chart:** Total Profit vs Discount with the selected scenario marked

### 📋 4. Descriptive Statistics (cell T42)

Mean, Standard Error, Median, Mode, Standard Deviation, Sample Variance, Kurtosis, Skewness, Range, Minimum, Maximum, Sum, Count and Confidence Level (95.0%) for **Sales, Quantity, Discount and Profit**.

### 📉 5. Linear Regression — Profit (Y) vs Sales (X) (cell T60)

| Block                  | Contents                                                                   |
| ---------------------- | -------------------------------------------------------------------------- |
| Regression Statistics  | Multiple R, R Square, Adjusted R Square, Standard Error, Observations      |
| ANOVA                  | df, SS, MS, F, Significance F                                              |
| Coefficients           | Intercept and Sales slope with standard error, t Stat, P-value, 95% limits |
| Residual Output (AD60) | Predicted profit, residuals and standard residuals for all 200 rows        |
| Prediction (T80)       | Enter a sales amount to get predicted profit                               |

### 📊 6. Histogram (cell T85)

Bins of 200 from 400 to 2,000 plus "More", with frequency, cumulative % and a column chart.

---

## 📊 Part B — Pivot Table Sheet

| # | Section                   | Description                                                               |
| - | ------------------------- | ------------------------------------------------------------------------- |
| 1 | Region × Product Category | Total sales cross-tab with grand totals and heat-map colouring            |
| 2 | Region Summary            | Orders, sales, profit, margin, average discount                           |
| 3 | Product Category Summary  | Same measures per category                                                |
| 4 | Monthly Sales Trend       | Sales, profit, orders, **MoM change ▲▼** and **MoM growth % ▲▼**          |
| 5 | Customer Analysis         | `SUMIFS` group-by for 30 customers with rank and tier, filters on         |
| 6 | High-Value Customers      | Ranked list pulled with `INDEX` / `MATCH`                                 |
| 7 | Top 10 Customers          | Rank, ID, name, sales, profit (feeds the dashboard bar chart)             |
| 8 | **Native PivotTables**    | Region × Product Category (sales) and Product Category (sales and profit) |

**Arrow format used for growth:**

```text
▲ 0.0%;▼ 0.0%;0.0%
```

Green and red colours come from conditional formatting on positive and negative values.

---

## 📈 Part C — Dashboard Sheet

**12 KPI cards:** Total Sales · Total Profit · Profit Margin · Total Orders · Avg Order Value · High-Value Customers · Avg Discount · Latest MoM Sales Growth · Top Region · Top Product Category · Profit vs Sales R² · Top 10 Customer Share

**Charts**

| Chart | Type                | KPI shown                          |
| ----- | ------------------- | ---------------------------------- |
| 1     | Bar                 | Total Sales by Region              |
| 2     | Pie                 | Sales Share by Product Category    |
| 3     | Line                | Monthly Sales & Profit Trend       |
| 4     | Bar                 | Top 10 Customers by Total Purchase |
| 5     | Column              | Histogram of order sales           |
| 6     | Scatter + trendline | Regression — Profit vs Sales       |

**Key Insights block:** seven sentences written with formulas, so they change when the data changes.

---

## 🛠️ Tools & Excel Features Used

| Tool / Feature                       | Purpose                                                    |
| ------------------------------------ | ---------------------------------------------------------- |
| `SUMIFS` / `COUNTIFS` / `AVERAGEIFS` | Group-by totals and counts                                 |
| `INDEX` + `MATCH`                    | Customer rank lookup and high-value list                   |
| `RANK`, `PERCENTILE`                 | Customer ranking and the high-value threshold              |
| `NOW()`, `EDATE`, `DATE`, `TEXT`     | Timestamp and month handling                               |
| Conditional Formatting               | Top 10 highlight, colour scales, data bars, growth colours |
| Custom Number Formats                | ▲ / ▼ arrows on growth values                              |
| PivotTables                          | Region × Product and Product summaries                     |
| Charts                               | Bar, line, pie, histogram, scatter with trendline          |
| Data Validation                      | Input limits on the scenario discount                      |
| AutoFilter                           | Filtering on Dataset and customer table                    |

---

## 📈 Results & Insights

| Metric                  | Result                                         |
| ----------------------- | ---------------------------------------------- |
| Total sales             | **195,218**                                    |
| Total profit            | **68,287** (margin **35.0%**)                  |
| Orders                  | **200**                                        |
| Top region              | **West** — 44,084 (22.6% of sales)             |
| Lowest region           | **East** — 32,442                              |
| Top product category    | **Books** — 24.1% of sales                     |
| Lowest product category | **Office Supplies** — 14.7%                    |
| Best month              | **Sep-2024** — 20,534                          |
| Top 10 customers        | **47.3%** of total sales                       |
| High-value customers    | **8** (total sales ≥ 8,236)                    |
| Regression              | Profit ≈ −1.03 + 0.351 × Sales, **R² = 0.598** |
| What-if (10% discount)  | Profit changes by **+1,007 (+1.5%)**           |
| Target profit of 80,000 | Needs a discount of about **5.1%**             |

---

## ⚠️ Notes & Assumptions

* **ToolPak outputs are values.** Descriptive statistics, regression and histogram tables follow the Analysis ToolPak layout, and their values were computed outside Excel and cross-checked. To regenerate them with Excel itself, use *Data → Data Analysis*.
* **Customer names are inconsistent in the source data.** The same `Customer_ID` appears under several names, so customers are grouped by ID and the first recorded name is shown.
* **April 2025 is a partial month** (7 days of data), so its growth figure of ▼ 81.8% is understated.
* **What-if model:** cost (Sales − Profit) is held constant and gross sales are rebuilt from the recorded discount.
* **High-value customer:** total sales in the top quarter of customers (75th percentile).
* **Timestamps** use `NOW()`, so they refresh each time the workbook recalculates.

---

## 🏆 Advantages

| Advantage                   | Detail                                                                |
| --------------------------- | --------------------------------------------------------------------- |
| 🔄 **Live Model**           | Changing an input or the data updates the tables, charts and insights |
| 📚 **Covers Every Task**    | All ten instructions are completed in the sheets they belong to       |
| 🧭 **Easy Navigation**      | Jump links, colour coding and labelled sections                       |
| 🧮 **Transparent Formulas** | Grey cells show formulas, yellow cells show inputs                    |
| 🖥️ **One-Page Summary**    | The dashboard gives the full story at a glance                        |

---

## 👤 Author

### Gosai Kishal

**📅 Date:** 09 October, 2026

> *"Every dataset has a story — build the sheet that tells it."*

---

*Last updated: 09 October, 2026*
