# Sales Analysis Dashboard (Excel + Power Query)

An end-to-end Excel project: messy ERP sales export -> cleaned dataset -> formula-driven analysis sheets, charts and an executive summary.

**Period:** Jan 2023 - Dec 2025 | **Transactions:** 4,187 | **Locations:** 8 | **Managers:** 8 | **Manager|Salesman combinations:** 55

![Executive summary](docs/images/01_summary.png)

## Business questions answered

| # | Question | Sheet |
|---|----------|-------|
| 1 | Month-over-month (MOM) summary | `MOM_Summary` |
| 2 | MOM trend chart | `MOM_Trend` |
| 3 | Sales share by Manager / Manager\|Salesman (pie charts) | `Manager_Share` |
| 4 | Overall sales (by year, location, customer type) | `Overall_Sales` |
| 5 | Total cost ratio (Cost / Sales) | `Cost_Ratio` |
| 6 | VAT at 15% of total sales | `VAT_15` |
| 7 | Executive summary of all of the above | `Summary` |

## Key results

| Metric | Value |
|---|---|
| Total sales | 150,104,201 |
| Total cost | 81,903,527 |
| Total cost ratio | 54.56% (gross margin 45.44%) |
| VAT @ 15% | 22,515,630 |
| Sales growth | 2023: 50.2M -> 2024: 39.0M (-22.2%) -> 2025: 60.9M (+56.2%) |
| Top manager | Rajiv, 17.8% of sales (lowest: Oliver, 6.1%) |
| Top Manager\|Salesman | Michael\|Tyler, 12.1% of sales |
| Best / weakest month | Aug-25 (8.72M) / Apr-24 (2.00M) |

Observations: Oliver and Kapil have the highest cost ratios (58.0% and 57.6%), so they are worth a pricing/margin review. The top 10 Manager|Salesman pairs generate about 64% of total sales.

![MOM trend](docs/images/02_mom_trend.png)

## Data cleaning

The raw export (`Data` sheet) is a report-style layout, not a flat table. 4,292 rows contained 105 non-transaction rows:

| Problem | Rows | Fix |
|---|---|---|
| Blank spacer rows | 35 | removed |
| Month label rows ("February 2023") | 35 | removed |
| Repeated header rows | 35 | removed |
| `Manager\|Salesman` and `Customer Details` packed into one column | all | split into Manager, Salesman, Customer Type, Customer |
| No month / year fields | all | added `Month`, `Year` |

Result: 4,187 clean transactions (`Clean_Data`). A reconciliation check confirms the cleaned total equals the original `Data` sheet total. The Power Query version of these steps is in [`power_query/clean_sales_data.pq`](power_query/clean_sales_data.pq).

## Workbook structure

```
Summary          executive KPIs, findings, charts
MOM_Summary      selected-month vs previous-month + 36-month MOM table
MOM_Trend        Sales vs Cost line chart, MOM % chart
Manager_Share    pie charts + share tables
Overall_Sales    totals + year / location / customer-type breakdown
Cost_Ratio       cost ratio overall, by year, by manager
VAT_15           VAT calculation (rate is an input cell)
Clean_Data       cleaned transactions (+ formula columns)
PQ_Steps         cleaning log + Power Query script
Data / Sheet1    original raw export and task list (untouched)
```

Everything on the analysis sheets is a live formula (`SUMIFS`, `INDEX/MATCH`, `RANK`, `EDATE`, ...) on `Clean_Data` - about 26,000 formulas, zero errors. Inputs are blue (VAT rate, selected month).

![Manager share](docs/images/03_manager_share.png)

## Assumptions

- Sales are treated as **excluding VAT**; VAT = Total Sales x 15% (change `VAT_15!C4`).
- The export has no currency field, so amounts are treated as a single reporting currency.
- "MOM (Y2601)" is implemented as a selectable month (`MOM_Summary!C4`, default Dec-2025). The data ends Dec-2025, so Jan-2026 returns no data.

## Repository layout

```
excel/Sales_Analysis_Dashboard.xlsx   final workbook (open this)
power_query/clean_sales_data.pq       Power Query (M) cleaning script
scripts/build_workbook.py             Python script that generates the workbook
docs/                                 data dictionary + screenshots
```

## How to run / rebuild

1. Open `excel/Sales_Analysis_Dashboard.xlsx` in Excel (formulas recalculate automatically).
2. Optional rebuild from a raw export:
   ```bash
   pip install -r requirements.txt
   python scripts/build_workbook.py raw_export.xlsx Sales_Analysis_Dashboard.xlsx
   ```
   Then open in Excel (or recalculate in LibreOffice) so values are cached.

## Skills demonstrated

Excel (SUMIFS, INDEX/MATCH, dynamic inputs, charts), Power Query / M, data cleaning and validation, KPI design, MOM/YoY analysis, dashboard storytelling, Python (openpyxl, pandas).

## Author

Adeel - Data Analyst (Excel, Power BI, SQL) | [LinkedIn](https://www.linkedin.com/in/your-profile) | [Upwork](https://www.upwork.com/freelancers/your-profile)
