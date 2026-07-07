# CSAT Tabulation & Reporting Automation (VBA)

## Problem
Customer support teams collect large volumes of raw survey/ticket data (85,000+ records in the source dataset), but raw data alone isn't useful without regular tabulation, quality checks, and reporting. A manual process that's slow and error-prone if repeated every reporting cycle.

## Approach
Built a fully macro-driven Excel pipeline that takes raw CSAT (Customer Satisfaction) survey data and automatically produces a cleaned, validated, tabulated, and charted topline report. There is no manual pivot-table building required.

**Pipeline steps (all VBA-driven, run with one click via `RunFullPipeline`):**
1. **Import** — browse for the raw CSV and load it into RawData
2. **Data validation** — flags incomplete records (e.g. missing Order ID) and calculates response-time buckets (Under 1 hr / 1-4 hrs / 4-24 hrs / Over 24 hrs / Invalid-Missing)
3. **Tabulation** — dynamically builds cross-tabs (Avg CSAT + response count) by Category, Agent Shift, Tenure Bucket, and Response Time, each sorted by Avg CSAT and charted (bar + secondary-axis line for volume)
4. **Daily trend** — CSAT trend over time plus daily response volume, on its own sheet
5. **Topline summary** — key metrics (total responses, average CSAT, % satisfied) plus a data-quality section, auto-generated on a Summary sheet

All formulas and charts rebuild automatically if the raw data changes. 
This isn't a one-time analysis; it's a repeatable reporting tool.

## Tools
VBA, Microsoft Excel

## Data
Full [eCommerce Customer Service Satisfaction dataset](https://www.kaggle.com/datasets/ddosad/ecommerce-customer-service-satisfaction) from Kaggle - 85,907 records spanning ~30 days, which is what makes the daily CSAT trend chart meaningful.

## Key results
- Overall average CSAT and % satisfied calculated automatically across all 85,907 responses
- Cross-tabs surface where satisfaction is lowest (e.g. by category, response time) to support process improvement discussions
- CSAT trend over ~30 days shows day-to-day satisfaction alongside response volume
- Data quality check flags a meaningful share of records missing an Order ID, useful for identifying upstream data capture issues

## How to run
1. Open `CSAT_Tabulation_Automation.xlsm` and enable macros
2. Click the **Run** button on the README sheet (or press Alt+F8, select `RunFullPipeline`, click Run)
3. Browse for the raw CSV when prompted

(`VBA_Macros.txt` is included in this repo as a readable reference to the full macro code, in case you'd rather review it without opening Excel.)
