# Sales Data Analysis Dashboard — Excel Project

## Project Overview
This beginner-friendly project uses Microsoft Excel to explore a sample dataset of 300 sales transactions. The goal is to practice data quality checks, summarize sales performance, identify trends, and present results in a dashboard.

> **Dataset note:** The dataset is synthetic practice data. It does not represent real transactions or the performance of an actual company.

## Project Objectives
- Explore sales transactions and identify patterns in revenue and performance.
- Check data quality, including duplicates, blank cells, and date/number formats.
- Summarize revenue by product category, month, and region.
- Calculate key performance indicators (KPIs).
- Present results using PivotTables, PivotCharts, and a dashboard.
- Optionally add slicers for interactive filtering and a monthly revenue forecast.

## Tools and Excel Features
- Microsoft Excel
- Data filters and data-quality checks
- Excel formulas
- PivotTables and PivotCharts
- Charts and dashboard design
- 
## Dataset Fields
The sample sales data includes these fields:
- **Order ID** — transaction identifier
- **Order Date** — date of the transaction
- **City** — city associated with the transaction
- **Region** — sales region
- **Product** — product name
- **Category** — product category
- **Quantity** — units sold
- **Unit Price (₹)** — price per unit
- **Discount** — discount applied
- **Revenue (₹)** — revenue value supplied in the sample
- **Sales Channel** — channel used for the sale

## Project Workflow

### 1. Explore the Data
Open the `Sales Data` worksheet and inspect the available columns, dates, products, categories, regions, and sales channels.

### 2. Check Data Quality
Use filters and validation checks to investigate duplicate IDs, missing values, and incorrect date or number formats. Do not remove records unless the data rules confirm that they are duplicates.

### 3. Optional Calculated Column
A **Gross Sales** column can be added with the formula:

```excel
=G2*H2
```

This calculates quantity multiplied by unit price. In the supplied sample, the Revenue column already reflects the discount, so gross sales and revenue may differ.

### 4. Create PivotTables
- **Category analysis:** Place `Category` in Rows and `Revenue (₹)` in Values. Set the calculation to **Sum**.
- **Monthly trend:** Summarize revenue by month. You can group `Order Date` by Months (and Years if the data spans multiple years), or create a month label with `=TEXT(B2,"mmm yyyy")`.
- **Region analysis:** Place `Region` in Rows and `Revenue (₹)` in Values. Set the calculation to **Sum**.

### 5. Create KPIs
The following example formulas assume the dataset occupies rows 2–301 in `Sales Data`, and each row represents one order:

**Total revenue**
```excel
=SUM('Sales Data'!J2:J301)
```

**Order count**
```excel
=COUNTA('Sales Data'!A2:A301)
```

**Units sold**
```excel
=SUM('Sales Data'!G2:G301)
```

**Average revenue per order**
```excel
=IFERROR(SUM('Sales Data'!J2:J301)/COUNTA('Sales Data'!A2:A301),0)
```

If an order ID appears on multiple rows, count unique orders instead of counting every row. Update the ranges if the dataset size changes.

### 6. Build Charts
- Use a column or bar chart for revenue by category.
- Use a line chart for monthly revenue trends.
- Use a bar or column chart for revenue by region.
- Place the charts and KPI cards on a dashboard worksheet.

### 7. Add Interactivity (Optional)
Add slicers for fields such as `Region`, `Category`, or `Sales Channel`. Connect each slicer to the relevant PivotTables through **Report Connections** or **PivotTable Connections** where supported.

### 8. Monthly Revenue Forecast (Optional)
If your Excel version supports **Data → Forecast Sheet**, select a chronological monthly date series and its revenue values, then create a forecast. Exclude the PivotTable Grand Total from the selected data. Treat forecasts as estimates, especially when there are only a few historical months.

## Business Questions
- Which product categories generate the most revenue?
- Which regions have the highest and lowest revenue?
- How does revenue change month by month?
- Which sales channels contribute to revenue?
- What does a monthly forecast estimate for future revenue?

## How to Use
1. Open the Excel workbook.
2. Start with the `Sales Data` worksheet.
3. Review the data-quality checks.
4. Create or refresh the PivotTables after changes to the source data.
5. Review the charts and KPI cards on the dashboard.
6. If implemented, test the slicers and forecast sheet.

## Project Status
- [ ] Data cleaning and validation
- [ ] Category PivotTable and chart
- [ ] Monthly revenue PivotTable and chart
- [ ] Regional PivotTable and chart
- [ ] KPI cards
- [ ] Interactive slicers
- [ ] Monthly revenue forecast
