# Monthly Sales MIS Report (Excel)

An MIS-style sales report built in Microsoft Excel from a messy sample sales dataset (January to June 2026). The project covers the full workflow of an MIS Executive: cleaning raw data, joining it with a master table, summarising it, charting it and writing business insights.

> **Note:** The dataset is a synthetic (sample) dataset created for practice. It does not belong to any real company.

## Files in this repository

| File | Description |
|---|---|
| `Monthly_Sales_MIS_Report.xlsx` | Final Excel workbook with all sheets |
| `sales_data_messy.csv` | Raw sales data (512 rows, intentionally messy) |
| `product_master.csv` | Product Master (10 products: Product ID, Product Name, Category) |
| `Monthly Sales trend.png`, `Region-wise Sales.png`, `Top 5 Product by Sales.png` | Chart images used in this README |

## Dataset

**Sales data columns:** Order ID, Date, Region (North, South, East, West), Product ID, Quantity, Unit Price, Salesperson

**Product Master columns:** Product ID, Product Name, Category

The sales table has no product names or categories, so they are brought in from the Product Master using VLOOKUP.

## Workbook sheets

1. **Summary** - key numbers and 4 business insights
2. **Charts** - monthly trend, region-wise sales and top 5 products
3. **Pivot** - monthly sales by region
4. **SUMIF Analysis** - region-wise and product-wise total sales
5. **Clean data** - cleaned sales data with Total Sales, Product Name and Category
6. **product master** - reference table

## Charts

### Monthly Sales Trend (Jan to Jun 2026)
![Monthly Sales Trend](Monthly%20Sales%20trend.png)

### Region-wise Total Sales
![Region-wise Sales](Region-wise%20Sales.png)

### Top 5 Products by Sales
![Top 5 Products](Top%205%20Product%20by%20Sales.png)

## Steps followed

**1. Data cleaning**
- Removed 12 duplicate orders using Remove Duplicates
- Fixed inconsistent Region and Product ID entries (extra spaces, wrong letter case) using `TRIM`, `PROPER` and `UPPER`
- Converted text dates to proper dates and checked the range with `MIN` and `MAX`
- Converted Unit Price from text to numbers (removed "Rs." and commas)
- Removed rows with blank, negative or unrealistic Quantity values
- Replaced blank Salesperson names with "Unknown"

**2. Calculations**
- Total Sales = Quantity x Unit Price
- Product Name and Category added with `VLOOKUP` and `IFERROR`
- Region-wise and product-wise totals with `SUMIF`
- Grand totals cross-checked between SUMIF and the Pivot Table

**3. Reporting**
- Pivot Table for monthly sales by region
- 3 charts: monthly trend (line), region-wise sales (column), top 5 products (column)
- Top 5 products pulled automatically with `LARGE` and `INDEX-MATCH`
- Summary sheet with key numbers and insights

## Key results

- Total sales: **Rs. 1,39,33,930** from **489** clean orders (average order value about Rs. 28,495)
- **West** was the top region (30.95% of sales) and **East** the lowest (18.98%)
- **April** was the best month (Rs. 26,57,030) and **March** the weakest (Rs. 15,34,800)
- **Laptop** alone contributed 41.3% of sales; the top 5 products contributed 73.5%
- 6 orders (Rs. 1,83,300, 1.3% of sales) have Product IDs missing from the Product Master

## Business insights

1. West led with 31% of total sales while East was the lowest at 19%, so East needs more focus.
2. March sales were about 42% lower than April, so the reason for the March dip should be checked.
3. Heavy dependence on Laptop and a few other products is a risk, so other products should be promoted.
4. Six orders have Product IDs missing from the Product Master, and the master list should be updated to keep product-wise reporting complete.

## Excel skills used

`VLOOKUP`, `IFERROR`, `SUMIF`, `SUM`, `COUNTA`, `COUNTIF`, `LARGE`, `INDEX`, `MATCH`, `MIN`, `MAX`, `TRIM`, `PROPER`, `UPPER`, `ISNUMBER`, Remove Duplicates, Text to Columns, Find and Replace, Filters, Pivot Table, Charts

## Author

Riya Nashine
