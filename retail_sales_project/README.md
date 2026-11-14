# Retail Sales Analysis – VBA

A workbook that analyzes a month of retail sales with one click. It calculates the key business
metrics, ranks the top products, analyzes sales by region, highlights high- and low-value
transactions, and writes an analytical summary with charts. All data is synthetic.

## Files

| File | Purpose |
| --- | --- |
| `Retail_Sales_Analysis.xlsm` | Main workbook: **Sales Analysis** (report) and **Sales Data** (records + button) |
| `SalesAnalysis.bas` | The VBA module |
| `retail_sales_oct_2026.csv` | The 100-row sales dataset |
| `generate_sales.py` | Regenerates the dataset (`python generate_sales.py`) |
| `build_workbook.ps1` | Rebuilds the workbook from the CSV and `.bas` file |
| `Test-Sales.ps1` | Checks the results in Excel and writes `validation_results.txt` |

## Dataset (100 transactions, October 2026)

`Order ID | Order Date | Product | Category | Region | Quantity | Unit Price | Sales Amount`

- 12 products in 4 categories: Electronics, Accessories, Home & Kitchen, Clothing
- 4 regions: North, South, East, West
- Sales Amount = Quantity × Unit Price (between 25 and 840)

## How to use

1. Open `Retail_Sales_Analysis.xlsm` and click **Enable Content**.
2. On **Sales Data**, click **Run Sales Analysis**, or press **Alt+F8** and run `RunSalesAnalysis`.
3. To analyze a new month, paste the new records into **Sales Data** using the same columns, then
   click the button again. The report is rebuilt each time, and any number of rows is supported.

## What the macro produces

**Sales Analysis** sheet:

| Section | Contents |
| --- | --- |
| Key Metrics | Total sales, number of transactions, average sale, total units, highest and lowest transaction, count of high- and low-value transactions |
| Top 5 Products by Revenue | Rank, product, units sold, revenue and share of total sales |
| Sales by Region | Transactions, revenue, average sale and share for each region, from best to worst |
| Key Insights | Five automatically written summary sentences |
| Charts | A bar chart of the top 5 products and a column chart of sales by region |

**Sales Data** sheet: high-value transactions (**≥ 500**) are highlighted green, and low-value
transactions (**< 100**) are highlighted red. You can change both thresholds at the top of
`SalesAnalysis.bas` (`HIGH_VALUE`, `LOW_VALUE`).

The macro checks every row before it changes anything. If a date, quantity or sales amount is
invalid, it reports the row (for example, *Row 50: Sales Amount is not a number.*) and leaves the
report unchanged.

## Results for the sample data

- **Total sales:** 29,790.00 from 100 transactions, with an average sale of 297.90
- **Top product:** Air Fryer, at 4,900.00 (16.4% of sales); the top 5 products generated 58.1% of sales
- **Strongest region:** West (29.2%); the weakest is North (19.4%)
- **High-value transactions:** 17, contributing 37.0% of sales; 16 transactions were low-value

## Validation

`Test-Sales.ps1` ran in desktop Excel, and all 15 checks pass (see `validation_results.txt`). The checks
cover:

- every metric, the top 5 products and the region table, compared with figures calculated directly from the CSV
- the row highlighting, the insights and the charts
- reruns, a newly added record, and rejection of an invalid amount

The test adds a temporary error-trapping module to a throwaway copy of the workbook. To do that, it briefly
enables Excel's *Trust access to the VBA project object model* setting, and restores the original
setting afterwards.
