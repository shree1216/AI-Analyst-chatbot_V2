# Banking Customer Segmentation – VBA

A workbook that cleans banking customer records, finds duplicate and incomplete records, and
distributes customers into category worksheets with one click. All data is synthetic.

## Files

| File | Purpose |
| --- | --- |
| `Customer_Segmentation.xlsm` | Main workbook with the **Process Customers** button |
| `CustomerProcessing.bas` | The VBA module |
| `bank_customers_raw_100.csv` | The 100-row raw dataset |
| `generate_customers.py` | Regenerates the dataset (`python generate_customers.py`) |
| `build_workbook.ps1` | Rebuilds the workbook from the CSV and `.bas` file |
| `Test-Segmentation.ps1` | Runs the macro on a temporary copy and writes `validation_results.txt` |

## Dataset (100 raw records)

`Customer ID | Customer Name | Age | City | Phone | Email | Account Type | Loan Type | Loan Amount`

The raw extract intentionally contains common data-quality problems:

- **8 duplicates** – the same Customer ID entered again later, often in different letter case
- **7 incomplete records** – a missing name, age, city or phone number, an age of `N/A`, or a short phone number
- **Inconsistent formatting:**
  - names such as `chloe young`, `SARAH JACKSON` or names with extra spaces
  - city names in different letter cases
  - phone numbers written as `(555) 216-6158`, `555.412.0649`, `+1 555 961 6569` or `5551234567`
  - email addresses in upper case
  - loan types written as `home`, `Housing Loan`, `AUTO LOAN` or `none`

## How to use

1. Open `Customer_Segmentation.xlsm` and click **Enable Content**.
2. On **Raw Data**, click **Process Customers**, or press **Alt+F8** and run `ProcessCustomers`.
3. To process a new extract, paste it over **Raw Data** with the same column headers, then click the
   button again. The output sheets are rebuilt each time, and Raw Data itself is never changed.

## What the macro does

For each raw record, the macro:

1. **Standardizes** the record:
   - Customer ID → upper case
   - Name, City and Account Type → trimmed and Proper Case
   - Phone → `555-123-4567`
   - Email → lower case
   - Loan Type → `Home Loan`, `Car Loan`, `Personal Loan` or `None`
2. **Checks completeness** – Customer ID, Name, valid Age (18–120), City and a 10-digit phone number
   are required. If any is missing or invalid, the record goes to **Incomplete**, and the Issue column
   lists the fields that need fixing.
3. **Checks for duplicates** – if the Customer ID was already seen, the record goes to **Duplicates**,
   with the row number of the original.
4. **Categorizes** the clean records. A customer can appear in more than one category sheet (for
   example, a senior citizen with a car loan).

| Sheet | Contents |
| --- | --- |
| Summary | Count of customers in each category, with a bar chart |
| Clean Data | Every valid, unique customer (standardized) |
| Senior Citizens | Clean customers aged 60 or over |
| Home Loan | Clean customers with a home loan |
| Car Loan | Clean customers with a car loan |
| Duplicates | Repeated records, each with the row of the original |
| Incomplete | Records with missing or invalid fields, each with the reason |

## Results for the sample data

| Category | Customers |
| --- | --- |
| Total records received | 100 |
| Duplicate records | 8 |
| Incomplete records | 7 |
| Clean customers | 85 |
| Senior citizens (60+) | 21 |
| Home loan holders | 33 |
| Car loan holders | 21 |

## Validation

`Test-Segmentation.ps1` ran in desktop Excel, and all 21 checks pass (see `validation_results.txt`).

- **Counts:** every category count matches a separate calculation made directly from the CSV.
- **Completeness:** every raw record is accounted for, and the summary matches the category sheets.
- **Cleaning:** names, phone numbers, emails and loan types are standardized, and no duplicate IDs remain in Clean Data.
- **Reruns and new records:** rerunning doesn't double any rows, and a newly added raw record is cleaned and placed correctly.
