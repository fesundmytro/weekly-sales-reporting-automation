# Weekly Sales Reporting Automation

A Python workflow that processes the newest weekly sales CSV, validates and cleans the data, calculates business metrics, and generates a formatted Excel report.

## Business Scenario

A new sales CSV is received every week. Preparing the management report manually takes approximately three hours: checking data quality, removing duplicates, calculating KPIs, preparing summary tables, and formatting the Excel file.

The script reduces the workflow to:

1. Place the new CSV in the `input/` folder.
2. Run one command or wait for the scheduled task.
3. Open the generated Excel report in `reports/`.

## Automation Impact

Demonstration business scenario:

- estimated manual preparation: approximately 3 hours per week;
- automated report generation: measured locally after setup;
- potential annual time saving depends on the real reporting process.

The manual-time estimate is illustrative and does not represent
a production deployment.

## Workflow

```text
Weekly CSV
    ↓
Data validation and cleaning
    ↓
KPI calculation
    ↓
Comparison with the previous weekly file
    ↓
Formatted Excel report
    ↓
Log file
```

## Data Source and Privacy

The demonstration workflow uses the public Global Superstore dataset:

[Global Superstore dataset on Kaggle](https://www.kaggle.com/datasets/laibaanwer/superstore-sales-dataset)

The project reproduces a realistic weekly reporting scenario using
public data. No employer, customer, or confidential business data
is included.

## Data Cleaning Rules

The script:

- normalizes column names;
- removes exact duplicate rows;
- converts dates and numeric fields;
- fills missing category, market, region, and product names with `Unknown`;
- rejects rows with missing critical fields;
- rejects negative sales, non-positive quantity, and discounts outside 0–1;
- stores rejected rows in a separate Excel sheet;
- records data-quality statistics.

The script does not treat repeated `order_id` values as duplicates because one order may contain multiple products.

## Required Columns

```text
order_id
order_date
sales
quantity
profit
```

The following columns are used in report breakdowns and are created as `Unknown` when absent:

```text
category
market
region
product_name
```

## Excel Report

The workbook contains:

- `Dashboard` — KPIs, previous-week comparison, and charts;
- `By Category`;
- `By Market`;
- `By Region`;
- `Top Products`;
- `Data Quality`;
- `Clean Data`;
- `Rejected Rows` — added only when invalid rows are found.

Main KPIs:

- total sales;
- total profit;
- quantity sold;
- unique orders;
- average order value;
- profit margin.

## Project Structure

```text
weekly-sales-reporting/
├── weekly_report.py
├── prepare_sample_data.py
├── requirements.txt
├── README.md
├── .gitignore
├── run_report.bat
├── run_report.sh
├── input/
├── reports/
└── logs/
```

## Installation

Create and activate a virtual environment:

### Windows

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### Linux or macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Test Data

Place the cleaned Global Superstore file next to the scripts:

```text
sales_clean.csv
```

Create two weekly sample files:

```bash
python prepare_sample_data.py
```

The script writes the last two complete weeks to `input/`, using names such as:

```text
sales_2014-12-21.csv
sales_2014-12-28.csv
```

The main script uses the newest file as the current week and the previous file for week-over-week comparison.

## Run the Report

```bash
python weekly_report.py
```

The generated report appears in:

```text
reports/weekly_sales_report_YYYY-MM-DD.xlsx
```

Execution details and errors are written to:

```text
logs/weekly_report.log
```

## Custom Folders

```bash
python weekly_report.py --input-dir path/to/input --output-dir path/to/reports
```

## Windows Task Scheduler

Create a weekly task and configure:

- Trigger: Weekly;
- Day: Monday;
- Time: 08:00;
- Program: the full path to `python.exe`;
- Arguments: the full path to `weekly_report.py`;
- Start in: the project folder.

Example:

```text
Program:
C:\Users\YourName\weekly-sales-reporting\.venv\Scripts\python.exe

Arguments:
C:\Users\YourName\weekly-sales-reporting\weekly_report.py

Start in:
C:\Users\YourName\weekly-sales-reporting
```

Before enabling the schedule, run the task manually from Task Scheduler and confirm that the report and log file are created.

## Linux cron

Open the cron editor:

```bash
crontab -e
```

Example: run every Monday at 08:00:

```cron
0 8 * * 1 /full/path/.venv/bin/python /full/path/weekly_report.py >> /full/path/logs/cron.log 2>&1
```

Use absolute paths because scheduled tasks may not start in the project directory.

## Repository Files

Recommended files for GitHub:

```text
weekly_report.py
prepare_sample_data.py
README.md
requirements.txt
.gitignore
run_report.bat
run_report.sh
input/.gitkeep
reports/.gitkeep
logs/.gitkeep
```

Do not publish real sales files or generated reports when they contain confidential business data.

## Possible Improvements

- email the report automatically;
- upload the result to Google Drive or SharePoint;
- add budget-versus-actual metrics;
- include alerts when sales fall below a threshold;
- process files from cloud storage;
- store historical KPIs in a database;
- add automated tests and a CI workflow.
