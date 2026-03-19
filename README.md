# Excel Sales Automation Tool

Excel Sales Automation Tool is a Python project that reads a sales Excel file, cleans the data, calculates key metrics, and generates a formatted Excel report. It now includes both a command-line workflow and a Streamlit web interface.

## Features

- Loads sales data from an Excel file using `pandas`
- Cleans the dataset by removing empty rows and handling missing values
- Calculates `Total Sales` for each row
- Generates key analytics:
  - Total revenue
  - Best selling product
  - Top salesperson
  - Sales by category
- Creates summary tables and charts
- Exports a new Excel report with `Cleaned Data`, `Summary`, and `Report` sheets
- Provides a Streamlit dashboard for upload, preview, KPI cards, charts, and report download

## Expected Input Columns

The source Excel file must contain these columns:

- `Date`
- `Product`
- `Category`
- `Quantity`
- `Price`
- `Salesperson`

## Project Structure

```text
excel_sales_automation_tool/
  __init__.py
  main.py
  streamlit_app.py
  create_sample_input.py
  smoke_test.py
  requirements.txt
  README.md
```

## Installation

```bash
pip install -r requirements.txt
```

## Command-Line Usage

Run the tool by passing the path to your sales Excel file:

```bash
python main.py sample_sales_data.xlsx
```

Optional output path:

```bash
python main.py sample_sales_data.xlsx --output custom_report.xlsx
```

## Streamlit Web App

Start the Streamlit interface:

```bash
streamlit run streamlit_app.py
```

Inside the app you can:

- Upload an Excel workbook
- Review the raw and cleaned data
- See summary metrics instantly
- Explore product, category, and salesperson performance
- Download the generated Excel report

## Sample Input File

Generate a ready-to-use sample Excel file:

```bash
python create_sample_input.py
```

This creates `sample_sales_data.xlsx` in the project folder.

## Smoke Test

Run a small end-to-end smoke test after generating the sample input file:

```bash
python smoke_test.py
```

The smoke test verifies that the output workbook is created and contains these sheets:

- `Cleaned Data`
- `Summary`
- `Report`

## Output

The generated Excel file includes:

- `Cleaned Data` sheet
- `Summary` sheet
- `Report` sheet with summary highlights and embedded charts

## Notes

- Missing numeric values in `Quantity` and `Price` are filled with `0`
- Missing text values in `Product`, `Category`, and `Salesperson` are filled with `Unknown`
- Invalid dates are converted safely and preserved where possible

## Example Workflow

```bash
python create_sample_input.py
python main.py sample_sales_data.xlsx
streamlit run streamlit_app.py
```