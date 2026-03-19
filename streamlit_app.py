from __future__ import annotations

import tempfile
from io import BytesIO
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

from main import REQUIRED_COLUMNS, prepare_report_data, write_excel_report


st.set_page_config(
    page_title="Excel Sales Automation Tool",
    layout="wide",
)


CURRENCY_SYMBOL = "INR"
# SAMPLE_FILE = Path("sample_sales_data.xlsx")
SAMPLE_FILE = Path(__file__).parent / "sample_sales_data.xlsx"

@st.cache_data(show_spinner=False)
def load_uploaded_dataframe(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(BytesIO(file_bytes))


def format_currency(value: float) -> str:
    return f"{CURRENCY_SYMBOL} {value:,.2f}"


def build_report_download(dataframe: pd.DataFrame) -> bytes:
    report_data = prepare_report_data(dataframe)
    with tempfile.TemporaryDirectory() as temp_dir:
        output_path = Path(temp_dir) / "sales_report.xlsx"
        write_excel_report(
            cleaned_df=report_data.cleaned_df,
            summary_table=report_data.summary_table,
            product_sales=report_data.product_sales,
            salesperson_sales=report_data.salesperson_sales,
            category_sales=report_data.category_sales,
            output_file=output_path,
        )
        return output_path.read_bytes()


def render_category_chart(category_sales: pd.DataFrame) -> None:
    fig, ax = plt.subplots(figsize=(5, 5))
    values = category_sales["Total Sales"]
    labels = category_sales["Category"]

    if category_sales.empty or float(values.sum()) == 0:
        values = [1]
        labels = ["No Sales Data"]

    ax.pie(values, labels=labels, autopct="%1.1f%%", startangle=90)
    ax.set_title("Category Distribution")
    st.pyplot(fig, clear_figure=True)


def main() -> None:
    st.title("Excel Sales Automation Tool")
    st.caption(
        "Upload a sales workbook, review insights instantly, and export a formatted Excel report."
    )

    with st.sidebar:
        st.subheader("Input")
        uploaded_file = st.file_uploader("Upload a sales Excel file", type=["xlsx", "xls"])
        st.markdown("Required columns:")
        st.code(", ".join(REQUIRED_COLUMNS))

        if SAMPLE_FILE.exists():
            st.download_button(
                label="Download sample input",
                data=SAMPLE_FILE.read_bytes(),
                file_name=SAMPLE_FILE.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    if uploaded_file is None:
        st.info("Upload an Excel file to preview your sales data and generate the report.")
        return

    file_bytes = uploaded_file.getvalue()

    try:
        raw_df = load_uploaded_dataframe(file_bytes)
        report_data = prepare_report_data(raw_df)
    except Exception as exc:
        st.error(f"Unable to process the uploaded file: {exc}")
        return

    summary_values = dict(zip(report_data.summary_table["Metric"], report_data.summary_table["Value"]))

    metric_col_1, metric_col_2, metric_col_3, metric_col_4 = st.columns(4)
    metric_col_1.metric("Rows Processed", f"{len(report_data.cleaned_df):,}")
    metric_col_2.metric(
        "Total Revenue",
        format_currency(float(summary_values.get("Total Revenue", 0.0))),
    )
    metric_col_3.metric(
        "Best Selling Product",
        str(summary_values.get("Best Selling Product", "N/A")),
    )
    metric_col_4.metric(
        "Top Salesperson",
        str(summary_values.get("Top Salesperson", "N/A")),
    )

    preview_tab, insights_tab, export_tab = st.tabs(["Data Preview", "Insights", "Export Report"])

    with preview_tab:
        left_col, right_col = st.columns(2)
        with left_col:
            st.subheader("Uploaded Data")
            st.dataframe(raw_df, use_container_width=True)
        with right_col:
            st.subheader("Cleaned Data")
            st.dataframe(report_data.cleaned_df, use_container_width=True)

    with insights_tab:
        chart_col, pie_col = st.columns([1.3, 1])
        with chart_col:
            st.subheader("Product Sales")
            product_chart = report_data.product_sales.set_index("Product")
            st.bar_chart(product_chart["Total Sales"], use_container_width=True)
        with pie_col:
            st.subheader("Sales by Category")
            render_category_chart(report_data.category_sales)

        bottom_left, bottom_right = st.columns(2)
        with bottom_left:
            st.subheader("Sales by Product")
            st.dataframe(report_data.product_sales, use_container_width=True)
            st.subheader("Sales by Category")
            st.dataframe(report_data.category_sales, use_container_width=True)
        with bottom_right:
            st.subheader("Sales by Salesperson")
            st.dataframe(report_data.salesperson_sales, use_container_width=True)
            st.subheader("Summary Metrics")
            st.dataframe(report_data.summary_table, use_container_width=True)

    with export_tab:
        st.subheader("Generate Excel Report")
        st.write(
            "Export the same formatted workbook produced by the command-line tool, including summary sheets and embedded charts."
        )

        if st.button("Build Report", type="primary"):
            with st.spinner("Generating Excel report..."):
                try:
                    report_bytes = build_report_download(raw_df)
                except Exception as exc:
                    st.error(f"Failed to generate report: {exc}")
                else:
                    st.success("Report generated successfully.")
                    st.download_button(
                        label="Download sales_report.xlsx",
                        data=report_bytes,
                        file_name="sales_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )


if __name__ == "__main__":
    main()