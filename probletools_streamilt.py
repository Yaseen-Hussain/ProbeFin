import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import xlrd
import io
import re
from datetime import datetime

# ================== Helper Functions ==================

def process_probe_data(uploaded_files):
    """Tool 1: Probe Data Processor"""
    financial_fields = [
        "Net Revenue", "Cost of Materials Consumed", "Gross Profit Margin (%)",
        "EBITDA Margin (%)", "Depreciation and Amortization Expense",
        "Finance Costs", "Profit for the Period", "Total Non-current Liabilities",
        "Total Current Liabilities", "Total Equity", "Intangible Assets",
        "Current Ratio", "Short Term Borrowings", "Long Term Borrowings",
        "Operating Profit ( EBITDA )", "Interest Coverage Ratio",
        "Payables / Sales (Days)", "Debtors / Sales (Days)",
        "Inventory / Sales (Days)", "Cash Conversion Cycle (Days)",
        "Return on Capital Employed (%)", "Return on Equity (%)",
        "Total Net Fixed Assets"
    ]

    output_data = []
    for uploaded_file in uploaded_files:
        try:
            wb_about = xlrd.open_workbook(file_contents=uploaded_file.read())

            # Sheet 1: "About the Company"
            about_sheet = wb_about.sheet_by_name("About the Company")
            company_name = about_sheet.cell_value(0, 1)

            incorporation_date = None
            for row in range(about_sheet.nrows):
                if str(about_sheet.cell_value(row, 0)).strip() == "Date of Incorporation":
                    incorporation_date = str(about_sheet.cell_value(row, 1)).strip()
                    break

            incorporation_year = None
            vintage_years = None
            if incorporation_date:
                try:
                    parsed_date = datetime.strptime(incorporation_date, "%d %b, %Y")
                    incorporation_year = parsed_date.year
                    vintage_years = datetime.now().year - incorporation_year
                except:
                    incorporation_year = incorporation_date

            # Sheet 2: "Standalone Financial Data"
            fin_sheet = wb_about.sheet_by_name("Standalone Financial Data")
            latest_col = None
            date_of_report = None
            for col in reversed(range(fin_sheet.ncols)):
                if fin_sheet.cell_value(0, col):
                    latest_col = col
                    date_of_report = str(fin_sheet.cell_value(0, col)).strip()
                    break

            fin_data = {}
            for row in range(fin_sheet.nrows):
                key = str(fin_sheet.cell_value(row, 0)).strip()
                if key in financial_fields:
                    try:
                        val = fin_sheet.cell_value(row, latest_col)
                        val = float(val) if isinstance(val, (int, float)) else None
                    except:
                        val = None
                    fin_data[key] = val

            net_revenue = fin_data.get("Net Revenue", 0)
            depreciation = fin_data.get("Depreciation and Amortization Expense", 0)
            finance_costs = fin_data.get("Finance Costs", 0)
            pat = fin_data.get("Profit for the Period", 0)
            total_non_current = fin_data.get("Total Non-current Liabilities", 0)
            total_current = fin_data.get("Total Current Liabilities", 0)
            total_equity = fin_data.get("Total Equity", 0)
            intangible_assets = fin_data.get("Intangible Assets", 0)
            current_ratio = fin_data.get("Current Ratio", 0)
            short_term = fin_data.get("Short Term Borrowings", 0)
            long_term = fin_data.get("Long Term Borrowings", 0)
            ebitda = fin_data.get("Operating Profit ( EBITDA )", 0)

            total_outside_liabilities = total_non_current + total_current
            tangible_net_worth = total_equity - intangible_assets
            tol_tnw = total_outside_liabilities / tangible_net_worth if tangible_net_worth else None
            total_debt = short_term + long_term
            debt_ebitda = total_debt / ebitda if ebitda else None
            depreciation_pct = depreciation / net_revenue if net_revenue else None
            finance_pct = finance_costs / net_revenue if net_revenue else None
            pat_margin = pat / net_revenue if net_revenue else None
            fatr = net_revenue / fin_data.get("Total Net Fixed Assets", 0) if fin_data.get("Total Net Fixed Assets", 0) else None

            output_data.append({
                "Company Name": company_name,
                "Net Revenue": net_revenue,
                "Gross Margin(%)": fin_data.get("Gross Profit Margin (%)"),
                "EBITDA (%)": fin_data.get("EBITDA Margin (%)"),
                "Depreciation (% of sales)": depreciation_pct,
                "Finance Cost (% of sales)": finance_pct,
                "PAT %": pat_margin,
                "TOL": total_outside_liabilities,
                "TNW": tangible_net_worth,
                "Current Ratio": current_ratio,
                "TOL/TNW ": tol_tnw,
                "Debt": total_debt,
                "Debt/EBITDA": debt_ebitda,
                "Interest Coverage Ratio": fin_data.get("Interest Coverage Ratio"),
                "ROCE (%)": fin_data.get("Return on Capital Employed (%)"),
                "ROE (%)": fin_data.get("Return on Equity (%)"),
                "Fixed Asset Turnover Ratio": fatr,
                "Net Fixed Assets": fin_data.get("Total Net Fixed Assets"),
                "Vintage (Years)": vintage_years,
                "Date of Incorporation": incorporation_date,
                "Date of Report": date_of_report,
            })
        except Exception as e:
            st.error(f"Error processing {uploaded_file.name}: {e}")

    return pd.DataFrame(output_data)


def convert_to_fy_format(column_name):
    if "31 Mar" in str(column_name):
        year_match = re.search(r'(\d{4})', str(column_name))
        if year_match:
            full_year = year_match.group(1)
            return f"FY{full_year[-2:]}"
    return column_name


def process_three_years(uploaded_file):
    """Tool 2: 3-Year Financials + Combo Chart"""
    df = pd.read_excel(uploaded_file, sheet_name="Standalone Financial Data", header=0)
    df.columns = df.columns.map(convert_to_fy_format)

    rows_needed = [
        "Net Revenue", "Total Equity", "Long Term Borrowings", "Short Term Borrowings",
        "EBITDA Margin (%)", "Profit for the Period", "Total Net Fixed Assets"
    ]
    df_filtered = df[df.iloc[:, 0].isin(rows_needed)].set_index(df.columns[0])

    fy_columns = [col for col in df_filtered.columns if str(col).startswith("FY")]
    last_3_years = sorted(fy_columns)[-3:]

    processed = pd.DataFrame({
        "Net Revenue": df_filtered.loc["Net Revenue", last_3_years],
        "Total Equity": df_filtered.loc["Total Equity", last_3_years],
        "Debt": df_filtered.loc["Long Term Borrowings", last_3_years] + df_filtered.loc["Short Term Borrowings", last_3_years],
        "Net Fixed Asset": df_filtered.loc["Total Net Fixed Assets", last_3_years],
        "EBITDA (%)": df_filtered.loc["EBITDA Margin (%)", last_3_years],
        "PAT (%)": (df_filtered.loc["Profit for the Period", last_3_years] / df_filtered.loc["Net Revenue", last_3_years]) * 100
    }).T

    processed = processed.T.reset_index().rename(columns={'index': 'Year'})

    # --- Chart ---
    fig, ax1 = plt.subplots(figsize=(10, 6))
    bar_width = 0.2
    x = range(len(processed["Year"]))

    bars1 = ax1.bar([i - 1.5*bar_width for i in x], processed["Net Revenue"], width=bar_width, label="Net Revenue", color="teal")
    bars2 = ax1.bar([i - 0.5*bar_width for i in x], processed["Total Equity"], width=bar_width, label="Total Equity", color="navy")
    bars3 = ax1.bar([i + 0.5*bar_width for i in x], processed["Debt"], width=bar_width, label="Debt", color="lightsteelblue")
    bars4 = ax1.bar([i + 1.5*bar_width for i in x], processed["Net Fixed Asset"], width=bar_width, label="Net Fixed Asset", color="indigo")

    ax1.set_xticks(x)
    ax1.set_xticklabels(processed["Year"])
    ax1.set_ylabel("INR (₹ Cr)")
    ax1.set_title("Financial Performance (Last 3 FYs)")

    for bars in [bars1, bars2, bars3, bars4]:
        for bar in bars:
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2, height, f"{height:.0f}",
                     ha="center", va="bottom", fontsize=8)

    ax2 = ax1.twinx()
    ax2.plot(processed["Year"], processed["EBITDA (%)"], marker="o", color="powderblue", linewidth=2, label="EBITDA %")
    ax2.plot(processed["Year"], processed["PAT (%)"], marker="o", color="limegreen", linewidth=2, label="PAT %")
    ax2.set_ylabel("Percentage (%)")

    for i, val in enumerate(processed["EBITDA (%)"]):
        ax2.text(i, val, f"{val:.0f}%", ha="center", va="bottom", fontsize=9)
    for i, val in enumerate(processed["PAT (%)"]):
        ax2.text(i, val, f"{val:.0f}%", ha="center", va="bottom", fontsize=9)

    bars_labels, bars_handles = ax1.get_legend_handles_labels()
    lines_labels, lines_handles = ax2.get_legend_handles_labels()
    ax1.legend(bars_labels + lines_labels, bars_handles + lines_handles, loc="upper left")

    plt.tight_layout()
    return processed, fig


# ================== Streamlit App ==================

st.title("📊 Financial Analysis Tools")

tool_choice = st.selectbox("Choose a Tool:", ["Probe Data Processor", "3-Year Financials + Chart"])

if tool_choice == "Probe Data Processor":
    uploaded_files = st.file_uploader("Upload Multiple Financial Excel (.xls) Files", type=["xls"], accept_multiple_files=True)
    if uploaded_files:
        df = process_probe_data(uploaded_files)
        st.dataframe(df)

        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        st.download_button("📥 Download Processed Excel", data=output,
                           file_name="financial_summary_output.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif tool_choice == "3-Year Financials + Chart":
    uploaded_file = st.file_uploader("Upload Single Financial Excel (.xls or .xlsx) File", type=["xls", "xlsx"])
    if uploaded_file:
        df, fig = process_three_years(uploaded_file)
        st.dataframe(df)
        st.pyplot(fig)

        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        st.download_button("📥 Download Processed 3Y Excel", data=output,
                           file_name="processed_financials.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
