import pandas as pd
import streamlit as st

# 1. Page Configuration
st.set_page_config(
    page_title="EthSwitch Performance Reporting Hub",
    page_icon="🏢",
    layout="wide",
)

# 2. Custom Top Header with Branding Logos
col1, col2, col3 = st.columns([1, 3, 1])

with col1:
  st.markdown("## 🏢 **ETHSWITCH**")

with col2:
  st.markdown(
      "<h1 style='text-align: center; color: #E65100; font-weight: bold;"
      " margin-bottom: 0;'>EthSwitch Performance Reporting Hub</h1>",
      unsafe_allow_html=True,
  )
  st.markdown(
      "<p style='text-align: center; color: #666666; margin-top: 0;'>Automated"
      " Bi-weekly & Monthly Report Generator</p>",
      unsafe_allow_html=True,
  )

with col3:
  st.markdown(
      "<h2 style='text-align: right;'>🦜 <span"
      " style='color:#C2185B;'>EthioPay</span></h2>",
      unsafe_allow_html=True,
  )

st.markdown("<br>", unsafe_allow_html=True)

# 3. Orange Highlight Banner
st.markdown(
    """
    <div style="background-color: #FFF3E0; border-left: 6px solid #E65100; padding: 12px; border-radius: 4px; text-align: center;">
        <h3 style="color: #E65100; margin: 0; font-size: 18px; font-weight: bold; letter-spacing: 1px;">
            BI-WEEKLY & MONTHLY REPORT GENERATOR
        </h3>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("<br>", unsafe_allow_html=True)


# 4. Global Excel File Caching & Loading Function
@st.cache_data
def load_excel_file(uploaded_file):
  xls = pd.ExcelFile(uploaded_file)
  return xls


# 5. File Uploader Section
st.markdown("### 📁 **Upload Source Data**")
uploaded_file = st.file_uploader(
    "Upload your ATM Transaction Excel Report (e.g., ATM Declined Transaction"
    " Report)",
    type=["xlsx", "xls"],
)

st.markdown("---")

# 6. Report Selection (Stacked Ordered List)
st.markdown("### **Report Type**")

report_options = [
    "Home / Overview",
    "1. ATM decline summery",
    "2. ATM Transaction Success Rate Summary",
    "3. ATM success rate ranking",
    "4. ATM decline contribution by participant",
    "5. Major ATM declined transactions",
]

selected_report = st.radio(
    "Select Report Type",
    report_options,
    horizontal=False,  # Vertically stacked list
    label_visibility="collapsed",
)

st.markdown("---")

# 7. Core Execution Logic
if uploaded_file is None:
  st.info(
      "👈 Please upload your transaction Excel file above to generate and"
      " view reports."
  )
else:
  try:
    xls = load_excel_file(uploaded_file)
    sheet_names = xls.sheet_names

    # --- HOME / OVERVIEW ---
    if selected_report == "Home / Overview":
      st.subheader("Welcome to the Central Dashboard")
      st.write(
          "Successfully loaded workbook containing"
          f" **{len(sheet_names)} sheets**."
      )
      st.write("Available sheets in workbook:")
      st.write(", ".join(sheet_names[:10]) + ("..." if len(sheet_names) > 10 else ""))

      # Preview monthly summary sheet if available
      summary_sheet_name = (
          "Sheet16" if "Sheet16" in sheet_names else sheet_names[0]
      )
      df_preview = pd.read_excel(uploaded_file, sheet_name=summary_sheet_name)
      st.markdown(f"### Preview: {summary_sheet_name}")
      st.dataframe(df_preview.head(10), use_container_width=True)

    # --- 1. ATM DECLINE SUMMARY ---
    elif selected_report == "1. ATM decline summery":
      st.subheader("1. ATM Decline Summary")
      st.write(
          "Aggregated transaction decline response summary across participating"
          " banks."
      )

      sheet_choice = st.radio(
          "Select Data Source for Decline Summary:",
          ["Monthly Summary Sheet (Sheet16)", "Aggregate All Daily Sheets"],
          horizontal=True,
      )

      if sheet_choice == "Monthly Summary Sheet (Sheet16)":
        target_sheet = "Sheet16" if "Sheet16" in sheet_names else sheet_names[0]
        df_decline = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=2)
      else:
        daily_sheets = [s for s in sheet_names if s != "Sheet16"]
        all_dfs = []
        for s in daily_sheets:
          temp_df = pd.read_excel(uploaded_file, sheet_name=s, header=2)
          all_dfs.append(temp_df)
        combined = pd.concat(all_dfs, ignore_index=True)
        df_decline = combined.groupby(combined.columns[0], as_index=False).sum(
            numeric_only=True
        )

      # Clean column headers
      df_decline.columns = [str(c).strip() for c in df_decline.columns]

      st.markdown("### 📊 Decline Summary Matrix")
      st.dataframe(df_decline, use_container_width=True)

      if "Total" in df_decline.columns:
        total_declines = df_decline["Total"].sum()
        st.metric(
            label="Total Declined Transactions (Period)",
            value=f"{total_declines:,.0f}",
        )

      # Download Button
      csv_export = df_decline.to_csv(index=False).encode("utf-8")
      st.download_button(
          label="Download Decline Summary Report (CSV)",
          data=csv_export,
          file_name="atm_decline_summary_report.csv",
          mime="text/csv",
      )

    # --- 2. SUCCESS RATE SUMMARY ---
    elif selected_report == "2. ATM Transaction Success Rate Summary":
      st.subheader("2. ATM Transaction Success Rate Summary")
      st.info(
          "Report logic for Success Rate Summary will be implemented here next."
      )

    # --- 3. SUCCESS RATE RANKING ---
    elif selected_report == "3. ATM success rate ranking":
      st.subheader("3. ATM Success Rate Ranking")
      st.info(
          "Report logic for Success Rate Ranking (vs 98.5% benchmark) will be"
          " implemented here."
      )

    # --- 4. DECLINE CONTRIBUTION BY PARTICIPANT ---
    elif selected_report == "4. ATM decline contribution by participant":
      st.subheader("4. ATM Decline Contribution by Participant")
      st.info(
          "Report logic for participant decline contribution shares will be"
          " implemented here."
      )

    # --- 5. MAJOR DECLINED TRANSACTIONS ---
    elif selected_report == "5. Major ATM declined transactions":
      st.subheader("5. Major ATM Declined Transactions")
      st.info(
          "Report logic for major response codes and error descriptions will"
          " be implemented here."
      )

  except Exception as e:
    st.error(
        "An error occurred while processing the Excel file. Please check file"
        f" format. Details: {e}"
    )
