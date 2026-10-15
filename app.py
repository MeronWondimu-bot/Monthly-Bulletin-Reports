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


# 4. Multi-File Processing Function
@st.cache_data
def process_multiple_daily_files(uploaded_files):
  """Reads multiple daily report files and aggregates them into a single summary."""
  all_daily_dfs = []

  for file in uploaded_files:
    # Read each daily file (assuming format matches your structure where header is at row index 2)
    df_day = pd.read_excel(file, header=2)
    all_daily_dfs.append(df_day)

  if not all_daily_dfs:
    return None

  # Combine all uploaded daily reports
  combined_df = pd.concat(all_daily_dfs, ignore_index=True)

  # Group by the first column (Response Code) and sum up numeric bank values across all days
  rc_col = combined_df.columns[0]
  aggregated_df = combined_df.groupby(rc_col, as_index=False).sum(
      numeric_only=True
  )

  return aggregated_df


# 5. Multi-File Uploader Section
st.markdown("### 📁 **Upload Daily Source Reports**")
uploaded_files = st.file_uploader(
    "Upload multiple daily ATM report files (Select all daily Excel files for"
    " the month)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,  # Allows uploading many files at once!
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
    horizontal=False,
    label_visibility="collapsed",
)

st.markdown("---")

# 7. Core Execution Logic
if not uploaded_files:
  st.info(
      "👈 Please upload your **daily report files** above to aggregate and"
      " generate reports."
  )
else:
  try:
    # Process and aggregate all daily files
    master_decline_df = process_multiple_daily_files(uploaded_files)

    # --- HOME / OVERVIEW ---
    if selected_report == "Home / Overview":
      st.subheader("Welcome to the Central Dashboard")
      st.success(
          f"Successfully uploaded and processed **{len(uploaded_files)} daily"
          " report files**."
      )

      st.markdown("### 📊 Aggregated Monthly Preview")
      st.dataframe(master_decline_df.head(10), use_container_width=True)

    # --- 1. ATM DECLINE SUMMARY ---
    elif selected_report == "1. ATM decline summery":
      st.subheader("1. ATM Decline Summary (Multi-Day Aggregation)")
      st.write(
          "Aggregated transaction decline response summary combining all"
          f" uploaded daily reports ({len(uploaded_files)} files)."
      )

      # Clean column headers
      master_decline_df.columns = [
          str(c).strip() for c in master_decline_df.columns
      ]

      st.markdown("### 📊 Combined Decline Summary Matrix")
      st.dataframe(master_decline_df, use_container_width=True)

      if "Total" in master_decline_df.columns:
        total_declines = master_decline_df["Total"].sum()
        st.metric(
            label="Total Declined Transactions (Selected Period)",
            value=f"{total_declines:,.0f}",
        )

      # Download Button
      csv_export = master_decline_df.to_csv(index=False).encode("utf-8")
      st.download_button(
          label="Download Combined Decline Summary Report (CSV)",
          data=csv_export,
          file_name="combined_atm_decline_summary.csv",
          mime="text/csv",
      )

    # --- 2. SUCCESS RATE SUMMARY ---
    elif selected_report == "2. ATM Transaction Success Rate Summary":
      st.subheader("2. ATM Transaction Success Rate Summary")
      st.info(
          "Report logic for Success Rate Summary across daily uploads will be"
          " implemented here next."
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
        "An error occurred while processing the daily files. Please ensure"
        f" they have matching structures. Details: {e}"
    )
