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

# Define standard bank columns in exact expected order
STANDARD_BANKS = [
    "CBE",
    "BOA",
    "ABAY",
    "DASHEN",
    "WEGAGEN",
    "ZEMEN",
    "BUNNA",
    "ENAT",
    "DGB",
    "LION",
    "OIB",
    "BRIB",
    "UB",
    "AIB",
    "CBO",
    "NIB",
    "ADIB",
    "ZAMZAM",
    "HIJRA",
    "AHADU",
    "SINQEE",
    "AMHARA",
    "Tseday",
]


# 4. Strict Line-by-Line, RC-to-RC, Bank-to-Bank Aggregation Function
@st.cache_data
def process_strict_daily_matrix(uploaded_files):
  """Reads each daily report, aligns every line (Response Code) and column (Bank)

  precisely, removes duplicates from description tables, and aggregates them.
  """
  daily_matrices = []
  all_rcs_set = set()

  for file in uploaded_files:
    df_raw = pd.read_excel(file, header=None)
    header_row_idx = None
    for idx, row in df_raw.iterrows():
      if (
          row.astype(str).str.contains("RC/BANK NAME", case=False, na=False).any()
      ):
        header_row_idx = idx
        break

    if header_row_idx is not None:
      df_day = pd.read_excel(file, header=header_row_idx)
      df_day.columns = [str(c).strip() for c in df_day.columns]
      rc_col = df_day.columns[0]

      # Drop nulls and keep only numeric response code rows
      df_day = df_day.dropna(subset=[rc_col])
      df_day = df_day[
          df_day[rc_col].astype(str).str.replace(".", "", regex=False).str.isdigit()
      ].copy()

      # Standardize Response Code as string
      df_day[rc_col] = df_day[rc_col].astype(str).str.strip()

      # Drop duplicate RCs if the description table repeats them at the bottom
      df_day = df_day.drop_duplicates(subset=[rc_col], keep="first")

      # Ensure all standard banks exist in the daily dataframe columns
      for bank in STANDARD_BANKS:
        if bank not in df_day.columns:
          df_day[bank] = 0

      df_day.set_index(rc_col, inplace=True)
      df_numeric = (
          df_day[STANDARD_BANKS]
          .apply(pd.to_numeric, errors="coerce")
          .fillna(0)
      )

      all_rcs_set.update(df_numeric.index)
      daily_matrices.append(df_numeric)

  if not daily_matrices:
    return None

  # Sort RCs numerically/consistently
  sorted_rcs = sorted(list(all_rcs_set), key=lambda x: int(x) if x.isdigit() else 0)

  # Reindex all daily matrices to master grid and sum them up
  master_grid = pd.DataFrame(0.0, index=sorted_rcs, columns=STANDARD_BANKS)

  for mat in daily_matrices:
    aligned_mat = mat.reindex(
        index=sorted_rcs, columns=STANDARD_BANKS, fill_value=0
    )
    master_grid = master_grid.add(aligned_mat, fill_value=0)

  # Reset index to make RC/BANK NAME a regular column
  final_df = master_grid.reset_index()
  final_df.rename(columns={final_df.columns[0]: "RC/BANK NAME"}, inplace=True)

  # Convert bank columns to integer
  final_df[STANDARD_BANKS] = final_df[STANDARD_BANKS].astype(int)

  # Calculate Total column for each response code row
  final_df["Total"] = final_df[STANDARD_BANKS].sum(axis=1)

  # Calculate TOTAL summary row at the bottom
  total_row = final_df[STANDARD_BANKS].sum(numeric_only=True).to_frame().T
  total_row["RC/BANK NAME"] = "TOTAL"
  total_row["Total"] = total_row[STANDARD_BANKS].sum(axis=1)

  # Ensure TOTAL row is integer
  total_row[STANDARD_BANKS] = total_row[STANDARD_BANKS].astype(int)
  total_row["Total"] = total_row["Total"].astype(int)

  # Append TOTAL row
  final_df = pd.concat([final_df, total_row], ignore_index=True)

  return final_df


# 5. Styling Function: Highlight values > 100 in Red
def color_red_if_above_100(val):
  try:
    num = float(val)
    if num > 100:
      return "color: red; font-weight: bold;"
  except ValueError:
    pass
  return ""


# 6. Multi-File Uploader Section
st.markdown("### 📁 **Upload Daily Source Reports**")
uploaded_files = st.file_uploader(
    "Upload all daily ATM decline report files for the month",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

st.markdown("---")

# 7. Report Selection (Stacked Ordered List)
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

# 8. Core Execution Logic
if not uploaded_files:
  st.info(
      "👈 Please upload your **daily report files** above to aggregate and"
      " generate reports."
  )
else:
  try:
    master_df = process_strict_daily_matrix(uploaded_files)

    # --- HOME / OVERVIEW ---
    if selected_report == "Home / Overview":
      st.subheader("Welcome to the Central Dashboard")
      st.success(
          f"Successfully processed and aligned **{len(uploaded_files)} daily"
          " report files** line-by-line."
      )
      st.write(
          "Select **1. ATM decline summery** from the menu above to view the"
          " full aggregated monthly matrix."
      )

    # --- 1. ATM DECLINE SUMMARY ---
    elif selected_report == "1. ATM decline summery":
      st.subheader("1. ATM Decline Summary (Monthly Aggregated Matrix)")
      st.write(
          "Complete line-by-line RC-to-RC and bank-to-bank aggregation across"
          f" {len(uploaded_files)} daily reports. Transactions exceeding"
          " **100** are highlighted in red."
      )

      # Apply conditional styling and format numbers as integers without decimals
      styled_df = (
          master_df.style.map(
              color_red_
