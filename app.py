import os
import re
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

# Comprehensive bank alias mapping for full names & short codes
BANK_ALIASES = {
    "CBE": ["CBE", "COMMERCIAL BANK OF ETHIOPIA", "COMMERCIAL", "COMM"],
    "BOA": ["BOA", "BANK OF ABYSSINIA", "ABYSSINIA"],
    "ABAY": ["ABAY", "ABAY BANK"],
    "DASHEN": ["DASHEN", "DASHEN BANK"],
    "WEGAGEN": ["WEGAGEN", "WEGAGEN BANK", "WGB"],
    "ZEMEN": ["ZEMEN", "ZEMEN BANK"],
    "BUNNA": ["BUNNA", "BUNNA BANK"],
    "ENAT": ["ENAT", "ENAT BANK"],
    "DGB": ["DGB", "DEVELOPMENT BANK OF ETHIOPIA", "DBE"],
    "LION": ["LION", "LION INTERNATIONAL BANK", "LIB"],
    "OIB": ["OIB", "OROMIA INTERNATIONAL BANK", "OROMIA BANK"],
    "BRIB": ["BRIB", "BERHAN", "BERHAN BANK"],
    "UB": ["UB", "UNITED BANK", "COOP UB", "UNITED"],
    "AIB": ["AIB", "AWASH", "AWASH BANK", "AWASH INTERNATIONAL BANK"],
    "CBO": ["CBO", "COOPERATIVE BANK OF OROMIA"],
    "NIB": ["NIB", "NIB INTERNATIONAL BANK", "NIB BANK"],
    "ADIB": ["ADIB", "ADDIS INTERNATIONAL BANK", "ADDIS BANK"],
    "ZAMZAM": ["ZAMZAM", "ZAMZAM BANK"],
    "HIJRA": ["HIJRA", "HIJRA BANK"],
    "AHADU": ["AHADU", "AHADU BANK"],
    "SINQEE": ["SINQEE", "SINQEE BANK"],
    "AMHARA": ["AMHARA", "AMHARA BANK"],
    "Tseday": ["TSEDAY", "TSEDAY BANK", "TSEDAY BANK S.C."],
}


def standardize_bank_name(col_name):
  cleaned = str(col_name).strip().upper()
  for std_code, aliases in BANK_ALIASES.items():
    if cleaned == std_code or cleaned in aliases:
      return std_code
    for alias in aliases:
      if alias in cleaned:
        return std_code
  return cleaned


# Helper function to extract date label like '1-Sep' from filename
def extract_date_label(filename, index):
  m = re.search(
      r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s\-_]*(\d{1,2})",
      filename,
      re.IGNORECASE,
  )
  if m:
    return f"{int(m.group(2))}-{m.group(1)[:3].capitalize()}"
  m2 = re.search(
      r"(\d{1,2})[\s\-_]*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*",
      filename,
      re.IGNORECASE,
  )
  if m2:
    return f"{int(m2.group(1))}-{m2.group(2)[:3].capitalize()}"
  return f"{index}-Sep"


# 4. Strict Line-by-Line Aggregation Function (Report 1)
@st.cache_data
def process_strict_daily_matrix(uploaded_files):
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
      df_day.columns = [
          standardize_bank_name(c) if i > 0 else str(c).strip()
          for i, c in enumerate(df_day.columns)
      ]
      rc_col = df_day.columns[0]

      df_day = df_day.dropna(subset=[rc_col])
      df_day = df_day[
          df_day[rc_col].astype(str).str.replace(".", "", regex=False).str.isdigit()
      ].copy()
      df_day[rc_col] = df_day[rc_col].astype(str).str.strip()
      df_day = df_day.drop_duplicates(subset=[rc_col], keep="first")

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

  sorted_rcs = sorted(list(all_rcs_set), key=lambda x: int(x) if x.isdigit() else 0)
  master_grid = pd.DataFrame(0.0, index=sorted_rcs, columns=STANDARD_BANKS)

  for mat in daily_matrices:
    aligned_mat = mat.reindex(
        index=sorted_rcs, columns=STANDARD_BANKS, fill_value=0
    )
    master_grid = master_grid.add(aligned_mat, fill_value=0)

  final_df = master_grid.reset_index()
  final_df.rename(columns={final_df.columns[0]: "RC/BANK NAME"}, inplace=True)
  final_df[STANDARD_BANKS] = final_df[STANDARD_BANKS].astype(int)
  final_df["Total"] = final_df[STANDARD_BANKS].sum(axis=1)

  total_row = final_df[STANDARD_BANKS].sum(numeric_only=True).to_frame().T
  total_row["RC/BANK NAME"] = "TOTAL"
  total_row["Total"] = total_row[STANDARD_BANKS].sum(axis=1)
  total_row[STANDARD_BANKS] = total_row[STANDARD_BANKS].astype(int)
  total_row["Total"] = total_row["Total"].astype(int)

  final_df = pd.concat([final_df, total_row], ignore_index=True)
  return final_df


# 5. Success Rate Processing Function (Report 2)
@st.cache_data
def process_success_rate_matrix(uploaded_files):
  daily_success_records = []

  for idx, file in enumerate(uploaded_files, start=1):
    filename = getattr(file, "name", f"file_{idx}.xlsx")
    date_label = extract_date_label(filename, idx)

    df_raw = pd.read_excel(file, header=None)
    success_row = None

    for r_idx, row in df_raw.iterrows():
      row_str = " ".join([str(val) for val in row.values]).upper()
      if "ISSU" in row_str and ("SUCC" in row_str or "RATE" in row_str):
        success_row = row
        break

    bank_success_rates = {"Bank/Date": date_label}

    if success_row is not None:
      header_row_idx = None
      for r_idx, row in df_raw.iterrows():
        if (
            row.astype(str)
            .str.contains("RC/BANK NAME", case=False, na=False)
            .any()
        ):
          header_row_idx = r_idx
          break

      if header_row_idx is not None:
        header_cols = [
            standardize_bank_name(c)
            for c in df_raw.iloc[header_row_idx].values
        ]
        for bank in STANDARD_BANKS:
          if bank in header_cols:
            col_pos = header_cols.index(bank)
            val = success_row.iloc[col_pos]
            try:
              val_str = str(val).replace("%", "").strip()
              num = float(val_str)
              if num <= 1.0 and num > 0:
                num = num * 100
              bank_success_rates[bank] = round(num, 1)
            except ValueError:
              bank_success_rates[bank] = 100.0
          else:
            bank_success_rates[bank] = 100.0
    else:
      for bank in STANDARD_BANKS:
        bank_success_rates[bank] = 100.0

    valid_rates = [
        v for k, v in bank_success_rates.items() if k in STANDARD_BANKS
    ]
    weighted_avg = (
        round(sum(valid_rates) / len(valid_rates), 1) if valid_rates else 100.0
    )
    bank_success_rates["Weighted Average"] = weighted_avg

    daily_success_records.append(bank_success_rates)

  if not daily_success_records:
    return None

  success_df = pd.DataFrame(daily_success_records)
  cols = ["Bank/Date"] + STANDARD_BANKS + ["Weighted Average"]
  success_df = success_df[cols]

  avg_row = {"Bank/Date": "Average"}
  for bank in STANDARD_BANKS + ["Weighted Average"]:
    col_vals = pd.to_numeric(success_df[bank], errors="coerce")
    avg_row[bank] = round(col_vals.mean(), 1)

  success_df = pd.concat(
      [success_df, pd.DataFrame([avg_row])], ignore_index=True
  )
  return success_df


# 6. Success Rate Ranking Processing Function (Report 3)
@st.cache_data
def process_success_rate_ranking(success_df):
  if success_df is None:
    return None

  avg_row = success_df[success_df["Bank/Date"].astype(str).str.upper() == "AVERAGE"]
  if avg_row.empty:
    return None

  ranking_data = []
  for bank in STANDARD_BANKS:
    if bank in avg_row.columns:
      val = round(float(avg_row[bank].values[0]), 1)
      target_status = "Meets target" if val >= 98.5 else "Below target"
      ranking_data.append(
          {"Participant": bank, "Success rate": val, "Target status": target_status}
      )

  ranking_df = pd.DataFrame(ranking_data)
  ranking_df = ranking_df.sort_values(
      by="Success rate", ascending=False
  ).reset_index(drop=True)
  ranking_df.insert(0, "Rank", range(1, len(ranking_df) + 1))
  ranking_df["Success rate formatted"] = ranking_df["Success rate"].map(
      "{:.1f}%".format
  )

  return ranking_df


# 7. Decline Contribution Processing Function (Report 4)
@st.cache_data
def process_decline_contribution(master_df):
  if master_df is None:
    return None

  total_row_df = master_df[
      master_df["RC/BANK NAME"].astype(str).str.upper() == "TOTAL"
  ]
  if total_row_df.empty:
    return None

  contribution_data = []
  total_declines_all = float(total_row_df["Total"].values[0])

  for bank in STANDARD_BANKS:
    if bank in total_row_df.columns:
      declines = int(total_row_df[bank].values[0])
      share = (
          (declines / total_declines_all) * 100 if total_declines_all > 0 else 0.0
      )
      contribution_data.append(
          {
              "Participant": bank,
              "Declined transactions": declines,
              "Share of ATM declines": round(share, 2),
          }
      )

  contrib_df = pd.DataFrame(contribution_data)
  contrib_df = contrib_df.sort_values(
      by="Declined transactions", ascending=False
  ).reset_index(drop=True)

  contrib_df["Declined transactions formatted"] = contrib_df[
      "Declined transactions"
  ].map("{:,}".format)
  contrib_df["Share formatted"] = contrib_df["Share of ATM declines"].map(
      lambda x: f"{x:.2f}%"
  )

  return contrib_df


# 8. Major Declined Transactions Processing Function (Report 5)
@st.cache_data
def process_major_declines(master_df):
  """Aggregates declined transactions by Response Code from the master grid."""
  if master_df is None:
    return None

  # Exclude the 'TOTAL' row at the bottom
  valid_matrix = master_df[
      master_df["RC/BANK NAME"].astype(str).str.upper() != "TOTAL"
  ].copy()
  if valid_matrix.empty:
    return None

  # Group by Response Code (RC/BANK NAME column) and sum across all banks
  rc_totals = []
  grand_total = 0.0

  for idx, row in valid_matrix.iterrows():
    rc_code = str(row["RC/BANK NAME"]).strip()
    total_declines = float(row["Total"])
    rc_totals.append({"Response code": rc_code, "Declined transactions": total_declines})
    grand_total += total_declines

  major_df = pd.DataFrame(rc_totals)
  major_df = major_df.sort_values(
      by="Declined transactions", ascending=False
  ).reset_index(drop=True)

  # Calculate share
  major_df["Share of ATM declines"] = major_df["Declined transactions"].apply(
      lambda x: (x / grand_total) * 100 if grand_total > 0 else 0.0
  )

  # Formatting
  major_df["Declined transactions formatted"] = major_df[
      "Declined transactions"
  ].astype(int).map("{:,}".format)
  major_df["Share formatted"] = major_df["Share of ATM declines"].map(
      lambda x: f"{x:.2f}%"
  )

  return major_df


# 9. Multi-File Uploader Section
st.markdown("### 📁 **Upload Daily Source Reports**")
uploaded_files = st.file_uploader(
    "Upload all daily ATM decline report files for the month",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

st.markdown("---")

# 10. Report Selection (Stacked Ordered List)
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

# 11. Core Execution Logic
if not uploaded_files:
  st.info(
      "👈 Please upload your **daily report files** above to aggregate and"
      " generate reports."
  )
else:
  try:
    if selected_report == "Home / Overview":
      st.subheader("Welcome to the Central Dashboard")
      st.success(
          f"Successfully processed and aligned **{len(uploaded_files)} daily"
          " report files** line-by-line."
      )
      st.write(
          "Select a report from the menu above to view detailed monthly"
          " analytics."
      )

    elif selected_report == "1. ATM decline summery":
      master_df = process_strict_daily_matrix(uploaded_files)
      st.subheader("1. ATM Decline Summary (Monthly Aggregated Matrix)")

      html_table = (
          "<div style='max-height: 600px; overflow-y: auto; overflow-x:"
          " auto;'><table style='width: 100%; border-collapse: collapse; font-family:"
          " sans-serif; font-size: 13px; text-align: center;'>"
      )
      html_table += "<thead><tr style='background-color: #f2f2f2;'>"
      for col in master_df.columns:
        html_table += f"<th style='border: 1px solid #ddd; padding: 8px; position: sticky; top: 0; background-color: #f2f2f2; z-index: 1;'>{col}</th>"
      html_table += "</tr></thead><tbody>"

      for idx, row in master_df.iterrows():
        is_total_row = str(row["RC/BANK NAME"]).upper() == "TOTAL"
        row_bg = (
            "background-color: #f9f9f9; font-weight: bold;"
            if is_total_row
            else ""
        )
        html_table += f"<tr style='{row_bg}'>"

        for col_idx, col in enumerate(master_df.columns):
          val = row[col]
          if col_idx == 0:
            html_table += f"<td style='border: 1px solid #ddd; padding: 6px; font-weight: bold; text-align: left;'>{val}</td>"
          else:
            try:
              num_val = float(val)
              formatted_val
