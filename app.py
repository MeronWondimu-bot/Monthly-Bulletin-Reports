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
      "<h1 style='text-align: center; color: #E65100; font-weight: bold; margin-bottom: 0;'>EthSwitch Performance Reporting Hub</h1>",
      unsafe_allow_html=True,
  )
  st.markdown(
      "<p style='text-align: center; color: #666666; margin-top: 0;'>Automated Bi-weekly & Monthly Report Generator</p>",
      unsafe_allow_html=True,
  )

with col3:
  st.markdown(
      "<h2 style='text-align: right;'>🦜 <span style='color:#C2185B;'>EthioPay</span></h2>",
      unsafe_allow_html=True,
  )

st.markdown("<br>", unsafe_allow_html=True)

# 3. Orange Highlight Banner
st.markdown(
    """
    <div style="background-color: #E65100; border-left: 6px solid #B7410E; padding: 12px; border-radius: 4px; text-align: center;">
        <h3 style="color: #000000; margin: 0; font-size: 18px; font-weight: bold; letter-spacing: 1px;">
            BI-WEEKLY & MONTHLY REPORT GENERATOR
        </h3>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("<br>", unsafe_allow_html=True)

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
  if "GADAA" in cleaned:
    return None
  for std_code, aliases in BANK_ALIASES.items():
    if cleaned == std_code or cleaned in aliases:
      return std_code
    for alias in aliases:
      if alias in cleaned:
        return std_code
  return cleaned


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


def extract_datetime_from_file(file):
  filename = getattr(file, "name", "")
  m = re.search(
      r"(\d{1,2})[\s\-_]*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s\-_]*(\d{4})?",
      filename,
      re.IGNORECASE,
  )
  if m:
    day, mon, yr = m.group(1), m.group(2), m.group(3)
    year_str = yr if yr else "2026"
    try:
      return pd.to_datetime(f"{day} {mon} {year_str}")
    except:
      pass
  return None


# 9. Multi-File Uploader Section with Date Checking & Filtering
st.markdown("### 📁 **Upload Daily Source Reports**")
uploaded_files = st.file_uploader(
    "Upload all daily ATM decline report files for the month",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

filtered_files = uploaded_files
if uploaded_files:
  file_dates = []
  file_map = {}
  for idx, f in enumerate(uploaded_files):
    dt = extract_datetime_from_file(f)
    if dt is not None:
      file_dates.append(dt)
      file_map[dt] = f
    else:
      fallback_dt = pd.to_datetime(f"2026-09-{idx+1:02d}")
      file_dates.append(fallback_dt)
      file_map[fallback_dt] = f

  min_dt = min(file_dates)
  max_dt = max(file_dates)

  st.markdown("---")
  st.subheader("📅 **File Date Verification & Filtering**")
  date_filter_option = st.radio(
      "Filter source reports by date?", ["All Uploaded Files", "Custom Date Range"], horizontal=True
  )

  if date_filter_option == "Custom Date Range":
    selected_range = st.date_input(
        "Select Report Date Range",
        value=(min_dt.date(), max_dt.date()),
        min_value=min_dt.date(),
        max_value=max_dt.date(),
    )
    if isinstance(selected_range, tuple) and len(selected_range) == 2:
      start_d, end_d = selected_range
      filtered_files = [
          file_map[dt]
          for dt in file_dates
          if start_d <= dt.date() <= end_d
      ]

st.markdown("---")

# 10. Report Selection
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


# 4. Strict Line-by-Line Aggregation Function (Report 1) - GADAA Excluded
@st.cache_data
def process_strict_daily_matrix(file_tuples_names):
  daily_matrices = []
  all_rcs_set = set()

  for file in file_tuples_names:
    df_raw = pd.read_excel(file, header=None)
    header_row_idx = None
    for idx, row in df_raw.iterrows():
      if row.astype(str).str.contains("RC/BANK NAME", case=False, na=False).any():
        header_row_idx = idx
        break

    if header_row_idx is not None:
      df_day = pd.read_excel(file, header=header_row_idx)
      new_cols = []
      for i, c in enumerate(df_day.columns):
        if i == 0:
          new_cols.append(str(c).strip())
        else:
          std = standardize_bank_name(c)
          if std == "GADAA" or "GADAA" in str(c).upper():
            new_cols.append(None)
          else:
            new_cols.append(std if std in STANDARD_BANKS else None)
      df_day.columns = new_cols

      df_day = df_day.loc[:, df_day.columns.notna()]
      rc_col = df_day.columns[0]
      df_day = df_day.dropna(subset=[rc_col])
      df_day = df_day[~df_day[rc_col].astype(str).str.upper().str.contains("GADAA", na=False)]
      
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
          df_day[STANDARD_BANKS].apply(pd.to_numeric, errors="coerce").fillna(0)
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


# 5. Success Rate Processing Function (Report 2) - Exact Weighted Average from File
@st.cache_data
def process_success_rate_matrix(file_tuples_names):
  daily_success_records = []

  for idx, file in enumerate(file_tuples_names, start=1):
    filename = getattr(file, "name", f"file_{idx}.xlsx")
    date_label = extract_date_label(filename, idx)

    df_raw = pd.read_excel(file, header=None)
    success_row = None
    weighted_avg_from_file = None

    for r_idx, row in df_raw.iterrows():
      row_str = " ".join([str(val) for val in row.values]).upper()
      if "ISSU" in row_str and ("SUCC" in row_str or "RATE" in row_str):
        success_row = row
        break

    bank_success_rates = {"Bank/Date": date_label}

    if success_row is not None:
      header_row_idx = None
      for r_idx, row in df_raw.iterrows():
        if row.astype(str).str.contains("RC/BANK NAME", case=False, na=False).any():
          header_row_idx = r_idx
          break

      if header_row_idx is not None:
        header_vals = df_raw.iloc[header_row_idx].values
        header_cols = [standardize_bank_name(c) for c in header_vals]
        
        for col_idx, col_name in enumerate(header_vals):
          c_str = str(col_name).upper()
          if "WEIGHTED" in c_str or "W.AVG" in c_str or "WEIGHTED AVERAGE" in c_str:
            w_val = success_row.iloc[col_idx]
            try:
              w_str = str(w_val).replace("%", "").strip()
              val_num = float(w_str)
              if val_num <= 1.0 and val_num > 0:
                val_num *= 100
              weighted_avg_from_file = round(val_num, 1)
            except:
              pass

        for bank in STANDARD_BANKS:
          if bank in header_cols:
            col_pos = header_cols.index(bank)
            val = success_row.iloc[col_pos]
            try:
              if pd.isna(val):
                bank_success_rates[bank] = 100.0
              else:
                val_str = str(val).replace("%", "").strip()
                if val_str.lower() == "nan" or val_str == "":
                  bank_success_rates[bank] = 100.0
                else:
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

    if weighted_avg_from_file is not None:
      bank_success_rates["Weighted Average"] = weighted_avg_from_file
    else:
      valid_rates = [v for k, v in bank_success_rates.items() if k in STANDARD_BANKS]
      bank_success_rates["Weighted Average"] = (
          round(sum(valid_rates) / len(valid_rates), 1) if valid_rates else 100.0
      )

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
      val_raw = avg_row[bank].values[0]
      val = round(float(val_raw), 1) if pd.notna(val_raw) else 100.0
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
  if master_df is None:
    return None

  valid_matrix = master_df[
      master_df["RC/BANK NAME"].astype(str).str.upper() != "TOTAL"
  ].copy()
  if valid_matrix.empty:
    return None

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

  major_df["Share of ATM declines"] = major_df["Declined transactions"].apply(
      lambda x: (x / grand_total) * 100 if grand_total > 0 else 0.0
  )

  major_df["Declined transactions formatted"] = (
      major_df["Declined transactions"].astype(int).map("{:,}".format)
  )
  major_df["Share formatted"] = major_df["Share of ATM declines"].map(
      lambda x: f"{x:.2f}%"
  )

  return major_df


# 11. Core Execution Logic
if not filtered_files:
  st.info(
      "👈 Please upload your **daily report files** above to aggregate and"
      " generate reports."
  )
else:
  try:
    if selected_report == "Home / Overview":
      st.subheader("Welcome to the Central Dashboard")
      st.success(
          f"Successfully processed and aligned **{len(filtered_files)} daily report files** line-by-line."
      )
      st.write(
          "Select a report from the menu above to view detailed monthly analytics."
      )

    elif selected_report == "1. ATM decline summery":
      master_df = process_strict_daily_matrix(tuple(filtered_files))
      st.subheader("1. ATM Decline Summary (Monthly Aggregated Matrix)")

      html_table = "<div style='max-height: 600px; overflow-y: auto; overflow-x: auto;'><table style='width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 13px; text-align: center;'>"
      html_table += "<thead><tr style='background-color: #f2f2f2;'>"
      for col in master_df.columns:
        html_table += f"<th style='border: 1px solid #ddd; padding: 8px; position: sticky; top: 0; background-color: #f2f2f2; z-index: 1;'>{col}</th>"
      html_table += "</tr></thead><tbody>"

      for idx, row in master_df.iterrows():
        is_total_row = str(row["RC/BANK NAME"]).upper() == "TOTAL"
        row_bg = "background-color: #f9f9f9; font-weight: bold;" if is_total_row else ""
        html_table += f"<tr style='{row_bg}'>"

        for col_idx, col in enumerate(master_df.columns):
          val = row[col]
          if col_idx == 0:
            html_table += f"<td style='border: 1px solid #ddd; padding: 6px; font-weight: bold; text-align: left;'>{val}</td>"
          else:
            try:
              num_val = float(val)
              formatted_val = f"{int(num_val):,}"
              if num_val > 100:
                html_table += f"<td style='border: 1px solid #ddd; padding: 6px; color: red; font-weight: bold;'>{formatted_val}</td>"
              else:
                html_table += f"<td style='border: 1px solid #ddd; padding: 6px;'>{formatted_val}</td>"
            except ValueError:
              html_table += f"<td style='border: 1px solid #ddd; padding: 6px;'>{val}</td>"

        html_table += "</tr>"

      html_table += "</tbody></table></div>"
      st.markdown(html_table, unsafe_allow_html=True)

      st.markdown("<br>", unsafe_allow_html=True)
      csv_export = master_df.to_csv(index=False).encode("utf-8")
      st.download_button(
          label="Download Complete Decline Summary Report (CSV)",
          data=csv_export,
          file_name="strict_aligned_atm_decline_summary.csv",
          mime="text/csv",
      )

    elif selected_report == "2. ATM Transaction Success Rate Summary":
      success_df = process_success_rate_matrix(tuple(filtered_files))
      st.subheader("2. ATM Transaction Success Rate Summary")
      st.write(
          "Day-to-day issuer success rates (%) by bank extracted directly from daily files. Values strictly below **78%** are highlighted in red."
      )

      html_table = "<div style='max-height: 600px; overflow-y: auto; overflow-x: auto;'><table style='width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 13px; text-align: center;'>"
      html_table += "<thead><tr style='background-color: #f2f2f2;'>"
      for col in success_df.columns:
        html_table += f"<th style='border: 1px solid #ddd; padding: 8px; position: sticky; top: 0; background-color: #f2f2f2; z-index: 1;'>{col}</th>"
      html_table += "</tr></thead><tbody>"

      for idx, row in success_df.iterrows():
        is_avg_row = str(row["Bank/Date"]).upper() == "AVERAGE"
        row_bg = "background-color: #e8f5e9; font-weight: bold;" if is_avg_row else ""
        html_table += f"<tr style='{row_bg}'>"

        for col_idx, col in enumerate(success_df.columns):
          val = row[col]
          if col_idx == 0:
            html_table += f"<td style='border: 1px solid #ddd; padding: 6px; font-weight: bold; text-align: left;'>{val}</td>"
          else:
            try:
              num_val = float(val) if pd.notna(val) else 100.0
              formatted_val = f"{num_val:.1f}%"
              if num_val < 78.0:
                html_table += f"<td style='border: 1px solid #ddd; padding: 6px; color: red; font-weight: bold;'>{formatted_val}</td>"
              else:
                html_table += f"<td style='border: 1px solid #ddd; padding: 6px;'>{formatted_val}</td>"
            except (ValueError, TypeError):
              html_table += "<td style='border: 1px solid #ddd; padding: 6px;'>100.0%</td>"

        html_table += "</tr>"

      html_table += "</tbody></table></div>"
      st.markdown(html_table, unsafe_allow_html=True)

      st.markdown("<br>", unsafe_allow_html=True)
      csv_export = success_df.to_csv(index=False).encode("utf-8")
      st.download_button(
          label="Download Success Rate Summary Report (CSV)",
          data=csv_export,
          file_name="atm_success_rate_summary.csv",
          mime="text/csv",
      )

    elif selected_report == "3. ATM success rate ranking":
      success_df = process_success_rate_matrix(tuple(filtered_files))
      ranking_df = process_success_rate_ranking(success_df)

      st.subheader("3. ATM Success Rate Ranking")
      st.write(
          "Participant ranking based on monthly average success rates evaluated against the **98.5%** target threshold."
      )

      html_table = "<div style='max-height: 600px; overflow-y: auto; overflow-x: auto;'><table style='width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 13px; text-align: center;'>"
      html_table += "<thead><tr style='background-color: #f2f2f2;'>"
      html_table += "<th style='border: 1px solid #ddd; padding: 8px; position: sticky; top: 0; background-color: #f2f2f2; z-index: 1;'>Rank</th>"
      html_table += "<th style='border: 1px solid #ddd; padding: 8px; position: sticky; top: 0; background-color: #f2f2f2; z-index: 1;'>Participant</th>"
      html_table += "<th style='border: 1px solid #ddd; padding: 8px; position: sticky; top: 0; background-color: #f2f2f2; z-index: 1;'>Success rate</th>"
      html_table += "<th style='border: 1px solid #ddd; padding: 8px; position: sticky; top: 0; background-color: #f2f2f2; z-index: 1;'>Target status</th>"
      html_table += "</tr></thead><tbody>"

      for idx, row in ranking_df.iterrows():
        is_below = row["Target status"] == "Below target"
        status_color = "color: red; font-weight: bold;" if is_below else ""
        html_table += "<tr>"
        html_table += f"<td style='border: 1px solid #ddd; padding: 6px; font-weight: bold;'>{row['Rank']}</td>"
        html_table += f"<td style='border: 1px solid #ddd; padding: 6px; text-align: left; font-weight: bold;'>{row['Participant']}</td>"
        html_table += f"<td style='border: 1px solid #ddd; padding: 6px;'>{row['Success rate formatted']}</td>"
        html_table += f"<td style='border: 1px solid #ddd; padding: 6px; {status_color}'>{row['Target status']}</td>"
        html_table += "</tr>"

      html_table += "</tbody></table></div>"
      st.markdown(html_table, unsafe_allow_html=True)

      st.markdown("<br>", unsafe_allow_html=True)
      export_df = ranking_df[
          ["Rank", "Participant", "Success rate formatted", "Target status"]
      ].copy()
      export_df.rename(
          columns={"Success rate formatted": "Success rate"}, inplace=True
      )
      csv_export = export_df.to_csv(index=False).encode("utf-8")
      st.download_button(
          label="Download Success Rate Ranking Report (CSV)",
          data=csv_export,
          file_name="atm_success_rate_ranking.csv",
          mime="text/csv",
      )

    elif selected_report == "4. ATM decline contribution by participant":
      master_df = process_strict_daily_matrix(tuple(filtered_files))
      contrib_df = process_decline_contribution(master_df)

      st.subheader("4. ATM Decline Contribution by Participant")
      st.write(
          "Participant contribution share of total ATM declined transactions computed from the monthly aggregated decline matrix."
      )

      html_table = "<div style='max-height: 600px; overflow-y: auto; overflow-x: auto;'><table style='width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 13px; text-align: center;'>"
      html_table += "<thead><tr style='background-color: #f2f2f2;'>"
      html_table += "<th style='border: 1px solid #ddd; padding: 8px; position: sticky; top: 0; background-color: #f2f2f2; z-index: 1;'>Participant</th>"
      html_table += "<th style='border: 1px solid #ddd; padding: 8px; position: sticky; top: 0; background-color: #f2f2f2; z-index: 1;'>Declined transactions</th>"
      html_table += "<th style='border: 1px solid #ddd; padding: 8px; position: sticky; top: 0; background-color: #f2f2f2; z-index: 1;'>Share of ATM declines</th>"
      html_table += "</tr></thead><tbody>"

      for idx, row in contrib_df.iterrows():
        html_table += "<tr>"
        html_table += f"<td style='border: 1px solid #ddd; padding: 6px; text-align: left; font-weight: bold;'>{row['Participant']}</td>"
        html_table += f"<td style='border: 1px solid #ddd; padding: 6px;'>{row['Declined transactions formatted']}</td>"
        html_table += f"<td style='border: 1px solid #ddd; padding: 6px;'>{row['Share formatted']}</td>"
        html_table += "</tr>"

      html_table += "</tbody></table></div>"
      st.markdown(html_table, unsafe_allow_html=True)

      st.markdown("<br>", unsafe_allow_html=True)
      export_df = contrib_df[
          ["Participant", "Declined transactions", "Share formatted"]
      ].copy()
      export_df.rename(columns={"Share formatted": "Share of ATM declines"}, inplace=True)
      csv_export = export_df.to_csv(index=False).encode("utf-8")
      st.download_button(
          label="Download Decline Contribution Report (CSV)",
          data=csv_export,
          file_name="atm_decline_contribution_by_participant.csv",
          mime="text/csv",
      )

    elif selected_report == "5. Major ATM declined transactions":
      master_df = process_strict_daily_matrix(tuple(filtered_files))
      major_df = process_major_declines(master_df)

      st.subheader("5. Major ATM Declined Transactions")
      st.write(
          "Aggregation of major ATM declined transactions grouped by response code and ranked by decline volume share."
      )

      html_table = "<div style='max-height: 600px; overflow-y: auto; overflow-x: auto;'><table style='width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 13px; text-align: center;'>"
      html_table += "<thead><tr style='background-color: #f2f2f2;'>"
      html_table += "<th style='border: 1px solid #ddd; padding: 8px; position: sticky; top: 0; background-color: #f2f2f2; z-index: 1;'>Response code</th>"
      html_table += "<th style='border: 1px solid #ddd; padding: 8px; position: sticky; top: 0; background-color: #f2f2f2; z-index: 1;'>Declined transactions</th>"
      html_table += "<th style='border: 1px solid #ddd; padding: 8px; position: sticky; top: 0; background-color: #f2f2f2; z-index: 1;'>Share of ATM declines</th>"
      html_table += "</tr></thead><tbody>"

      for idx, row in major_df.iterrows():
        html_table += "<tr>"
        html_table += f"<td style='border: 1px solid #ddd; padding: 6px; font-weight: bold;'>{row['Response code']}</td>"
        html_table += f"<td style='border: 1px solid #ddd; padding: 6px;'>{row['Declined transactions formatted']}</td>"
        html_table += f"<td style='border: 1px solid #ddd; padding: 6px;'>{row['Share formatted']}</td>"
        html_table += "</tr>"

      html_table += "</tbody></table></div>"
      st.markdown(html_table, unsafe_allow_html=True)

      st.markdown("<br>", unsafe_allow_html=True)
      export_df = major_df[
          ["Response code", "Declined transactions formatted", "Share formatted"]
      ].copy()
      export_df.rename(
          columns={
              "Declined transactions formatted": "Declined transactions",
              "Share formatted": "Share of ATM declines",
          },
          inplace=True,
      )
      csv_export = export_df.to_csv(index=False).encode("utf-8")
      st.download_button(
          label="Download Major Declined Transactions Report (CSV)",
          data=csv_export,
          file_name="major_atm_declined_transactions.csv",
          mime="text/csv",
      )

  except Exception as e:
    st.error(f"An error occurred while processing daily files: {e}")
