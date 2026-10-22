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
  st.markdown("<h1 style='text-align: center; color: #E65100; font-weight: bold; margin-bottom: 0;'>EthSwitch Performance Reporting Hub</h1>", unsafe_allow_html=True)
  st.markdown("<p style='text-align: center; color: #666666; margin-top: 0;'>Automated Bi-weekly & Monthly Report Generator</p>", unsafe_allow_html=True)

with col3:
  st.markdown("<h2 style='text-align: right;'>🦜 <span style='color:#C2185B;'>EthioPay</span></h2>", unsafe_allow_html=True)

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
    "CBE", "BOA", "ABAY", "DASHEN", "WEGAGEN", "ZEMEN", "BUNNA", "ENAT",
    "DGB", "LION", "OIB", "BRIB", "UB", "AIB", "CBO", "NIB", "ADIB",
    "ZAMZAM", "HIJRA", "AHADU", "SINQEE", "AMHARA", "Tseday"
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
    "Tseday": ["TSEDAY", "TSEDAY BANK", "TSEDAY BANK S.C."]
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
  m = re.search(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s\-_]*(\d{1,2})", filename, re.IGNORECASE)
  if m:
    return f"{int(m.group(2))}-{m.group(1)[:3].capitalize()}"
  m2 = re.search(r"(\d{1,2})[\s\-_]*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*", filename, re.IGNORECASE)
  if m2:
    return f"{int(m2.group(1))}-{m2.group(2)[:3].capitalize()}"
  return f"{index}-Sep"

def extract_datetime_from_file(file):
  filename = getattr(file, "name", "")
  m = re.search(r"(\d{1,2})[\s\-_]*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[\s\-_]*(\d{4})?", filename, re.IGNORECASE)
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
uploaded_files = st.file_uploader("Upload all daily ATM decline report files for the month", type=["xlsx", "xls"], accept_multiple_files=True)

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
  date_filter_option = st.radio("Filter source reports by date?", ["All Uploaded Files", "Custom Date Range"], horizontal=True)

  if date_filter_option == "Custom Date Range":
    selected_range = st.date_input("Select Report Date Range", value=(min_dt.date(), max_dt.date()), min_value=min_dt.date(), max_value=max_dt.date())
    if isinstance(selected_range, tuple) and len(selected_range) == 2:
      start_d, end_d = selected_range
      filtered_files = [file_map[dt] for dt in file_dates if start_d <= dt.date() <= end_d]

st.markdown("---")

# 10. Report Selection
st.markdown("### **Report Type**")
report_options = [
    "Home / Overview",
    "1. ATM decline summery",
    "2. ATM Transaction Success Rate Summary",
    "3. ATM success rate ranking",
    "4. ATM decline contribution by participant",
    "5. Major ATM declined transactions"
]
selected_report = st.radio("Select Report Type", report_options, horizontal=False, label_visibility="collapsed")
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
      
      df_day = df_day[df_day[rc_col].astype(str).str.replace(".", "", regex=False).str.isdigit()].copy()
      df_day[rc_col] = df_day[rc_col].astype(str).str.strip()
      df_day = df_day.drop_duplicates(subset=[rc_col], keep="first")

      for bank in STANDARD_BANKS:
        if bank not in df_day.columns:
          df_day[bank] = 0

      df_day.set_index(rc_col, inplace=True)
      df_numeric = df_day[STANDARD_BANKS].apply(pd.to_numeric, errors="coerce").fillna(0)

      all_rcs_set.update(df_numeric.index)
      daily_matrices.append(df_numeric)

  if not daily_matrices:
    return None

  sorted_rcs = sorted(list(all_rcs_set), key=lambda x: int(x) if x.isdigit() else 0)
  master_grid = pd.DataFrame(0.0, index=sorted_rcs, columns=STANDARD_BANKS)

  for mat in daily_matrices:
    aligned_mat = mat.reindex(index=sorted_rcs, columns=STANDARD_BANKS, fill_value=0)
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

# 5. Success Rate Processing Function (Report 2) - Direct Extraction from "Issu. SUCC. RATE (%)" row
@st.cache_data
def process_success_rate_matrix(file_tuples_names):
  daily_success_records = []

  for idx, file in enumerate(file_tuples_names, start=1):
    filename = getattr(file, "name", f"file_{idx}.xlsx")
    date_label = extract_date_label(filename, idx)

    df_raw = pd.read_excel(file, header=None)
    success_row = None
    weighted_avg_from_file = None

    # Scan for the exact "Issu. SUCC. RATE (%)" row or similar variants
    for r_idx, row in df_raw.iterrows():
      row_str = " ".join([str(val) for val in row.values]).upper()
      if "ISSU" in row_str and "SUCC" in row_str:
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
        
        # 1. Try to find column containing weighted average / total label in the header row
        for col_idx, col_name in enumerate(header_vals):
          c_str = str(col_name).upper()
          if any(term in c_str for term in ["WEIGHTED", "W.AVG", "TOTAL", "AVERAGE"]):
            w_val = success_row.iloc[col_idx]
            try:
              w_str = str(w_val).replace("%", "").strip()
              val_num = float(w_str)
              if val_num <= 1.0 and val_num > 0:
                val_num *= 100
              weighted_avg_from_file = round(val_num, 1)
              break
            except:
              pass

        # 2. If not found, look at the last non-empty numeric column in the success row (summary column)
        if weighted_avg_from_file is None:
          for col_idx in range(len(success_row) - 1, -1, -1):
            val = success_row.iloc[col_idx]
            try:
              val_str = str(val).replace("%", "").strip()
              num = float(val_str)
              if 0 < num <= 100 or 0 < num <= 1.0:
                if num <= 1.0:
                  num *= 100
                weighted_avg_from_file = round(num, 1)
                break
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
      bank_success_rates["Weighted Average"] = round(sum(valid_rates) / len(valid_rates), 1) if valid_rates else 100.0

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

  success_df = pd
