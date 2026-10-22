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
          for dt in
