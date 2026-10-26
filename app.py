import io
import pandas as pd
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="EthSwitch Performance Reporting Dashboard",
    page_icon="💳",
    layout="wide",
)

st.markdown("""
    <style>
        .banner {
            background-color: #d35400;
            padding: 15px;
            border-radius: 5px;
            text-align: center;
            font-weight: bold;
            color: white;
            font-size: 22px;
            margin-bottom: 25px;
        }
        .section-header {
            font-size: 18px;
            font-weight: bold;
            color: #2c3e50;
            margin-top: 15px;
            margin-bottom: 10px;
        }
    </style>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="banner">BI-WEEKLY & MONTHLY REPORT GENERATOR</div>',
    unsafe_allow_html=True,
)

# File Uploader
st.markdown(
    '<div class="section-header">📁 Upload Daily Source Reports</div>',
    unsafe_allow_html=True,
)
uploaded_files = st.file_uploader(
    "Upload daily files",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

st.markdown("---")

HEADER_MAPPING = {
    "CBE": "CBE",
    "BOA": "BOA",
    "AIB": "Awash",
    "SINQEE": "Siinqee",
    "OIB": "Oromia",
    "DASHEN": "Dashen",
    "CBO": "CBO",
    "AMHARA": "Amhara",
    "ABAY": "Abay",
    "BRIB": "Berhan",
    "WEGAGEN": "Wegagen",
    "WGB": "Wegagen",
    "UB": "Hibret",
    "NIB": "NIB",
    "Tseday": "Tseday",
    "ZEMEN": "Zemen",
    "BUNNA": "Bunna",
    "ZAMZAM": "ZamZam",
    "ADIB": "Addis",
    "HIJRA": "Hijra",
    "LION": "Lion",
    "ENAT": "Enat",
    "AHADU": "Ahadu",
    "DGB": "Global",
    "OB": "Oromia",
}


def process_all_decline_files(files):
  all_frames = []
  for file in files:
    try:
      xls = pd.ExcelFile(file)
      for sheet_name in xls.sheet_names:
        if "sheet16" in sheet_name.lower() or "summary" in sheet_name.lower():
          continue
        df_sheet = pd.read_excel(file, sheet_name=sheet_name)
        header_idx = None
        for idx, row in df_sheet.iterrows():
          row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
          if "RC/BANK NAME" in row_str or "RC" in row_str:
            header_idx = idx
            break
        if header_idx is not None:
          df_sheet.columns = [
              str(c).strip() for c in df_sheet.iloc[header_idx].values
          ]
          sub_df = df_sheet.iloc[header_idx + 1 :].copy()
          rc_col = sub_df.columns[0]
          sub_df = sub_df.dropna(subset=[rc_col])
          sub_df = sub_df[
              ~sub_df[rc_col]
              .astype(str)
              .str.contains(
                  "Total|RC|Response Code|Issu", case=False, na=False
              )
          ]
          all_frames.append(sub_df)
    except Exception as e:
      st.error(f"Error processing {file.name}: {e}")
  if all_frames:
    return pd.concat(all_frames, ignore_index=True)
  return None


combined_decline = None
if uploaded_files:
  combined_decline = process_all_decline_files(uploaded_files)

# Report Selection Menu
report_type = st.radio(
    "Report Selection",
    [
        "1. ATM Decline Summary",
        "2. ATM Transaction Success Rate Summary",
    ],
    label_visibility="collapsed",
    horizontal=True,
)

st.markdown("---")

if report_type == "1. ATM Decline Summary":
  st.subheader("📊 ATM Decline Summary")
  if combined_decline is not None:
    rc_col = combined_decline.columns[0]

    valid_cols = []
    for c in combined_decline.columns:
      c_str = str(c).strip().upper()
      if (
          c != rc_col
          and "GEDAA" not in c_str
          and "GADAA" not in c_str
          and "TOTAL" not in c_str
      ):
        clean_name = HEADER_MAPPING.get(c_str, c_str)
        if clean_name not in valid_cols:
          valid_cols.append(clean_name)
          combined_decline.rename(columns={c: clean_name}, inplace=True)

    for col in valid_cols:
      combined_decline[col] = pd.to_numeric(
          combined_decline[col], errors="coerce"
      ).fillna(0)

    summary_pivot = (
        combined_decline.groupby(rc_col)[valid_cols].sum().reset_index()
    )
    summary_pivot.rename(columns={rc_col: "RC"}, inplace=True)

    summary_pivot["Total"] = summary_pivot[valid_cols].sum(axis=1)

    total_row = {"RC": "TOTAL"}
    for col in valid_cols + ["Total"]:
      total_row[col] = summary_pivot[col].sum()

    summary_pivot = pd.concat(
        [summary_pivot, pd.DataFrame([total_row])], ignore_index=True
    )


    def format_and_highlight(val):
      try:
        num = float(val)
        formatted = f"{int(num):,}"
        if num > 100:
          return f"<span style='color: red; font-weight: bold;'>{formatted}</span>"
        return formatted
      except:
        return str(val)


    display_cols = valid_cols + ["Total"]
    display_df = summary_pivot.copy()
    for col in display_cols:
      display_df[col] = display_df[col].apply(format_and_highlight)

    st.markdown(
        display_df.to_html(escape=False, index=False), unsafe_allow_html=True
    )

    out_sum = io.BytesIO()
    with pd.ExcelWriter(out_sum, engine="openpyxl") as writer:
      summary_pivot.to_excel(
          writer, sheet_name="ATM Decline Summary", index=False
      )
    st.download_button(
        "📥 Download ATM Decline Summary (.xlsx)",
        out_sum.getvalue(),
        "ATM_Decline_Summary.xlsx",
        key="dl_decline_sum",
    )
  else:
    st.info("👆 Please upload your daily ATM decline files at the top.")

elif report_type == "2. ATM Transaction Success Rate Summary":
  st.subheader("📈 ATM Transaction Success Rate Summary")
  if uploaded_files:
    success_rows = []
    for file in uploaded_files:
      try:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
          df_sheet = pd.read_excel(file, sheet_name=sheet_name)
          for idx, row in df_sheet.iterrows():
            row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
            if "Issu. SUCC. RATE (%)" in row_str or "SUCC. RATE" in row_str:
              # Extract this row and use sheet name or file name as date indicator
              row_data = df_sheet.iloc[idx].to_dict()
              row_data["Date"] = sheet_name
              success_rows.append(row_data)
      except Exception:
        pass

    if success_rows:
      df_success_summary = pd.DataFrame(success_rows)

      # Highlight values <= 79% (or 0.79) in red
      def highlight_success_rate(val):
        try:
          clean_val = str(val).replace("%", "").strip()
          num = float(clean_val)
          if num <= 0.79 or num <= 79:
            return f"<span style='color: red; font-weight: bold;'>{val}</span>"
        except:
          pass
        return str(val)

      display_success = df_success_summary.copy()
      for col in display_success.columns:
        if col != "Date":
          display_success[col] = display_success[col].apply(
              highlight_success_rate
          )

      st.markdown(
          display_success.to_html(escape=False, index=False),
          unsafe_allow_html=True,
      )

      out_succ = io.BytesIO()
      with pd.ExcelWriter(out_succ, engine="openpyxl") as writer:
        df_success_summary.to_excel(
            writer, sheet_name="Success Rate Summary", index=False
        )
      st.download_button(
          "📥 Download Success Rate Summary (.xlsx)",
          out_succ.getvalue(),
          "ATM_Success_Rate_Summary.xlsx",
          key="dl_success_sum",
      )
    else:
      st.warning(
          "Could not locate 'Issu. SUCC. RATE (%)' rows in the uploaded files."
      )
  else:
    st.info("👆 Please upload your daily source report files at the top.")
