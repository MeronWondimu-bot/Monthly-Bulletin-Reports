import io
import re
import pandas as pd
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="EthSwitch Performance Reporting Dashboard",
    page_icon="💳",
    layout="wide",
)

# Custom CSS for UI styling
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

# Orange Top Banner
st.markdown(
    '<div class="banner">BI-WEEKLY & MONTHLY REPORT GENERATOR</div>',
    unsafe_allow_html=True,
)

# ----------------------------------------------------
# 1. FILE UPLOADER SECTION (Top)
# ----------------------------------------------------
st.markdown(
    '<div class="section-header">📁 Upload Daily Source Reports</div>',
    unsafe_allow_html=True,
)
st.markdown("Upload all daily ATM report files for the month below:")

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
    "TSEDAY": "Tseday",
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

# ----------------------------------------------------
# REPORT TYPE SELECTION
# ----------------------------------------------------
report_type = st.radio(
    "Report Selection",
    [
        "1. ATM Decline Summary",
        "2. ATM Transaction Success Rate Summary",
        "3. ATM Success Rate Ranking",
        "4. ATM decline contribution by participant",
        "5. Major ATM declined transactions",
    ],
    label_visibility="collapsed",
)

st.markdown("---")

# ==========================================
# 1. ATM DECLINE SUMMARY (Untouched & Perfect)
# ==========================================
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

# ==========================================
# 2. ATM TRANSACTION SUCCESS RATE SUMMARY
# ==========================================
elif report_type == "2. ATM Transaction Success Rate Summary":
  st.subheader("📈 ATM Transaction Success Rate Summary")
  if uploaded_files:
    success_rows = []
    for file in uploaded_files:
      try:
        xls = pd.ExcelFile(file)
        raw_name = file.name.split(".")[0]
        cleaned_name = (
            raw_name.replace("ATM Declined Transaction Report for", "")
            .replace("ATM Success Rate", "")
            .strip()
        )
        file_date = re.sub(r",\s*", ", ", cleaned_name)

        for sheet_name in xls.sheet_names:
          df_sheet = pd.read_excel(file, sheet_name=sheet_name)
          target_row_series = None
          for idx, row in df_sheet.iterrows():
            row_str = " ".join([str(v) for v in row.values if pd.notna(v)]).upper()
            if "SUCC" in row_str and "RATE" in row_str:
              target_row_series = row
              break

          if target_row_series is not None:
            header_idx = None
            for h_idx in range(len(df_sheet)):
              h_str = " ".join([str(v) for v in df_sheet.iloc[h_idx].values if pd.notna(v)]).upper()
              if "CBE" in h_str or "BOA" in h_str:
                header_idx = h_idx
                break

            cols = (
                df_sheet.iloc[header_idx].tolist()
                if header_idx is not None
                else df_sheet.columns.tolist()
            )
            vals = target_row_series.values.tolist()

            record = {"Bank/Date": file_date}
            for c, val in zip(cols, vals):
              c_str = re.sub(r'\s+', ' ', str(c).strip().upper())
              if c_str in HEADER_MAPPING:
                b_name = HEADER_MAPPING[c_str]
                try:
                  num = float(str(val).replace("%", "").strip())
                  record[b_name] = num
                except:
                  record[b_name] = None
              elif "TSEDAY" in c_str:
                try:
                  num = float(str(val).replace("%", "").strip())
                  record["Tseday"] = num
                except:
                  record["Tseday"] = None

            if len(record) > 1:
              success_rows.append(record)
              break
      except Exception as e:
        st.error(f"Error reading {file.name}: {e}")

    if success_rows:
      df_success = pd.DataFrame(success_rows)

      # Chronological sorting
      try:
        norm_dates = df_success["Bank/Date"].str.replace(r",\s*", ", ", regex=True)
        df_success["_dt"] = pd.to_datetime(norm_dates, format="mixed", errors="coerce")
        df_success = df_success.sort_values(by="_dt").drop(columns=["_dt"]).reset_index(drop=True)
      except Exception:
        pass

      # Explicitly force Weighted Average column calculation
      bank_cols = [c for c in df_success.columns if c != "Bank/Date"]
      numeric_df = df_success[bank_cols].apply(pd.to_numeric, errors="coerce")
      df_success["Weighted Average"] = numeric_df.mean(axis=1)

      ordered_cols = ["Bank/Date"] + bank_cols + ["Weighted Average"]
      ordered_cols = list(dict.fromkeys(ordered_cols))
      df_success = df_success[[c for c in ordered_cols if c in df_success.columns]]

      # Average row at bottom
      avg_dict = {"Bank/Date": "Average"}
      calc_cols = [c for c in df_success.columns if c != "Bank/Date"]
      for col in calc_cols:
        numeric_vals = pd.to_numeric(df_success[col], errors="coerce")
        avg_dict[col] = numeric_vals.mean()

      df_success = pd.concat([df_success, pd.DataFrame([avg_dict])], ignore_index=True)

      display_df = df_success.astype(object)

      for idx, row in df_success.iterrows():
        is_avg = str(row["Bank/Date"]).strip().lower() == "average"
        for col in calc_cols:
          val = row[col]
          try:
            num = float(val)
            if pd.isna(num):
              display_df.at[idx, col] = ""
              continue

            pct_val = num * 100 if num <= 1.0 else num
            pct_str = f"{pct_val:.1f}%"

            if is_avg:
              ratio = num if num <= 1.0 else num / 100.0
              if ratio >= 0.985:
                display_df.at[idx, col] = f"<span style='background-color: #2ecc71; color: white; font-weight: bold; padding: 2px 4px; border-radius: 3px;'>{pct_str}</span>"
              elif ratio >= 0.90:
                display_df.at[idx, col] = f"<span style='background-color: #f1c40f; color: black; font-weight: bold; padding: 2px 4px; border-radius: 3px;'>{pct_str}</span>"
              elif ratio >= 0.80:
                display_df.at[idx, col] = f"<span style='background-color: #fcf3cf; color: black; font-weight: bold; padding: 2px 4px; border-radius: 3px;'>{pct_str}</span>"
              else:
                display_df.at[idx, col] = f"<span style='background-color: #e74c3c; color: white; font-weight: bold; padding: 2px 4px; border-radius: 3px;'>{pct_str}</span>"
            else:
              display_df.at[idx, col] = pct_str
          except:
            pass

      st.markdown(display_df.to_html(escape=False, index=False), unsafe_allow_html=True)

      out_succ = io.BytesIO()
      with pd.ExcelWriter(out_succ, engine="openpyxl") as writer:
        df_success.to_excel(writer, sheet_name="Success Rate Summary", index=False)
      st.download_button(
          "📥 Download Success Rate Summary (.xlsx)",
          out_succ.getvalue(),
          "ATM_Success_Rate_Summary.xlsx",
          key="dl_success_sum",
      )
    else:
      st.warning("Could not locate 'SUCC. RATE' rows in the uploaded files. Please check file structure.")
  else:
    st.info("👆 Please upload your daily source report files at the top.")

# ==========================================
# 3. ATM SUCCESS RATE RANKING
# ==========================================
elif report_type == "3. ATM Success Rate Ranking":
  st.subheader("🏆 ATM Success Rate Ranking")
  st.info("Report 3 is ready for review once Report 2 is locked down.")
