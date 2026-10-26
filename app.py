import io
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
        file_date = file.name.split(".")[0].strip()
        for sheet_name in xls.sheet_names:
          df_sheet = pd.read_excel(file, sheet_name=sheet_name)
          # Ensure sheet has at least 36 rows (index 35)
          if len(df_sheet) > 35:
            # Row 36 in Excel corresponds to index 35 in pandas
            target_row = df_sheet.iloc[35]

            # Find header row (usually contains bank names, let's look around row index 2-5 or use sheet columns)
            # We map columns using HEADER_MAPPING against column names or header row
            record = {"Bank/Date": file_date}

            for col_idx, val in enumerate(target_row.values):
              # Check column header name
              col_name = (
                  str(df_sheet.columns[col_idx]).strip().upper()
                  if col_idx < len(df_sheet.columns)
                  else ""
              )
              # Also check if header exists in row 0-5
              for h_idx in range(min(6, len(df_sheet))):
                h_val = str(df_sheet.iloc[h_idx, col_idx]).strip().upper()
                if h_val in HEADER_MAPPING:
                  col_name = h_val
                  break

              if col_name in HEADER_MAPPING:
                b_name = HEADER_MAPPING[col_name]
                try:
                  num = float(str(val).replace("%", "").strip())
                  if num > 1.0:
                    num = num / 100.0
                  record[b_name] = num
                except:
                  record[b_name] = None
              elif "WEIGHTED" in col_name or "AVERAGE" in col_name:
                try:
                  num = float(str(val).replace("%", "").strip())
                  if num > 1.0:
                    num = num / 100.0
                  record["Weighted Average"] = num
                except:
                  pass

            if len(record) > 1:
              success_rows.append(record)
              break  # Found in this file, move to next file
      except Exception as e:
        st.error(f"Error reading {file.name}: {e}")

    if success_rows:
      df_success = pd.DataFrame(success_rows)
      bank_cols = [c for c in df_success.columns if c != "Bank/Date"]

      # Calculate Average row at the bottom
      avg_dict = {"Bank/Date": "Average"}
      for col in bank_cols:
        avg_dict[col] = df_success[col].mean()

      df_success = pd.concat(
          [df_success, pd.DataFrame([avg_dict])], ignore_index=True
      )


      # Color Coding Rules:
      # 98.5% - 100% (>= 0.985): Green (#2ecc71)
      # 90% - 98% (0.90 - 0.9849): Yellow (#f1c40f)
      # 80% - 89% (0.80 - 0.899): Light Yellow (#fef9e7 / dark text)
      # <=79% (<= 0.79): Red (#e74c3c)
      def apply_success_styling(val):
        try:
          num = float(val)
          pct_str = f"{num*100:.1f}%"
          if num >= 0.985:
            return f"<span style='background-color: #2ecc71; color: white; font-weight: bold; padding: 2px 4px; border-radius: 3px;'>{pct_str}</span>"
          elif num >= 0.90:
            return f"<span style='background-color: #f1c40f; color: black; font-weight: bold; padding: 2px 4px; border-radius: 3px;'>{pct_str}</span>"
          elif num >= 0.80:
            return f"<span style='background-color: #fcf3cf; color: black; font-weight: bold; padding: 2px 4px; border-radius: 3px;'>{pct_str}</span>"
          else:
            return f"<span style='background-color: #e74c3c; color: white; font-weight: bold; padding: 2px 4px; border-radius: 3px;'>{pct_str}</span>"
        except:
          return str(val)


      display_df = df_success.copy()
      for col in bank_cols:
        display_df[col] = display_df[col].apply(apply_success_styling)

      st.markdown(
          display_df.to_html(escape=False, index=False), unsafe_allow_html=True
      )

      out_succ = io.BytesIO()
      with pd.ExcelWriter(out_succ, engine="openpyxl") as writer:
        df_success.to_excel(
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
          "Could not locate row 36 in the uploaded files. Please check file"
          " structure."
      )
  else:
    st.info("👆 Please upload your daily source report files at the top.")

# ==========================================
# 3. ATM SUCCESS RATE RANKING
# ==========================================
elif report_type == "3. ATM Success Rate Ranking":
  st.subheader("🏆 ATM Success Rate Ranking")
  st.info("Report 3 is ready for review once Report 2 is locked down.")

# ==========================================
# 4. ATM DECLINE CONTRIBUTION BY PARTICIPANT
# ==========================================
elif report_type == "4. ATM decline contribution by participant":
  st.subheader("👤 ATM Decline Contribution by Participant")
  if combined_decline is not None:
    bank_totals = {}
    for col in combined_decline.columns:
      col_clean = str(col).strip().upper()
      if "gedaa" in col_clean.lower() or "gadaa" in col_clean.lower():
        continue
      if col_clean in HEADER_MAPPING:
        target_name = HEADER_MAPPING[col_clean]
        vals = pd.to_numeric(combined_decline[col], errors="coerce").fillna(0)
        bank_totals[target_name] = bank_totals.get(target_name, 0) + vals.sum()

    part_df = pd.DataFrame(
        list(bank_totals.items()),
        columns=["Participant", "Declined transactions"],
    )
    part_df = part_df.sort_values(
        by="Declined transactions", ascending=False
    ).reset_index(drop=True)

    total_declines = part_df["Declined transactions"].sum()
    part_df["Share of ATM declines"] = (
        part_df["Declined transactions"] / total_declines
        if total_declines > 0
        else 0
    )

    display_part = part_df.copy()
    display_part["Declined transactions"] = display_part[
        "Declined transactions"
    ].apply(lambda x: f"{int(x):,}")
    display_part["Share of ATM declines"] = display_part[
        "Share of ATM declines"
    ].apply(lambda x: f"{x:.2%}")

    st.dataframe(display_part, use_container_width=True)
  else:
    st.info("👆 Please upload your daily decline Excel files at the top.")

# ==========================================
# 5. MAJOR ATM DECLINED TRANSACTIONS
# ==========================================
elif report_type == "5. Major ATM declined transactions":
  st.subheader("🔍 Major ATM Declined Transactions")
  if combined_decline is not None:
    rc_col = combined_decline.columns[0]
    combined_decline[rc_col] = (
        pd.to_numeric(combined_decline[rc_col], errors="coerce")
        .fillna(0)
        .astype(int)
        .astype(str)
    )

    total_col = [
        c for c in combined_decline.columns if "total" in str(c).lower()
    ]
    if total_col:
      combined_decline["Total_Numeric"] = pd.to_numeric(
          combined_decline[total_col[0]], errors="coerce"
      ).fillna(0)
    else:
      bank_cols = [
          c
          for c in combined_decline.columns
          if c != rc_col
          and "gedaa" not in str(c).lower()
          and "gadaa" not in str(c).lower()
      ]
      combined_decline["Total_Numeric"] = (
          combined_decline[bank_cols]
          .apply(pd.to_numeric, errors="coerce")
          .fillna(0)
          .sum(axis=1)
      )

    rc_summary = (
        combined_decline.groupby(rc_col)["Total_Numeric"].sum().reset_index()
    )
    rc_summary.columns = ["Response code", "Declined transactions"]
    rc_summary = rc_summary.sort_values(
        by="Declined transactions", ascending=False
    ).reset_index(drop=True)

    total_declines_rc = rc_summary["Declined transactions"].sum()
    rc_summary["Share of ATM declines"] = (
        rc_summary["Declined transactions"] / total_declines_rc
        if total_declines_rc > 0
        else 0
    )

    display_rc = rc_summary.copy()
    display_rc["Declined transactions"] = display_rc[
        "Declined transactions"
    ].apply(lambda x: f"{int(x):,}")
    display_rc["Share of ATM declines"] = display_rc[
        "Share of ATM declines"
    ].apply(lambda x: f"{x:.2%}")

    st.dataframe(display_rc, use_container_width=True)
  else:
    st.info("👆 Please upload your daily decline Excel files at the top.")
