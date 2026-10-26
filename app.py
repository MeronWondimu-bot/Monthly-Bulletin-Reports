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
st.markdown("Upload all daily ATM decline report files for the month below:")

uploaded_files = st.file_uploader(
    "Upload daily files",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

st.markdown("---")

# Header Mapping for Clean Bank Names (Gedaa excluded entirely)
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


# Helper function to parse multiple uploaded daily decline files
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
              .str.contains("Total|RC", case=False, na=False)
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
# 2. REPORT TYPE SELECTION
# ----------------------------------------------------
st.markdown(
    '<div class="section-header">Report Type</div>', unsafe_allow_html=True
)

report_type = st.radio(
    "Report Selection",
    [
        "Home / Overview",
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
# REPORT VIEWS & CALCULATIONS
# ==========================================

if report_type == "Home / Overview":
  st.subheader("Welcome to the EthSwitch Reporting Hub")
  st.markdown("""
    * **1. ATM Decline Summary:** Daily responses aggregated per bank (Gedaa excluded, clean headers, highlighted if > 100, and TOTAL row).
    * **2. ATM Transaction Success Rate Summary:** Daily success rates.
    * **3. ATM Success Rate Ranking:** Ranked descending from top to bottom automatically from uploaded files.
    * **4. ATM decline contribution by participant:** Participant shares.
    * **5. Major ATM declined transactions:** Response code summaries.
    """)

elif report_type == "1. ATM Decline Summary":
  st.subheader("📊 ATM Decline Summary")
  if combined_decline is not None:
    rc_col = combined_decline.columns[0]
    
    # Filter out Gedaa / Gadaa entirely from columns and calculations
    valid_cols = []
    for c in combined_decline.columns:
      c_str = str(c).strip().upper()
      if c != rc_col and "GEDAA" not in c_str and "GADAA" not in c_str:
        # Map to clean header name if available
        clean_name = HEADER_MAPPING.get(c_str, c_str)
        if clean_name not in valid_cols:
          valid_cols.append(clean_name)
          # Rename column in dataframe temporarily for grouping
          combined_decline.rename(columns={c: clean_name}, inplace=True)

    for col in valid_cols:
      combined_decline[col] = pd.to_numeric(
          combined_decline[col], errors="coerce"
      ).fillna(0)

    summary_pivot = (
        combined_decline.groupby(rc_col)[valid_cols].sum().reset_index()
    )

    # Clean the Response Code column name so it doesn't display "RC/BANK NAME" or "Response Code" as extra header text
    summary_pivot.rename(columns={rc_col: "RC"}, inplace=True)

    # Append TOTAL row at the bottom (strictly excluding Gedaa)
    total_row = {"RC": "TOTAL"}
    for col in valid_cols:
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


    display_df = summary_pivot.copy()
    for col in valid_cols:
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
    daily_success_frames = []
    for file in uploaded_files:
      try:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
          df_sheet = pd.read_excel(file, sheet_name=sheet_name)
          for idx, row in df_sheet.iterrows():
            row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
            if "Issu. SUCC. RATE (%)" in row_str or "SUCC. RATE" in row_str:
              daily_success_frames.append(df_sheet.iloc[[idx]])
      except Exception:
        pass

    if daily_success_frames:
      success_summary_df = pd.concat(daily_success_frames, ignore_index=True)


      def highlight_low_success(val):
        try:
          num_val = float(str(val).replace("%", ""))
          if num_val <= 79:
            return "color: red; font-weight: bold;"
        except:
          pass
        return ""


      st.dataframe(
          success_summary_df.style.map(highlight_low_success),
          use_container_width=True,
      )
    else:
      st.warning(
          "Could not locate 'Issu. SUCC. RATE (%)' row across uploaded sheets."
      )
  else:
    st.info("👆 Please upload your daily source report files at the top.")

elif report_type == "3. ATM Success Rate Ranking":
  st.subheader("🏆 ATM Success Rate Ranking")
  if uploaded_files:
    success_data = {}
    for file in uploaded_files:
      try:
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
          df_sheet = pd.read_excel(file, sheet_name=sheet_name)
          header_idx, avg_idx = None, None
          for idx, row in df_sheet.iterrows():
            row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
            if "Bank/Date" in row_str or "Bank" in row_str:
              header_idx = idx
            if "Average" in row_str or "Issu. SUCC. RATE (%)" in row_str:
              avg_idx = idx

          if header_idx is not None and avg_idx is not None:
            cols = df_sheet.iloc[header_idx].tolist()
            avg_row = df_sheet.iloc[avg_idx].tolist()
            for c, v in zip(cols[1:], avg_row[1:]):
              if pd.notna(c) and str(c).strip() not in [
                  "Weighted Average",
                  "nan",
              ]:
                bank_name = HEADER_MAPPING.get(
                    str(c).strip().upper(), str(c).strip()
                )
                try:
                  val = float(str(v).replace("%", ""))
                  if val > 1:
                    val = val / 100.0
                  success_data.setdefault(bank_name, []).append(val)
                except:
                  pass
      except Exception:
        pass

    if success_data:
      avg_success = {
          bank: sum(rates) / len(rates) for bank, rates in success_data.items()
      }
      rank_df = pd.DataFrame(
          list(avg_success.items()), columns=["Participant", "Success rate"]
      )
      rank_df = rank_df.sort_values(
          by="Success rate", ascending=False
      ).reset_index(drop=True)
      rank_df.index += 1
      rank_df["Rank"] = rank_df.index
      rank_df = rank_df[["Rank", "Participant", "Success rate"]]

      display_rank = rank_df.copy()
      display_rank["Success rate"] = display_rank["Success rate"].apply(
          lambda x: f"{x*100:.2f}%"
      )

      st.dataframe(display_rank, use_container_width=True)

      out_rank = io.BytesIO()
      with pd.ExcelWriter(out_rank, engine="openpyxl") as writer:
        rank_df.to_excel(
            writer, sheet_name="Success Rate Ranking", index=False
        )
      st.download_button(
          "📥 Download Success Rate Ranking (.xlsx)",
          out_rank.getvalue(),
          "ATM_Success_Rate_Ranking.xlsx",
          key="dl_rank_3",
      )
    else:
      st.warning(
          "Could not compute rankings from the uploaded source files. Ensure"
          " success rate rows are present."
      )
  else:
    st.info(
        "👆 Please upload your daily source report files at the top to generate"
        " rankings automatically."
    )

elif report_type == "4. ATM decline contribution by participant":
  st.subheader("👤 4. ATM Decline Contribution by Participant")
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

    display_part = part_df
