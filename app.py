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
st.markdown(
    "Upload all daily ATM decline report files for the month below:"
)

uploaded_files = st.file_uploader(
    "Upload daily files",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

st.markdown("---")

# Header Mapping for Clean Bank Names
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


# Combine data if files are uploaded
combined_decline = None
if uploaded_files:
  combined_decline = process_all_decline_files(uploaded_files)

# ----------------------------------------------------
# 2. REPORT TYPE SELECTION (Bottom)
# ----------------------------------------------------
st.markdown(
    '<div class="section-header">Report Type</div>', unsafe_allow_html=True
)

report_type = st.radio(
    "Report Type Selection",
    [
        "Home / Overview",
        "1. ATM decline summery",
        "2. ATM Transaction Success Rate Summary",
        "3. ATM success rate ranking",
        "4. ATM decline contribution by participant",
        "5. Major ATM declined transactions",
    ],
    label_visibility="collapsed",
)

st.markdown("---")

# ==========================================
# REPORT VIEWS
# ==========================================

if report_type == "Home / Overview":
  st.subheader("Welcome to the EthSwitch Reporting Hub")
  st.markdown("""
    * **1. ATM decline summery:** Overview of consolidated daily declines.
    * **2. ATM Transaction Success Rate Summary:** Detailed bank success rate performance.
    * **3. ATM success rate ranking:** Bank rankings evaluated against the 98.5% target.
    * **4. ATM decline contribution by participant:** Participant-wise volume and percentage share of declines.
    * **5. Major ATM declined transactions:** Breakdown of major decline causes by Response Code.
    """)

elif report_type == "1. ATM decline summery":
  st.subheader("📊 1. ATM Decline Summary")
  if combined_decline is not None:
    st.success(
        f"Successfully processed {len(uploaded_files)} daily report files!"
    )
    st.dataframe(combined_decline.head(50), use_container_width=True)
  else:
    st.info("👆 Please upload your daily ATM decline Excel files at the top.")

elif report_type == "2. ATM Transaction Success Rate Summary":
  st.subheader("📈 2. ATM Transaction Success Rate Summary")
  success_file = st.file_uploader(
      "Upload ATM Success Rate Report (.xlsx)",
      type=["xlsx", "xls"],
      key="success_summary_file",
  )
  if success_file is not None:
    df_success = pd.read_excel(success_file, sheet_name=0)
    st.dataframe(df_success, use_container_width=True)
  else:
    st.info("Please upload your ATM Success Rate Excel file.")

elif report_type == "3. ATM success rate ranking":
  st.subheader("🏆 3. ATM Success Rate Ranking & Target Evaluation")
  target_pct = (
      st.slider(
          "Target Success Rate Threshold (%)",
          min_value=90.0,
          max_value=100.0,
          value=98.5,
          step=0.1,
      )
      / 100.0
  )

  success_file_rank = st.file_uploader(
      "Upload ATM Success Rate Report for Ranking (.xlsx)",
      type=["xlsx", "xls"],
      key="ranking_file_3",
  )
  if success_file_rank is not None:
    df_success = pd.read_excel(success_file_rank, sheet_name=0)
    header_idx, avg_idx = None, None
    for idx, row in df_success.iterrows():
      row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
      if "Bank/Date" in row_str or "Bank" in row_str:
        header_idx = idx
      if "Average" in row_str:
        avg_idx = idx

    if header_idx is not None and avg_idx is not None:
      cols = df_success.iloc[header_idx].tolist()
      avg_row = df_success.iloc[avg_idx].tolist()

      success_data = []
      for c, v in zip(cols[1:], avg_row[1:]):
        if pd.notna(c) and str(c).strip() != "Weighted Average":
          bank_name = HEADER_MAPPING.get(
              str(c).strip().upper(), str(c).strip()
          )
          rate = float(v)
          success_data.append({
              "Participant": bank_name,
              "Success rate": rate,
              "Target status": (
                  "Meets target" if rate >= target_pct else "Below target"
              ),
          })

      rank_df = pd.DataFrame(success_data)
      rank_df = rank_df.sort_values(
          by="Success rate", ascending=False
      ).reset_index(drop=True)
      rank_df.index += 1
      rank_df["Rank"] = rank_df.index
      rank_df = rank_df[
          ["Rank", "Participant", "Success rate", "Target status"]
      ]

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
    st.info("Please upload your ATM Success Rate report to generate rankings.")

elif report_type == "4. ATM decline contribution by participant":
  st.subheader("👤 4. ATM Decline Contribution by Participant")
  if combined_decline is not None:
    bank_totals = {}
    for col in combined_decline.columns:
      col_clean = str(col).strip().upper()
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

    out_part = io.BytesIO()
    with pd.ExcelWriter(out_part, engine="openpyxl") as writer:
      part_df.to_excel(
          writer, sheet_name="Participant Contribution", index=False
      )
    st.download_button(
        "📥 Download Participant Contribution Report (.xlsx)",
        out_part.getvalue(),
        "Decline_Contribution_By_Participant.xlsx",
        key="dl_part_4",
    )
  else:
    st.info(
        "👆 Please upload your daily decline Excel files at the top to view"
        " participant contributions."
    )

elif report_type == "5. Major ATM declined transactions":
  st.subheader("🔍 5. Major ATM Declined Transactions (Response Codes)")
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
      bank_cols = combined_decline.columns[1:-1]
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

    out_rc = io.BytesIO()
    with pd.ExcelWriter(out_rc, engine="openpyxl") as writer:
      rc_summary.to_excel(writer, sheet_name="Major Causes", index=False)
    st.download_button(
        "📥 Download Major Causes Report (.xlsx)",
        out_rc.getvalue(),
        "Major_Causes_Of_Declines.xlsx",
        key="dl_rc_5",
    )
  else:
    st.info(
        "👆 Please upload your daily decline Excel files at the top to view"
        " major declined transactions."
    )
