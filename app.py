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
        .main-header {
            font-size: 26px;
            font-weight: bold;
            color: #d35400;
            text-align: center;
            margin-bottom: 0px;
        }
        .sub-header {
            font-size: 14px;
            color: #7f8c8d;
            text-align: center;
            margin-bottom: 20px;
        }
        .banner {
            background-color: #fdf2e9;
            padding: 12px;
            border-radius: 5px;
            border-left: 5px solid #d35400;
            text-align: center;
            font-weight: bold;
            color: #d35400;
            font-size: 18px;
            margin-bottom: 20px;
        }
    </style>
""", unsafe_allow_html=True)

# Top Header Section
col_logo1, col_title, col_logo2 = st.columns([1, 3, 1])
with col_logo1:
  st.markdown("### 🏢 ETHSWITCH")
with col_title:
  st.markdown(
      '<div class="main-header">EthSwitch Performance Reporting Hub</div>',
      unsafe_allow_html=True,
  )
  st.markdown(
      '<div class="sub-header">Automated Bi-weekly & Monthly Report'
      " Generator</div>",
      unsafe_allow_html=True,
  )
with col_logo2:
  st.markdown("### 🐦 EthioPay")

st.markdown(
    '<div class="banner">BI-WEEKLY & MONTHLY REPORT GENERATOR</div>',
    unsafe_allow_html=True,
)

# ----------------------------------------------------
# REPORT SELECTOR (Default to Decline Summary first)
# ----------------------------------------------------
st.markdown("### 📌 Select Report")
report_type = st.radio(
    "Report Selection",
    ["1. Decline transactions summary", "2. Success Rate Ranking & Summary"],
    horizontal=True,
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


# Helper function to parse multi-sheet decline workbooks
def process_decline_file(file):
  xls = pd.ExcelFile(file)
  daily_frames = []

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
      daily_frames.append(sub_df)

  if daily_frames:
    return pd.concat(daily_frames, ignore_index=True)
  return None


# ==========================================
# 1. DECLINE TRANSACTIONS SUMMARY
# ==========================================
if report_type == "1. Decline transactions summary":
  st.subheader("📊 Step 1: Upload ATM Decline Excel Workbook")

  decline_file = st.file_uploader(
      "Drop your multi-sheet ATM Decline Excel file here (.xlsx / .xls):",
      type=["xlsx", "xls"],
      key="decline_summary_main",
  )

  if decline_file is not None:
    st.success("File uploaded successfully! Processing data...")
    combined_decline = process_decline_file(decline_file)

    if combined_decline is not None:
      tab1, tab2 = st.tabs(
          ["👤 Participant Contribution", "🔍 Major Causes (Response Codes)"]
      )

      with tab1:
        st.markdown("#### Decline Contribution by Participant")
        bank_totals = {}
        for col in combined_decline.columns:
          col_clean = str(col).strip().upper()
          if col_clean in HEADER_MAPPING:
            target_name = HEADER_MAPPING[col_clean]
            vals = pd.to_numeric(
                combined_decline[col], errors="coerce"
            ).fillna(0)
            bank_totals[target_name] = (
                bank_totals.get(target_name, 0) + vals.sum()
            )

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
            key="dl_part",
        )

      with tab2:
        st.markdown("#### Major Causes of ATM Declines (Response Codes)")
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
            combined_decline.groupby(rc_col)["Total_Numeric"]
            .sum()
            .reset_index()
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
            key="dl_rc",
        )
  else:
    st.info("👆 Please upload your ATM Decline Excel file above to generate the report.")

# ==========================================
# 2. SUCCESS RATE RANKING & SUMMARY
# ==========================================
elif report_type == "2. Success Rate Ranking & Summary":
  st.subheader("🏆 Step 1: Upload ATM Success Rate Excel File")

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

  success_file = st.file_uploader(
      "Drop your ATM Success Rate Excel file here (.xlsx / .xls):",
      type=["xlsx", "xls"],
      key="ranking_file_main",
  )

  if success_file is not None:
    st.success("File uploaded successfully! Processing data...")
    df_success = pd.read_excel(success_file, sheet_name=0)

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

      st.success(
          f"Success Rate Ranking & Target Evaluation (Target:"
          f" {target_pct*100:.1f}%)"
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
          key="dl_rank",
      )
  else:
    st.info("👆 Please upload your ATM Success Rate Excel file above to see bank rankings.")
