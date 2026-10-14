import io
import pandas as pd
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="EthSwitch Performance Reporting Dashboard",
    page_icon="💳",
    layout="wide",
)

# Sidebar Navigation
st.sidebar.title("📊 Bi-weekly Monthly Reports Menu")
report_selection = st.sidebar.radio(
    "Select Report:",
    [
        "Home / Overview",
        "1. Decline trasnactions summery",
        "2. Success Rate Ranking & Summary",
    ],
)

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


# Helper function to parse decline files across multiple sheets
def process_decline_file(file):
  xls = pd.ExcelFile(file)
  daily_frames = []

  for sheet_name in xls.sheet_names:
    if "sheet16" in sheet_name.lower() or "summary" in sheet_name.lower():
      continue

    df_sheet = pd.read_excel(file, sheet_name=sheet_name)

    # Locate header row containing 'RC/BANK NAME'
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
# HOME / OVERVIEW
# ==========================================
if report_selection == "Home / Overview":
  st.title("💳 EthSwitch Performance & Decline Central Dashboard")
  st.markdown("""
    Welcome to your unified automated reporting dashboard. 
    Use the sidebar menu to navigate between your key operational reports:
    
    * **1. Decline transactions summary:** Consolidates all daily decline sheets to analyze volumes by **Participant Contribution** and **Major Causes (Response Codes)**.
    * **2. Success Rate Ranking & Summary:** Automatically evaluates all participating banks against the **98.5% target threshold**, ranks them, and provides full summaries.
    """)

# ==========================================
# 1. DECLINE TRANSACTIONS SUMMARY
# ==========================================
elif report_selection == "1. Decline trasnactions summery":
  st.title("📊 ATM Decline Transactions Summary")
  st.markdown(
      "Upload your multi-sheet monthly or daily decline workbook below."
  )

  decline_file = st.file_uploader(
      "Upload ATM Decline Workbook (.xlsx / .xls)",
      type=["xlsx", "xls"],
      key="decline_summary",
  )

  if decline_file is not None:
    combined_decline = process_decline_file(decline_file)

    if combined_decline is not None:
      # Create sub-tabs for Participant vs Major Causes
      tab1, tab2 = st.tabs(
          ["👤 Participant Contribution", "🔍 Major Causes (Response Codes)"]
      )

      # --- TAB 1: PARTICIPANT CONTRIBUTION ---
      with tab1:
        st.subheader("Decline Contribution by Participant")
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

        # Download Participant Excel
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

      # --- TAB 2: MAJOR CAUSES ---
      with tab2:
        st.subheader("Major Causes of ATM Declines (Response Code Breakdown)")
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
          bank_cols = combined_decline.columns[1:]
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

        # Download Major Causes Excel
        out_rc = io.BytesIO()
        with pd.ExcelWriter(out_rc, engine="openpyxl") as writer:
          rc_summary.to_excel(writer, sheet_name="Major Causes", index=False)
        st.download_button(
            "📥 Download Major Causes Report (.xlsx)",
            out_rc.getvalue(),
            "Major_Causes_Of_Declines.xlsx",
            key="dl_rc",
        )

# ==========================================
# 2. SUCCESS RATE RANKING & SUMMARY
# ==========================================
elif report_selection == "2. Success Rate Ranking & Summary":
  st.title("🏆 ATM Success Rate Ranking & Target Evaluation")

  target_pct = (
      st.sidebar.slider(
          "Target Threshold (%)",
          min_value=90.0,
          max_value=100.0,
          value=98.5,
          step=0.1,
      )
      / 100.0
  )

  success_file = st.file_uploader(
      "Upload ATM Success Rate Report (.xlsx / .xls)",
      type=["xlsx", "xls"],
      key="ranking_file",
  )

  if success_file is not None:
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

      # Download Ranking Excel
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
