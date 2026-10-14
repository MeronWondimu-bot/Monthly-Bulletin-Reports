import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="ATM Consolidated Reporting Dashboard", layout="wide"
)

# Sidebar Navigation Menu
st.sidebar.title("📊 ATM Reports Menu")
report_selection = st.sidebar.radio(
    "Select Report:",
    [
        "Home / Overview",
        "1. Decline Contribution by Participant",
        "2. Major Causes of Declines",
        "3. Success Rate Ranking (>= 98.5%)",
    ],
)

# ==========================================
# HOME / OVERVIEW
# ==========================================
if report_selection == "Home / Overview":
  st.title("💳 EthSwitch ATM Performance & Decline Central Dashboard")
  st.markdown("""
    Welcome to your unified automated reporting dashboard. 
    Use the sidebar menu to navigate between your different monthly and daily reports:
    * **Participant Contribution:** View which banks contribute most to ATM declines.
    * **Major Causes:** Analyze decline transaction volumes grouped by response codes.
    * **Success Rate Ranking:** Evaluate participants against the **98.5% success rate target**.
    """)

# ==========================================
# REPORT 1: PARTICIPANT CONTRIBUTION
# ==========================================
elif report_selection == "1. Decline Contribution by Participant":
  st.title("📊 ATM Decline Contribution by Participant")
  decline_file = st.file_uploader(
      "Upload ATM Decline Report (.xlsx)", type=["xlsx"], key="part"
  )

  if decline_file is not None:
    xls_decline = pd.ExcelFile(decline_file)
    daily_frames = []
    for sheet_name in xls_decline.sheet_names:
      if sheet_name.lower() == "sheet16":
        continue
      df_s = pd.read_excel(decline_file, sheet_name=sheet_name)
      if len(df_s) > 2:
        df_s.columns = df_s.iloc[2]
        sub_df = df_s.iloc[3:].dropna(subset=["RC/BANK NAME"])
        daily_frames.append(sub_df)

    if daily_frames:
      combined_decline = pd.concat(daily_frames, ignore_index=True)
      header_mapping = {
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
      }

      bank_totals = {}
      for col, target_name in header_mapping.items():
        if col in combined_decline.columns:
          vals = pd.to_numeric(
              combined_decline[col], errors="coerce"
          ).fillna(0)
          bank_totals[target_name] = vals.sum()

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
      )

      display_part = part_df.copy()
      display_part["Declined transactions"] = display_part[
          "Declined transactions"
      ].apply(lambda x: f"{int(x):,}")
      display_part["Share of ATM declines"] = display_part[
          "Share of ATM declines"
      ].apply(lambda x: f"{x:.2%}")

      st.dataframe(display_part, use_container_width=True)

# ==========================================
# REPORT 2: MAJOR CAUSES OF DECLINES
# ==========================================
elif report_selection == "2. Major Causes of Declines":
  st.title("🔍 Major Causes of ATM Decline Transactions")
  decline_file2 = st.file_uploader(
      "Upload ATM Decline Report (.xlsx)", type=["xlsx"], key="causes"
  )

  if decline_file2 is not None:
    xls_decline2 = pd.ExcelFile(decline_file2)
    daily_frames2 = []
    for sheet_name in xls_decline2.sheet_names:
      if sheet_name.lower() == "sheet16":
        continue
      df_s = pd.read_excel(decline_file2, sheet_name=sheet_name)
      if len(df_s) > 2:
        df_s.columns = df_s.iloc[2]
        sub_df = df_s.iloc[3:].dropna(subset=["RC/BANK NAME"])
        daily_frames2.append(sub_df)

    if daily_frames2:
      combined_decline2 = pd.concat(daily_frames2, ignore_index=True)
      total_declines2 = pd.to_numeric(
          combined_decline2["Total"], errors="coerce"
      ).sum()

      rc_col = "RC/BANK NAME"
      if rc_col in combined_decline2.columns:
        rc_summary = (
            combined_decline2.groupby(rc_col)["Total"]
            .apply(lambda x: pd.to_numeric(x, errors="coerce").sum())
            .reset_index()
        )
        rc_summary.columns = ["Response code", "Declined transactions"]
        rc_summary = rc_summary.sort_values(
            by="Declined transactions", ascending=False
        ).reset_index(drop=True)
        rc_summary["Share of ATM declines"] = (
            rc_summary["Declined transactions"] / total_declines2
        )

        display_rc = rc_summary.copy()
        display_rc["Declined transactions"] = display_rc[
            "Declined transactions"
        ].apply(lambda x: f"{int(x):,}")
        display_rc["Share of ATM declines"] = display_rc[
            "Share of ATM declines"
        ].apply(lambda x: f"{x:.2%}")

        st.dataframe(display_rc, use_container_width=True)

# ==========================================
# REPORT 3: SUCCESS RATE RANKING
# ==========================================
elif report_selection == "3. Success Rate Ranking (>= 98.5%)":
  st.title("🏆 ATM Success Rate Ranking & Target Evaluation")
  success_file = st.file_uploader(
      "Upload ATM Success Rate Report (.xlsx)", type=["xlsx"], key="ranking"
  )

  if success_file is not None:
    df_success = pd.read_excel(success_file, sheet_name=0)
    cols = df_success.iloc[3].tolist()
    avg_row = df_success.iloc[34].tolist()

    success_data = []
    for c, v in zip(cols[1:], avg_row[1:]):
      if pd.notna(c) and c != "Weighted Average":
        success_data.append({"Participant": c, "Success rate": float(v)})

    rank_df = pd.DataFrame(success_data)
    rank_df["Target status"] = rank_df["Success rate"].apply(
        lambda x: "Meets target" if x >= 0.985 else "Below target"
    )
    rank_df = rank_df.sort_values(
        by="Success rate", ascending=False
    ).reset_index(drop=True)
    rank_df.index += 1
    rank_df["Rank"] = rank_df.index
    rank_df = rank_df[["Rank", "Participant", "Success rate", "Target status"]]

    display_rank = rank_df.copy()
    display_rank["Success rate"] = display_rank["Success rate"].apply(
        lambda x: f"{x*100:.2f}%"
    )

    st.dataframe(display_rank, use_container_width=True)