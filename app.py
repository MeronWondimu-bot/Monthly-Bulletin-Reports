import io
import re
import pandas as pd
import streamlit as st
from datetime import date

# Page Configuration
st.set_page_config(
    page_title="EthSwitch CRM Dashboard",
    page_icon="💳",
    layout="wide",
)

# Custom CSS for UI styling
st.markdown("""
    <style>
        .main-title {
            text-align: center;
            font-size: 28px;
            font-weight: 800;
            color: #d35400;
            letter-spacing: 1px;
            margin-bottom: 20px;
        }
    </style>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# HEADER WITH LOGOS / TITLES
# ----------------------------------------------------
col1, col2, col3 = st.columns([1.5, 3, 1.5])
with col1:
    st.markdown("### **ETHSWITCH**")
with col2:
    st.markdown('<div class="main-title">EthSwitch CRM Dashboard</div>', unsafe_allow_html=True)
with col3:
    st.markdown("<h3 style='text-align: right;'>EthioPay</h3>", unsafe_allow_html=True)

st.markdown("---")

# ----------------------------------------------------
# CRM OVERALL TASKS (CLEAN SELECTBOX SELECTOR)
# ----------------------------------------------------
cat_selection = st.selectbox(
    "CRM Overall Tasks",
    [
        "Bi- weekly and Monthly performance reports",
        "FI's Contacts",
        "Account_Managers",
        "Issue_Tracking"
    ]
)

st.markdown("---")

HEADER_MAPPING = {
    "CBE": "CBE", "BOA": "BOA", "AIB": "Awash", "SINQEE": "Siinqee", "OIB": "Oromia",
    "DASHEN": "Dashen", "CBO": "CBO", "AMHARA": "Amhara", "ABAY": "Abay", "BRIB": "Berhan",
    "WEGAGEN": "Wegagen", "WGB": "Wegagen", "UB": "Hibret", "NIB": "NIB", "TSEDAY": "Tseday",
    "ZEMEN": "Zemen", "BUNNA": "Bunna", "ZAMZAM": "ZamZam", "ADIB": "Addis", "HIJRA": "Hijra",
    "LION": "Lion", "ENAT": "Enat", "AHADU": "Ahadu", "DGB": "Global", "OB": "Oromia",
}

# ====================================================
# 1. BI-WEEKLY AND MONTHLY PERFORMANCE REPORTS
# ====================================================
if cat_selection == "Bi- weekly and Monthly performance reports":
    st.markdown("**UPLOAD DAILY SOURCE REPORTS**")
    uploaded_files = st.file_uploader(
        "Upload all daily ATM report files for the month below:",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )
    st.markdown("---")

    report_type = st.selectbox(
        "Select Report Type",
        [
            "1. ATM Decline Summary",
            "2. ATM Transaction Success Rate Summary",
            "3. ATM Success Rate Ranking",
            "4. ATM decline contribution by participant",
            "5. Major ATM declined transactions",
        ]
    )
    st.markdown("---")

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
                        df_sheet.columns = [str(c).strip() for c in df_sheet.iloc[header_idx].values]
                        sub_df = df_sheet.iloc[header_idx + 1:].copy()
                        rc_col = sub_df.columns[0]
                        sub_df = sub_df.dropna(subset=[rc_col])
                        sub_df = sub_df[~sub_df[rc_col].astype(str).str.contains("Total|RC|Response Code|Issu", case=False, na=False)]
                        all_frames.append(sub_df)
            except Exception as e:
                st.error(f"Error processing {file.name}: {e}")
        return pd.concat(all_frames, ignore_index=True) if all_frames else None

    combined_decline = process_all_decline_files(uploaded_files) if uploaded_files else None

    # Report 1
    if report_type == "1. ATM Decline Summary":
        st.subheader("📊 ATM Decline Summary")
        if combined_decline is not None:
            rc_col = combined_decline.columns[0]
            valid_cols = []
            for c in combined_decline.columns:
                c_str = str(c).strip().upper()
                if c != rc_col and "GEDAA" not in c_str and "GADAA" not in c_str and "TOTAL" not in c_str:
                    clean_name = HEADER_MAPPING.get(c_str, c_str)
                    if clean_name not in valid_cols:
                        valid_cols.append(clean_name)
                        combined_decline.rename(columns={c: clean_name}, inplace=True)

            for col in valid_cols:
                combined_decline[col] = pd.to_numeric(combined_decline[col], errors="coerce").fillna(0)

            summary_pivot = combined_decline.groupby(rc_col)[valid_cols].sum().reset_index()
            summary_pivot.rename(columns={rc_col: "RC"}, inplace=True)
            summary_pivot["Total"] = summary_pivot[valid_cols].sum(axis=1)

            total_row = {"RC": "TOTAL"}
            for col in valid_cols + ["Total"]:
                total_row[col] = summary_pivot[col].sum()
            summary_pivot = pd.concat([summary_pivot, pd.DataFrame([total_row])], ignore_index=True)

            display_df = summary_pivot.copy()
            for col in valid_cols + ["Total"]:
                display_df[col] = display_df[col].apply(lambda x: f"<span style='color: red; font-weight: bold;'>{int(float(x)):,}</span>" if float(x) > 100 else f"{int(float(x)):,}")

            st.markdown(display_df.to_html(escape=False, index=False), unsafe_allow_html=True)

            out_sum = io.BytesIO()
            with pd.ExcelWriter(out_sum, engine="openpyxl") as writer:
                summary_pivot.to_excel(writer, sheet_name="ATM Decline Summary", index=False)
            st.download_button("📥 Download ATM Decline Summary (.xlsx)", out_sum.getvalue(), "ATM_Decline_Summary.xlsx")
        else:
            st.info("👆 Please upload your daily ATM decline files at the top.")

    # Report 2
    elif report_type == "2. ATM Transaction Success Rate Summary":
        st.subheader("📈 ATM Transaction Success Rate Summary")
        if uploaded_files:
            success_rows = []
            for file in uploaded_files:
                try:
                    xls = pd.ExcelFile(file)
                    cleaned_name = file.name.split(".")[0].replace("ATM Declined Transaction Report for", "").replace("ATM Success Rate", "").strip()
                    file_date = re.sub(r",\s*", ", ", cleaned_name)

                    for sheet_name in xls.sheet_names:
                        df_sheet = pd.read_excel(file, sheet_name=sheet_name)
                        target_row_series = None
                        for idx, row in df_sheet.iterrows():
                            row_vals_str = " ".join([str(v) for v in row.values if pd.notna(v)]).upper()
                            if "SUCC" in row_vals_str and "RATE" in row_vals_str:
                                target_row
