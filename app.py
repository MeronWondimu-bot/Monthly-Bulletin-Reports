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

# Custom CSS matching your exact design preference
st.markdown("""
    <style>
        .main-title {
            text-align: center;
            font-size: 28px;
            font-weight: 800;
            color: #d35400;
            letter-spacing: 1px;
        }
        .hero-banner {
            background: linear-gradient(0deg, rgba(249,235,221,0.5) 0%, rgba(255,255,255,1) 100%);
            border: 1px solid #f9ebdd;
            padding: 25px;
            border-radius: 8px;
            text-align: center;
            margin-bottom: 25px;
        }
        .hero-title {
            font-size: 22px;
            font-weight: 700;
            color: #d35400;
            text-transform: uppercase;
            margin-bottom: 5px;
        }
        .hero-subtitle {
            font-size: 14px;
            color: #666;
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

st.markdown("""
  
""", unsafe_allow_html=True)

# ----------------------------------------------------
# TIER 1: MAIN NAVIGATION CATEGORY
# ----------------------------------------------------
st.markdown("**REPORT CATEGORY**")
cat_selection = st.radio(
    "Report Category",
    [
        "Bi-weekly and monthly reports",
        "FI_Contacts",
        "Account_Managers",
        "Issue_Tracking"
    ],
    horizontal=True,
    label_visibility="collapsed"
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
# 1. BI-WEEKLY AND MONTHLY REPORTS
# ====================================================
if cat_selection == "Bi-weekly and monthly reports":
    st.markdown("**UPLOAD DAILY SOURCE REPORTS**")
    uploaded_files = st.file_uploader(
        "Upload all daily ATM report files for the month below:",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )
    st.markdown("---")

    report_type = st.radio(
        "Report Selection",
        [
            "1. ATM Decline Summary",
            "2. ATM Transaction Success Rate Summary",
            "3. ATM Success Rate Ranking",
            "4. ATM decline contribution by participant",
            "5. Major ATM declined transactions",
        ],
        horizontal=True,
        label_visibility="collapsed"
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
                            if "SUCC" in " ".join([str(v) for v in row.values if pd.notna(v)]).upper() and "RATE" in " ".join([str(v) for v in row.values if pd.notna(v)]).upper():
                                target_row_series = row
                                break
                        if target_row_series is not None:
                            header_idx = next((h for h in range(len(df_sheet)) if "CBE" in " ".join([str(v) for v in df_sheet.iloc[h].values if pd.notna(v)]).upper()), None)
                            cols = df_sheet.iloc[header_idx].tolist() if header_idx is not None else df_sheet.columns.tolist()
                            vals = target_row_series.values.tolist()

                            record = {"Bank/Date": file_date}
                            for c, val in zip(cols, vals):
                                c_str = re.sub(r'\s+', ' ', str(c).strip().upper())
                                if c_str in HEADER_MAPPING:
                                    try: record[HEADER_MAPPING[c_str]] = float(str(val).replace("%", "").strip())
                                    except: record[HEADER_MAPPING[c_str]] = None
                            if len(vals) > 25:
                                try: record["Weighted Average"] = float(str(vals[25]).replace("%", "").strip())
                                except: pass
                            if len(record) > 1:
                                success_rows.append(record)
                                break
                except Exception as e:
                    st.error(f"Error reading {file.name}: {e}")

            if success_rows:
                df_success = pd.DataFrame(success_rows)
                bank_cols = [c for c in df_success.columns if c not in ["Bank/Date", "Weighted Average"]]
                ordered_cols = ["Bank/Date"] + bank_cols + (["Weighted Average"] if "Weighted Average" in df_success.columns else [])
                df_success = df_success[[c for c in ordered_cols if c in df_success.columns]]

                avg_dict = {"Bank/Date": "Average"}
                for col in [c for c in df_success.columns if c != "Bank/Date"]:
                    avg_dict[col] = pd.to_numeric(df_success[col], errors="coerce").mean()
                df_success = pd.concat([df_success, pd.DataFrame([avg_dict])], ignore_index=True)

                st.dataframe(df_success, use_container_width=True)
            else:
                st.warning("Could not locate 'SUCC. RATE' rows in the uploaded files.")
        else:
            st.info("👆 Please upload your daily source report files at the top.")

    # Report 3
    elif report_type == "3. ATM Success Rate Ranking":
        st.subheader("🏆 ATM Success Rate Ranking")
        st.info("Report 3 is ready for review once Report 2 is locked down.")

    # Report 4
    elif report_type == "4. ATM decline contribution by participant":
        st.subheader("👤 ATM Decline Contribution by Participant")
        if combined_decline is not None:
            bank_totals = {}
            for col in combined_decline.columns:
                col_clean = str(col).strip().upper()
                if "GEDAA" not in col_clean and "GADAA" not in col_clean and col_clean in HEADER_MAPPING:
                    bank_totals[HEADER_MAPPING[col_clean]] = bank_totals.get(HEADER_MAPPING[col_clean], 0) + pd.to_numeric(combined_decline[col], errors="coerce").fillna(0).sum()
            part_df = pd.DataFrame(list(bank_totals.items()), columns=["Participant", "Declined transactions"]).sort_values(by="Declined transactions", ascending=False).reset_index(drop=True)
            st.dataframe(part_df, use_container_width=True)
        else:
            st.info("👆 Please upload your daily decline Excel files at the top.")

    # Report 5
    elif report_type == "5. Major ATM declined transactions":
        st.subheader("🔍 Major ATM Declined Transactions")
        st.info("Major ATM declined transactions module active.")

# ====================================================
# 2. FI_CONTACTS
# ====================================================
elif cat_selection == "FI_Contacts":
    st.subheader("📇 FI's Contacts Directory")
    @st.cache_data
    def load_fi_contacts():
        try:
            xls = pd.ExcelFile("FI's contacts.xlsx")
            all_records = []
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, skiprows=1).iloc[:, 1:6].copy()
                df.columns = ["FI Name", "Name", "Position", "Phone/Office Number", "Email Address"]
                df["FI Name"] = df["FI Name"].ffill()
                df["Category"] = sheet.strip()
                all_records.append(df)
            return pd.concat(all_records, ignore_index=True)
        except:
            return pd.DataFrame(columns=["FI Name", "Category", "Name", "Position", "Phone/Office Number", "Email Address"])

    if "fi_df" not in st.session_state:
        st.session_state.fi_df = load_fi_contacts()

    search_fi = st.text_input("🔍 Search Bank / Executive Name:")
    df_view = st.session_state.fi_df.copy()
    if search_fi:
        df_view = df_view[df_view.apply(lambda r: r.astype(str).str.contains(search_fi, case=False).any(), axis=1)]
    st.dataframe(df_view, use_container_width=True, hide_index=True)

# ====================================================
# 3. ACCOUNT_MANAGERS
# ====================================================
elif cat_selection == "Account_Managers":
    st.subheader("📋 Account Managers List")
    if "am_df" not in st.session_state:
        st.session_state.am_df = pd.DataFrame([
            {"Bank / FI Name": "Abyssinia Bank", "Account Manager": "Abebe Kebede", "Phone": "0911000001", "Email": "abebe@ethswitch.com", "Tier": "Tier 1"},
            {"Bank / FI Name": "Awash Bank", "Account Manager": "Tigist Mamo", "Phone": "0911000002", "Email": "tigist@ethswitch.com", "Tier": "Tier 1"},
        ])
    st.dataframe(st.session_state.am_df, use_container_width=True, hide_index=True)

# ====================================================
# 4. ISSUE_TRACKING (Bank issue & Internal issue)
# ====================================================
elif cat_selection == "Issue_Tracking":
    st.subheader("🛠️ Issue Tracking List")
    issue_sub = st.radio(
        "Issue Sub-Category",
        ["1. Bank issue", "2. Internal issue"],
        horizontal=True,
        label_visibility="collapsed"
    )
    st.markdown("---")

    DEPARTMENTS = [
        "Application", "Digital", "Network", "System", "Shared Service",
        "Settlement and Reconciliation", "PIN and Card Department",
        "Security", "Dispute and Fraud", "Database"
    ]

    if "issues_df" not in st.session_state:
        st.session_state.issues_df = pd.DataFrame(columns=[
            "Issue ID", "Type", "Date Logged", "Entity / Department", "Severity", "Description", "Status"
        ])

    with st.expander(f"➕ Log New {issue_sub}", expanded=True):
        with st.form(f"form_{issue_sub}", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                entity_name = st.text_input("Bank / FI Name" if "Bank" in issue_sub else "Internal Department", placeholder="e.g., Dashen Bank or Network Team")
                dept = st.selectbox("Assigned Department", DEPARTMENTS)
            with c2:
                severity = st.selectbox("Severity", ["Low", "Medium", "High", "Critical"])
                status = st.selectbox("Status", ["Open", "In Progress", "Resolved", "Closed"])
            with c3:
                log_date = st.date_input("Date Logged", value=date.today())
            
            description = st.text_area("Issue Description", placeholder="Describe the problem...")
            submitted = st.form_submit_button("Submit Issue")
            
            if submitted and entity_name and description:
                issue_id = f"ISS-{len(st.session_state.issues_df) + 1:03d}"
                new_row = pd.DataFrame([{
                    "Issue ID": issue_id,
                    "Type": issue_sub,
                    "Date Logged": str(log_date),
                    "Entity / Department": entity_name.strip(),
                    "Severity": severity,
                    "Description": description.strip(),
                    "Status": status
                }])
                st.session_state.issues_df = pd.concat([st.session_state.issues_df, new_row], ignore_index=True)
                st.success(f"Logged {issue_sub} **{issue_id}** successfully!")

    st.divider()
    st.subheader(f"📋 {issue_sub} List")
    
    filtered_issues = st.session_state.issues_df[st.session_state.issues_df["Type"] == issue_sub]
    st.dataframe(filtered_issues, use_container_width=True, hide_index=True)
