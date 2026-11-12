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
# HELPER FUNCTIONS (Top-level scope)
# ----------------------------------------------------
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
    except Exception:
        return pd.DataFrame(columns=["FI Name", "Category", "Name", "Position", "Phone/Office Number", "Email Address"])


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

    # Fully visible radio selector listing all 5 report types explicitly
    st.markdown("**SELECT REPORT TYPE**")
    report_type = st.radio(
        "Select Report Type",
        [
            "1. ATM Decline Summary",
            "2. ATM Transaction Success Rate Summary",
            "3. ATM Success Rate Ranking",
            "4. ATM decline contribution by participant",
            "5. Major ATM declined transactions",
        ],
        label_visibility="collapsed"
    )
    st.markdown("---")

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
                                target_row_series = row
                                break
                        if target_row_series is not None:
                            header_idx = None
                            for h in range(len(df_sheet)):
                                h_str = " ".join([str(v) for v in df_sheet.iloc[h].values if pd.notna(v)]).upper()
                                if "CBE" in h_str:
                                    header_idx = h
                                    break
                            cols = df_sheet.iloc[header_idx].tolist() if header_idx is not None else df_sheet.columns.tolist()
                            vals = target_row_series.values.tolist()

                            record = {"Bank/Date": file_date}
                            for c, val in zip(cols, vals):
                                c_str = re.sub(r'\s+', ' ', str(c).strip().upper())
                                if c_str in HEADER_MAPPING:
                                    try:
                                        record[HEADER_MAPPING[c_str]] = float(str(val).replace("%", "").strip())
                                    except:
                                        record[HEADER_MAPPING[c_str]] = None
                            if len(vals) > 25:
                                try:
                                    record["Weighted Average"] = float(str(vals[25]).replace("%", "").strip())
                                except:
                                    pass
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
# 2. FI'S CONTACTS (CEO, CIO, Directors) + ADD/EDIT/SAVE
# ====================================================
elif cat_selection == "FI's Contacts":
    st.subheader("📇 FI's Contacts Management")
    
    contact_tier = st.selectbox(
        "Select Management Tier",
        ["CEO", "CIO", "Directors"]
    )
    st.markdown("---")

    if "fi_df" not in st.session_state:
        st.session_state.fi_df = load_fi_contacts()

    # Expander to add a new contact
    with st.expander("➕ Add New FI Contact", expanded=False):
        with st.form("add_contact_form", clear_on_submit=True):
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                new_fi_name = st.text_input("Bank / FI Name", placeholder="e.g., Abay Bank")
                new_name = st.text_input("Executive Name", placeholder="e.g., John Doe")
            with col_b:
                new_position = st.selectbox("Position / Category", ["CEO", "CIO", "Directors"])
                new_phone = st.text_input("Phone / Office Number", placeholder="0911000000")
            with col_c:
                new_email = st.text_input("Email Address", placeholder="john@bank.com")
            
            submitted_contact = st.form_submit_button("Save New Contact")
            if submitted_contact and new_fi_name and new_name:
                new_row = pd.DataFrame([{
                    "FI Name": new_fi_name.strip(),
                    "Category": new_position.strip(),
                    "Name": new_name.strip(),
                    "Position": new_position.strip(),
                    "Phone/Office Number": new_phone.strip(),
                    "Email Address": new_email.strip()
                }])
                st.session_state.fi_df = pd.concat([st.session_state.fi_df, new_row], ignore_index=True)
                st.success(f"Successfully added contact for **{new_fi_name}**!")

    st.markdown("### ✏️ Edit Contacts Below (Double-click any cell to edit)")

    df_full = st.session_state.fi_df.copy()
    if contact_tier == "CEO":
        mask = df_full["Category"].str.contains("CEO", case=False, na=False)
    elif contact_tier == "CIO":
        mask = df_full["Category"].str.contains("CIO", case=False, na=False)
    elif contact_tier == "Directors":
        mask = df_full["Category"].str.contains("Director", case=False, na=False)
    else:
        mask = pd.Series([True] * len(df_full))

    df_tier = df_full[mask].copy()

    search_fi = st.text_input("🔍 Search Bank / Executive Name:")
    if search_fi:
        search_mask = df_tier.apply(lambda r: r.astype(str).str.contains(search_fi, case=False).any(), axis=1)
        df_tier = df_tier[search_mask]

    # Render editable table and sync changes back to session state instantly
    edited_tier_df = st.data_editor(
        df_tier[["FI Name", "Name", "Position", "Phone/Office Number", "Email Address"]],
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key=f"fi_editor_{contact_tier}"
    )

    if not edited_tier_df.equals(df_tier[["FI Name", "Name", "Position", "Phone/Office Number", "Email Address"]]):
        df_full.loc[df_tier.index, ["FI Name", "Name", "Position", "Phone/Office Number", "Email Address"]] = edited_tier_df.values
        st.session_state.fi_df = df_full

    st.markdown("---")
    # Save/Download button to update the master Excel file
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        for cat in ["CEO", "CIO", "Directors"]:
            sub_save_df = st.session_state.fi_df[st.session_state.fi_df["Category"].str.contains(cat, case=False, na=False)]
            sub_save_df[["FI Name", "Name", "Position", "Phone/Office Number", "Email Address"]].to_excel(writer, sheet_name=cat, index=False)
    
    st.download_button(
        label="💾 Download Updated Contacts File (.xlsx)",
        data=output_excel.getvalue(),
        file_name="FI's contacts.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download this updated file and replace your root folder's 'FI's contacts.xlsx' to permanently save all edits and additions!"
    )

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
    issue_sub = st.selectbox(
        "Issue Sub-Category",
        ["1. Bank issue", "2. Internal issue"]
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
