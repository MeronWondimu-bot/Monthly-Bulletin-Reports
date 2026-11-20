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
# HELPER FUNCTIONS (Robust Data Loaders)
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
        res_df = pd.concat(all_records, ignore_index=True)
        if not res_df.empty:
            return res_df
    except Exception as e:
        st.warning(f"Could not load FI's contacts.xlsx: {e}")

    return pd.DataFrame([
        {"FI Name": "Abyssinia", "Category": "CEO", "Name": "Bekalu Zeleke", "Position": "CEO", "Phone/Office Number": "09-30-07-75-57", "Email Address": "ceo@bankofabyssinia.com"},
        {"FI Name": "Addis", "Category": "CEO", "Name": "Hailu Alemu", "Position": "CEO", "Phone/Office Number": "011-5-54-97-74", "Email Address": "Hailu@addisbanksc.com"},
        {"FI Name": "Awash", "Category": "CIO", "Name": "Tigist Mamo", "Position": "CIO", "Phone/Office Number": "011-5-57-00-00", "Email Address": "tigist@awashbank.com"},
        {"FI Name": "Dashen", "Category": "Directors", "Name": "Samuel Tadesse", "Position": "Director", "Phone/Office Number": "011-4-67-40-00", "Email Address": "samuel@dashenbank.com"}
    ])


@st.cache_data
def load_account_managers():
    try:
        filename = "Account Manager with FI's.xlsx"
        if not pd.io.common.file_exists(filename):
            filename = "Account Manager with FI's.xlsx"

        xls = pd.ExcelFile(filename)
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            df.columns = [str(c).strip() for c in df.columns]
            if "Account Manager" in df.columns and "Bank/MFI Name" in df.columns:
                df["Account Manager"] = df["Account Manager"].ffill()
                cleaned = df.dropna(subset=["Bank/MFI Name"]).reset_index(drop=True)
                if not cleaned.empty:
                    return cleaned
    except Exception as e:
        st.warning(f"Could not load Account Manager file: {e}")

    return pd.DataFrame([
        {"Account Manager": "Amare Eyase", "Bank/MFI Name": "ABAY"},
        {"Account Manager": "Amare Eyase", "Bank/MFI Name": "AWASH"},
        {"Account Manager": "Amare Eyase", "Bank/MFI Name": "NIB(NIB E-birr)"},
        {"Account Manager": "Amare Eyase", "Bank/MFI Name": "OROMIA"},
        {"Account Manager": "Biniyam Gossaye", "Bank/MFI Name": "DASHEN"},
        {"Account Manager": "Biniyam Gossaye", "Bank/MFI Name": "CBO"},
        {"Account Manager": "Meron Wondimu", "Bank/MFI Name": "CBE"},
        {"Account Manager": "Meron Wondimu", "Bank/MFI Name": "Berhan Bank"}
    ])


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


HEADER_MAPPING = {
    "CBE": "CBE", "BOA": "BOA", "AIB": "Awash", "SINQEE": "Siinqee", "OIB": "Oromia",
    "DASHEN": "Dashen", "CBO": "CBO", "AMHARA": "Amhara", "ABAY": "Abay", "BRIB": "Berhan",
    "WEGAGEN": "Wegagen", "WGB": "Wegagen", "UB": "Hibret", "NIB": "NIB", "TSEDAY": "Tseday",
    "ZEMEN": "Zemen", "BUNNA": "Bunna", "ZAMZAM": "ZamZam", "ADIB": "Addis", "HIJRA": "Hijra",
    "LION": "Lion", "ENAT": "Enat", "AHADU": "Ahadu", "DGB": "Global", "OB": "Oromia",
}


def extract_success_rate_df(files):
    success_rows = []
    for file in files:
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
                        c_str = re.sub(r"\s+", " ", str(c).strip().upper())
                        if c_str in HEADER_MAPPING:
                            try:
                                raw_num = float(str(val).replace("%", "").strip())
                                if "%" not in str(val) and 0.0 <= raw_num <= 1.0:
                                    raw_num = raw_num * 100.0
                                record[HEADER_MAPPING[c_str]] = round(raw_num, 2)
                            except Exception:
                                record[HEADER_MAPPING[c_str]] = None
                    if len(vals) > 25:
                        try:
                            wa_raw = float(str(vals[25]).replace("%", "").strip())
                            if "%" not in str(vals[25]) and 0.0 <= wa_raw <= 1.0:
                                wa_raw = wa_raw * 100.0
                            record["Weighted Average"] = round(wa_raw, 2)
                        except Exception:
                            pass
                    if len(record) > 1:
                        success_rows.append(record)
                        break
        except Exception as e:
            st.error(f"Error reading {file.name}: {e}")

    if not success_rows:
        return None

    df_success = pd.DataFrame(success_rows)
    bank_cols = [c for c in df_success.columns if c not in ["Bank/Date", "Weighted Average"]]
    ordered_cols = ["Bank/Date"] + bank_cols + (["Weighted Average"] if "Weighted Average" in df_success.columns else [])
    df_success = df_success[[c for c in ordered_cols if c in df_success.columns]]

    avg_dict = {"Bank/Date": "Average"}
    for col in [c for c in df_success.columns if c != "Bank/Date"]:
        mean_val = pd.to_numeric(df_success[col], errors="coerce").mean()
        avg_dict[col] = round(float(mean_val), 2) if pd.notna(mean_val) else None
    df_success = pd.concat([df_success, pd.DataFrame([avg_dict])], ignore_index=True)
    return df_success


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
        "Issue registration and Tracking"
    ]
)

st.markdown("---")

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
                display_df[col] = display_df[col].apply(
                    lambda x: f"<span style='color: red; font-weight: bold;'>{int(float(x)):,}</span>" if float(x) > 100 else f"{int(float(x)):,}"
                )

            st.markdown(display_df.to_html(escape=False, index=False), unsafe_allow_html=True)

            out_sum = io.BytesIO()
            with pd.ExcelWriter(out_sum, engine="openpyxl") as writer:
                summary_pivot.to_excel(writer, sheet_name="ATM Decline Summary", index=False)
            st.download_button("📥 Download ATM Decline Summary (.xlsx)", out_sum.getvalue(), "ATM_Decline_Summary.xlsx")
        else:
            st.info("👆 Please upload your daily ATM decline files at the top.")

    elif report_type == "2. ATM Transaction Success Rate Summary":
        st.subheader("📈 ATM Transaction Success Rate Summary")
        if uploaded_files:
            df_success = extract_success_rate_df(uploaded_files)
            if df_success is not None and not df_success.empty:
                num_cols = [c for c in df_success.columns if c != "Bank/Date"]
                formatted_success = df_success.copy()
                for col in num_cols:
                    formatted_success[col] = pd.to_numeric(formatted_success[col], errors="coerce").apply(
                        lambda x: f"{x:.2f}%" if pd.notna(x) else ""
                    )

                st.dataframe(formatted_success, use_container_width=True, hide_index=True)

                out_succ = io.BytesIO()
                with pd.ExcelWriter(out_succ, engine="openpyxl") as writer:
                    formatted_success.to_excel(writer, sheet_name="Success Rate Summary (%)", index=False)
                st.download_button(
                    "📥 Download Success Rate Summary (.xlsx)",
                    out_succ.getvalue(),
                    "ATM_Success_Rate_Summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Could not locate 'SUCC. RATE' rows in the uploaded files.")
        else:
            st.info("👆 Please upload your daily source report files at the top.")

    elif report_type == "3. ATM Success Rate Ranking":
        st.subheader("🏆 ATM Success Rate Ranking")
        if uploaded_files:
            df_success = extract_success_rate_df(uploaded_files)
            if df_success is not None and not df_success.empty:
                avg_rows = df_success[df_success["Bank/Date"] == "Average"]
                source_avg = avg_rows.iloc[0] if not avg_rows.empty else df_success.drop(columns=["Bank/Date"]).mean(numeric_only=True)
                bank_cols = [c for c in df_success.columns if c not in ["Bank/Date", "Weighted Average"]]

                ranking_rows = []
                for bank in bank_cols:
                    val = pd.to_numeric(source_avg.get(bank), errors="coerce")
                    if pd.notna(val):
                        ranking_rows.append({
                            "Participant": bank,
                            "Average Success Rate (%)": round(float(val), 2),
                            "vs. 98.5% Target": "Meets Target" if float(val) >= 98.5 else "Below Target"
                        })

                if ranking_rows:
                    rank_df = pd.DataFrame(ranking_rows).sort_values(
                        by="Average Success Rate (%)", ascending=False
                    ).reset_index(drop=True)
                    rank_df.insert(0, "Rank", rank_df.index + 1)

                    display_rank = rank_df.copy()
                    display_rank["Average Success Rate (%)"] = display_rank["Average Success Rate (%)"].apply(lambda x: f"{x:.2f}%")

                    st.dataframe(display_rank, use_container_width=True, hide_index=True)

                    out_rank = io.BytesIO()
                    with pd.ExcelWriter(out_rank, engine="openpyxl") as writer:
                        display_rank.to_excel(writer, sheet_name="ATM Success Rate Ranking", index=False)
                    st.download_button(
                        "📥 Download Success Rate Ranking (.xlsx)",
                        out_rank.getvalue(),
                        "ATM_Success_Rate_Ranking.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No participant success rate values found to rank.")
            else:
                st.warning("Could not locate 'SUCC. RATE' rows in the uploaded files.")
        else:
            st.info("👆 Please upload your daily source report files at the top.")

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

    elif report_type == "5. Major ATM declined transactions":
        st.subheader("🔍 Major ATM Declined Transactions")
        if combined_decline is not None and not combined_decline.empty:
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

            rc_grouped = combined_decline.groupby(rc_col)[valid_cols].sum().reset_index()
            rc_grouped.rename(columns={rc_col: "Response Code / Decline Reason"}, inplace=True)
            rc_grouped["Total Declines"] = rc_grouped[valid_cols].sum(axis=1)
            grand_total = rc_grouped["Total Declines"].sum()
            rc_grouped["Share (%)"] = rc_grouped["Total Declines"].apply(
                lambda x: f"{(x / grand_total * 100):.2f}%" if grand_total > 0 else "0.00%"
            )
            rc_grouped = rc_grouped.sort_values(by="Total Declines", ascending=False).reset_index(drop=True)

            top_n = st.slider("Show Top Decline Reasons", min_value=5, max_value=max(5, len(rc_grouped)), value=min(10, len(rc_grouped)))
            major_rc_df = rc_grouped.head(top_n).copy()
            major_rc_df.insert(0, "Rank", major_rc_df.index + 1)

            for col in valid_cols + ["Total Declines"]:
                major_rc_df[col] = major_rc_df[col].apply(lambda x: f"{int(float(x)):,}")

            st.markdown("### Top Declined Response Codes")
            st.dataframe(major_rc_df[["Rank", "Response Code / Decline Reason", "Total Declines", "Share (%)"] + valid_cols], use_container_width=True, hide_index=True)

            bank_contrib = []
            for col in valid_cols:
                b_total = pd.to_numeric(combined_decline[col], errors="coerce").fillna(0).sum()
                top_rc_idx = combined_decline.groupby(rc_col)[col].sum().idxmax() if b_total > 0 else "N/A"
                top_rc_val = combined_decline.groupby(rc_col)[col].sum().max() if b_total > 0 else 0
                bank_contrib.append({
                    "Participant": col,
                    "Total Declines": int(b_total),
                    "Share of Total (%)": round(float(b_total / grand_total * 100), 2) if grand_total > 0 else 0.0,
                    "Top Decline Reason": str(top_rc_idx),
                    "Top Reason Count": int(top_rc_val)
                })

            bank_major_df = pd.DataFrame(bank_contrib).sort_values(by="Total Declines", ascending=False).reset_index(drop=True)
            bank_major_df["Total Declines"] = bank_major_df["Total Declines"].apply(lambda x: f"{x:,}")
            bank_major_df["Share of Total (%)"] = bank_major_df["Share of Total (%)"].apply(lambda x: f"{x:.2f}%")
            bank_major_df["Top Reason Count"] = bank_major_df["Top Reason Count"].apply(lambda x: f"{x:,}")

            st.markdown("### Major Decline Drivers by Participant")
            st.dataframe(bank_major_df, use_container_width=True, hide_index=True)

            out_major = io.BytesIO()
            with pd.ExcelWriter(out_major, engine="openpyxl") as writer:
                rc_grouped.to_excel(writer, sheet_name="Top Decline Reasons", index=False)
                bank_major_df.to_excel(writer, sheet_name="By Participant", index=False)
            st.download_button(
                "📥 Download Major Declined Transactions (.xlsx)",
                out_major.getvalue(),
                "Major_ATM_Declined_Transactions.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("👆 Please upload your daily ATM decline files at the top.")

# ====================================================
# 2. FI'S CONTACTS
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

    st.markdown("### Edit Contacts Below")

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
    st.subheader("📋 Account Managers & FI Portfolio Management")

    if "am_df" not in st.session_state:
        st.session_state.am_df = load_account_managers()

    with st.expander("➕ Add New Account Manager Portfolio Assignment", expanded=False):
        with st.form("add_am_form", clear_on_submit=True):
            col_1, col_2 = st.columns(2)
            with col_1:
                am_name = st.text_input("Account Manager Name", placeholder="e.g., Meron Wondimu")
            with col_2:
                am_bank = st.text_input("Bank / MFI Name", placeholder="e.g., Dashen Bank")

            submitted_am = st.form_submit_button("Save Assignment")
            if submitted_am and am_name and am_bank:
                new_am_row = pd.DataFrame([{
                    "Account Manager": am_name.strip(),
                    "Bank/MFI Name": am_bank.strip()
                }])
                st.session_state.am_df = pd.concat([st.session_state.am_df, new_am_row], ignore_index=True)
                st.success(f"Successfully assigned **{am_bank}** to **{am_name}**!")

    st.markdown("### Edit Account Managers Below")

    search_am = st.text_input("🔍 Search Account Manager / Bank Name:")
    am_view = st.session_state.am_df.copy()
    if search_am:
        am_mask = am_view.apply(lambda r: r.astype(str).str.contains(search_am, case=False).any(), axis=1)
        am_view = am_view[am_mask]

    edited_am_df = st.data_editor(
        am_view,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        key="am_editor"
    )

    if not edited_am_df.equals(am_view):
        st.session_state.am_df = edited_am_df

    st.markdown("---")
    am_excel = io.BytesIO()
    with pd.ExcelWriter(am_excel, engine="openpyxl") as writer:
        st.session_state.am_df.to_excel(writer, sheet_name="Sheet1", index=False)

    st.download_button(
        label="💾 Download Updated Account Managers File (.xlsx)",
        data=am_excel.getvalue(),
        file_name="Account Manager with FI's.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download this updated file to replace your root folder's 'Account Manager with FI's.xlsx'!"
    )

# ====================================================
# 4. ISSUE REGISTRATION AND TRACKING
# ====================================================
elif cat_selection == "Issue registration and Tracking":
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

    issue_cols = [
        "Issue ID", "Type", "Date Logged", "Entity / Department",
        "Assigned Department", "Severity", "Description", "Status"
    ]

    if "issues_df" not in st.session_state:
        st.session_state.issues_df = pd.DataFrame(columns=issue_cols)
    else:
        for col in issue_cols:
            if col not in st.session_state.issues_df.columns:
                st.session_state.issues_df[col] = ""
        st.session_state.issues_df = st.session_state.issues_df[issue_cols]

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
                    "Assigned Department": dept,
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

    st.markdown("---")
    issues_excel = io.BytesIO()
    with pd.ExcelWriter(issues_excel, engine="openpyxl") as writer:
        st.session_state.issues_df.to_excel(writer, sheet_name="All Issues", index=False)
        for sub_cat in ["1. Bank issue", "2. Internal issue"]:
            sub_df = st.session_state.issues_df[st.session_state.issues_df["Type"] == sub_cat]
            safe_sheet = sub_cat[:31]
            sub_df.to_excel(writer, sheet_name=safe_sheet, index=False)

    st.download_button(
        label="📥 Download Issues Log (.xlsx)",
        data=issues_excel.getvalue(),
        file_name="EthSwitch_Issues_Log.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download all logged issues (including Bank and Internal issues) as an Excel workbook."
    )
