import streamlit as st
import pandas as pd
import io
import os

st.set_page_config(page_title="FI's Contacts", layout="wide")

st.markdown("""
    <style>
    .crm-banner {
        background-color: #C84B00;
        color: white;
        padding: 18px;
        border-radius: 8px;
        text-align: center;
        font-size: 24px;
        font-weight: 700;
        margin-bottom: 24px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="crm-banner">CRM Projects - FI\'s Contacts</div>', unsafe_allow_html=True)

# ---------------------------------------------------------
# 1. PARSE & CONSOLIDATE MULTI-SHEET EXCEL FILE
# ---------------------------------------------------------
@st.cache_data
def parse_fi_excel(file_source) -> pd.DataFrame:
    xls = pd.ExcelFile(file_source)
    all_records = []

    for sheet in xls.sheet_names:
        # Header row is on row index 1 (skiprows=1)
        df = pd.read_excel(xls, sheet_name=sheet, skiprows=1)
        
        # Standardize expected columns (ignoring the first index number column)
        df = df.iloc[:, 1:6].copy()
        df.columns = ["FI Name", "Name", "Position", "Phone/Office Number", "Email Address"]
        
        # Forward-fill Bank/FI names where multiple executives are listed under one FI
        df["FI Name"] = df["FI Name"].ffill()
        df["Category"] = sheet.strip()
        
        # Clean whitespace
        for col in ["FI Name", "Name", "Position", "Phone/Office Number", "Email Address"]:
            df[col] = df[col].astype(str).replace({"nan": "", "None": ""}).str.strip()
            
        all_records.append(df)

    combined = pd.concat(all_records, ignore_index=True)
    # Remove completely empty rows where FI Name is missing
    combined = combined[combined["FI Name"] != ""].reset_index(drop=True)
    return combined

def to_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="FI_Contacts")
    return output.getvalue()

# ---------------------------------------------------------
# 2. INITIALIZE DATA FROM REPO FILE OR UPLOAD
# ---------------------------------------------------------
DEFAULT_FILE = "FI's contacts.xlsx"

if "fi_contacts_df" not in st.session_state:
    if os.path.exists(DEFAULT_FILE):
        st.session_state.fi_contacts_df = parse_fi_excel(DEFAULT_FILE)
    else:
        st.session_state.fi_contacts_df = pd.DataFrame(
            columns=["FI Name", "Category", "Name", "Position", "Phone/Office Number", "Email Address"]
        )

# Optional: Re-upload updated master Excel file
with st.expander("📤 Upload / Replace Master Excel File (Optional)"):
    uploaded_file = st.file_uploader("Upload FI's contacts.xlsx", type=["xlsx", "xls"])
    if uploaded_file:
        st.session_state.fi_contacts_df = parse_fi_excel(uploaded_file)
        st.success("Master FI Contacts loaded from uploaded file!")

# ---------------------------------------------------------
# 3. ADD NEW FI OR CONTACT DETAILS FORM
# ---------------------------------------------------------
with st.expander("➕ Add New FI / Contact Details", expanded=False):
    with st.form("add_fi_contact_form", clear_on_submit=True):
        existing_fis = sorted(st.session_state.fi_contacts_df["FI Name"].dropna().unique().tolist())
        
        c1, c2, c3 = st.columns(3)
        with c1:
            fi_choice = st.selectbox(
                "Select Existing FI (or choose '➕ Add New FI')",
                options=["➕ Add New FI"] + existing_fis
            )
            new_fi_name = st.text_input("If New FI, Enter FI / Bank Name:", placeholder="e.g., New Bank Ethiopia")
        with c2:
            category = st.selectbox("Management Category", ["CEO", "CIO,CTO,VP", "Director", "Manager / Other"])
            contact_name = st.text_input("Executive / Contact Name", placeholder="e.g., Ato Abebe Kebede")
        with c3:
            position = st.text_input("Position / Title", placeholder="e.g., Director, Digital Banking")
            phone_num = st.text_input("Phone / Office Number", placeholder="e.g., 09-11-00-00-00")

        email_addr = st.text_input("Email Address", placeholder="e.g., name@bank.com.et")

        submitted = st.form_submit_button("Save FI Contact")
        if submitted:
            final_fi = new_fi_name.strip() if fi_choice == "➕ Add New FI" else fi_choice
            if not final_fi:
                st.warning("Please select an existing FI or type a new FI name.")
            else:
                new_row = pd.DataFrame([{
                    "FI Name": final_fi,
                    "Category": category,
                    "Name": contact_name.strip(),
                    "Position": position.strip(),
                    "Phone/Office Number": phone_num.strip(),
                    "Email Address": email_addr.strip()
                }])
                st.session_state.fi_contacts_df = pd.concat(
                    [st.session_state.fi_contacts_df, new_row], ignore_index=True
                )
                st.success(f"Added contact for **{final_fi}**!")

st.divider()

# ---------------------------------------------------------
# 4. SEARCH & FILTER DIRECTORY
# ---------------------------------------------------------
st.subheader("📇 FI Contact Directory")

col_f1, col_f2, col_f3 = st.columns([2, 2, 1])
with col_f1:
    search_fi = st.text_input("🔍 Find by FI's Name:", placeholder="e.g., Abyssinia, CBE, Dashen, Telebirr...")
with col_f2:
    search_pos = st.text_input("🔍 Find by Position:", placeholder="e.g., CEO, CIO, Digital, Infrastructure...")
with col_f3:
    cat_filter = st.selectbox("Filter by Tier:", ["All", "CEO", "CIO,CTO,VP", "Director", "Manager / Other"])

hide_blank_contacts = st.checkbox("Show only rows with filled contact details (hide empty placeholders)", value=True)

df_view = st.session_state.fi_contacts_df.copy()

# Reorder columns cleanly
display_cols = ["FI Name", "Category", "Name", "Position", "Phone/Office Number", "Email Address"]
df_view = df_view[display_cols]

if hide_blank_contacts:
    df_view = df_view[(df_view["Name"] != "") | (df_view["Phone/Office Number"] != "") | (df_view["Email Address"] != "")]

if search_fi:
    df_view = df_view[df_view["FI Name"].str.contains(search_fi, case=False, na=False)]
if search_pos:
    df_view = df_view[df_view["Position"].str.contains(search_pos, case=False, na=False)]
if cat_filter != "All":
    df_view = df_view[df_view["Category"] == cat_filter]

st.caption(f"Showing **{len(df_view)}** contact record(s)")
st.dataframe(df_view, use_container_width=True, hide_index=True)

# ---------------------------------------------------------
# 5. DOWNLOAD EXCEL BUTTON
# ---------------------------------------------------------
st.download_button(
    label="📥 Download Contacts (XLSX)",
    data=to_excel(df_view),
    file_name="EthSwitch_FI_Contacts.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
