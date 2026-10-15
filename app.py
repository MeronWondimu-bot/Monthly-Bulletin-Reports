import pandas as pd
import streamlit as st

# 1. Page Configuration
st.set_page_config(
    page_title="EthSwitch Performance Reporting Hub",
    page_icon="🏢",
    layout="wide",
)

# 2. Top Header & Branding
col1, col2, col3 = st.columns([1, 3, 1])
with col1:
  st.markdown("## 🏢 **ETHSWITCH**")
with col2:
  st.markdown(
      "<h1 style='text-align: center; color: #E65100; font-weight: bold;"
      " margin-bottom: 0;'>EthSwitch Performance Reporting Hub</h1>",
      unsafe_allow_html=True,
  )
  st.markdown(
      "<p style='text-align: center; color: #666666; margin-top: 0;'>Automated"
      " Bi-weekly & Monthly Report Generator</p>",
      unsafe_allow_html=True,
  )
with col3:
  st.markdown(
      "<h2 style='text-align: right;'>🦜 <span"
      " style='color:#C2185B;'>EthioPay</span></h2>",
      unsafe_allow_html=True,
  )

st.markdown("<br>", unsafe_allow_html=True)

# 3. Banner
st.markdown(
    """
    <div style="background-color: #FFF3E0; border-left: 6px solid #E65100; padding: 12px; border-radius: 4px; text-align: center;">
        <h3 style="color: #E65100; margin: 0; font-size: 18px; font-weight: bold; letter-spacing: 1px;">
            BI-WEEKLY & MONTHLY REPORT GENERATOR
        </h3>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("<br>", unsafe_allow_html=True)


# 4. Data Loading Function (Cached for fast switching)
@st.cache_data
def load_data(uploaded_file):
  """Reads Excel/CSV data once for all 5 reports."""
  if uploaded_file.name.endswith(".csv"):
    df = pd.read_csv(uploaded_file)
  else:
    df = pd.read_excel(uploaded_file)
  # Place any global data cleaning here (e.g. standardizing column names)
  return df


# 5. Data Upload Section
st.markdown("### 📁 **Upload Source Data**")
uploaded_file = st.file_uploader(
    "Upload your ATM Transaction Excel or CSV file (used across all 5 reports)",
    type=["xlsx", "xls", "csv"],
)

st.markdown("---")

# 6. Report Selection Options
st.markdown("### **Report Type**")

report_options = [
    "Home / Overview",
    "1. ATM decline summery",
    "2. ATM Transaction Success Rate Summary",
    "3. ATM success rate ranking",
    "4. ATM decline contribution by participant",
    "5. Major ATM declined transactions",
]

selected_report = st.radio(
    "Select Report Type",
    report_options,
    horizontal=False,
    label_visibility="collapsed",
)

st.markdown("---")

# 7. Report Generation Logic
if uploaded_file is None:
  st.info(
      "👈 Please upload your transaction data file above to begin generating"
      " reports."
  )

else:
  # Load the data once
  df = load_data(uploaded_file)
  st.success(
      f"Data loaded successfully! Total records: **{len(df):,} rows**"
  )

  # --- REPORT 0: HOME / OVERVIEW ---
  if selected_report == "Home / Overview":
    st.subheader("📊 Dataset Overview")
    st.write("Data Preview:")
    st.dataframe(df.head(10), use_container_width=True)

  # --- REPORT 1: ATM DECLINE SUMMARY ---
  elif selected_report == "1. ATM decline summery":
    st.subheader("1. ATM Decline Summary")
    # Example logic: Filter by declined transactions
    # df_declined = df[df['Status'] == 'Decline']
    # st.dataframe(df_declined)

  # --- REPORT 2: SUCCESS RATE SUMMARY ---
  elif selected_report == "2. ATM Transaction Success Rate Summary":
    st.subheader("2. ATM Transaction Success Rate Summary")
    # Example logic: Total Success vs Total Declined

  # --- REPORT 3: SUCCESS RATE RANKING ---
  elif selected_report == "3. ATM success rate ranking":
    st.subheader("3. ATM Success Rate Ranking")
    # Example logic: Group by Bank/Participant, compute success %, compare against 98.5% target threshold

  # --- REPORT 4: DECLINE CONTRIBUTION BY PARTICIPANT ---
  elif selected_report == "4. ATM decline contribution by participant":
    st.subheader("4. ATM Decline Contribution by Participant")
    # Example logic: Group by Participant bank to see contribution % to overall declines

  # --- REPORT 5: MAJOR DECLINED TRANSACTIONS ---
  elif selected_report == "5. Major ATM declined transactions":
    st.subheader("5. Major ATM Declined Transactions")
    # Example logic: Group by Response Code / Reason to find top decline causes
