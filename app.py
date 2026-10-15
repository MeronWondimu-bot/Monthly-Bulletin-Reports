import streamlit as st

# 1. Page Configuration
st.set_page_config(
    page_title="EthSwitch Performance Reporting Hub",
    page_icon="🏢",
    layout="wide",
)

# 2. Custom Top Header with Branding Logos
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

# 3. Highlight Banner
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

# 4. Report Selection (Stacked Ordered List)
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
    horizontal=False,  # Vertically stacked list
    label_visibility="collapsed",  # Clean layout without repetitive label
)

st.markdown("---")

# 5. Content Routing & Script Placeholders
if selected_report == "Home / Overview":
  st.subheader("Welcome to the Central Dashboard")
  st.write(
      "Please select a report type from the list above to view detailed"
      " performance metrics and analytics."
  )

elif selected_report == "1. ATM decline summery":
  st.subheader("1. ATM Decline Summary")
  # TODO: Paste your script for Report 1 here

elif selected_report == "2. ATM Transaction Success Rate Summary":
  st.subheader("2. ATM Transaction Success Rate Summary")
  # TODO: Paste your script for Report 2 here

elif selected_report == "3. ATM success rate ranking":
  st.subheader("3. ATM Success Rate Ranking")
  # TODO: Paste your script for Report 3 here

elif selected_report == "4. ATM decline contribution by participant":
  st.subheader("4. ATM Decline Contribution by Participant")
  # TODO: Paste your script for Report 4 here

elif selected_report == "5. Major ATM declined transactions":
  st.subheader("5. Major ATM Declined Transactions")
  # TODO: Paste your script for Report 5 here
