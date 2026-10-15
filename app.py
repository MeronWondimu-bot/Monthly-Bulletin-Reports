import streamlit as st

# Set page config
st.set_page_config(
    page_title="EthSwitch Performance Reporting Hub",
    page_icon="💳",
    layout="wide",
)

# --- Header Section ---
st.markdown(
    "<h1 style='text-align: center;'>ETHSWITCH</h1>", unsafe_allow_html=True
)
st.markdown(
    "<h3 style='text-align: center; color: orange;'>EthSwitch Performance Reporting Hub</h3>",
    unsafe_allow_html=True,
)
st.markdown(
    "<p style='text-align: center; color: gray;'>Automated Bi-weekly & Monthly Report Generator</p>",
    unsafe_allow_html=True,
)
st.markdown("---")

# --- Report Type Selection (Replacing Report Category & Old Report Section) ---
st.markdown("### Report Type")

report_options = [
    "Home / Overview",
    "1. ATM decline summery",
    "2. ATM Transaction Success Rate Summary",
    "3. ATM success rate ranking",
    "4. ATM decline contribution by participant",
    "5. Major ATM declined transactions",
]

selected_report = st.radio("Select Report Type", report_options, horizontal=True)

st.markdown("---")

# --- Content Routing Based on Selection ---
if selected_report == "Home / Overview":
  st.subheader("Welcome to the Central Dashboard")
  st.write(
      "Please select a report type above to view detailed performance metrics"
      " and analytics."
  )
  # Add your Home/Overview content here

elif selected_report == "1. ATM decline summery":
  st.subheader("ATM Decline Summary")
  # TODO: Insert your ATM decline summary script/logic here

elif selected_report == "2. ATM Transaction Success Rate Summary":
  st.subheader("ATM Transaction Success Rate Summary")
  # TODO: Insert your success rate summary script/logic here

elif selected_report == "3. ATM success rate ranking":
  st.subheader("ATM Success Rate Ranking")
  # TODO: Insert your ranking script/logic here (e.g., benchmark vs 98.5% target)

elif selected_report == "4. ATM decline contribution by participant":
  st.subheader("ATM Decline Contribution by Participant")
  # TODO: Insert your participant contribution script/logic here

elif selected_report == "5. Major ATM declined transactions":
  st.subheader("Major ATM Declined Transactions")
  # TODO: Insert your major declines analysis script/logic here
