import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Daily ATM Decline Monthly Aggregator", layout="wide"
)

st.title("📊 ATM Daily Decline Reports - Monthly Aggregator")
st.markdown(
    "Upload your **individual daily decline Excel files** below. The app will"
    " combine all daily reports and generate the complete **Monthly"
    " Summary Report** for each response code across all banks."
)

# 1. Multi-file uploader for individual daily reports
uploaded_files = st.sidebar.file_uploader(
    "Upload Daily Decline Reports (.xlsx)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
  st.success(f"📁 {len(uploaded_files)} daily report file(s) uploaded.")

  all_daily_dfs = []

  for file in uploaded_files:
    try:
      # Read the Excel file
      excel_data = pd.ExcelFile(file)

      # Process each sheet in the file (supports multi-sheet or single-sheet files)
      for sheet in excel_data.sheet_names:
        df_sheet = pd.read_excel(file, sheet_name=sheet)

        # Detect the row containing headers (where 'RC/BANK NAME' appears)
        header_row_idx = None
        for idx, row in df_sheet.iterrows():
          row_vals = [str(val).strip() for val in row.values if pd.notna(val)]
          if "RC/BANK NAME" in row_vals or "RC" in row_vals:
            header_row_idx = idx
            break

        if header_row_idx is not None:
          # Set correct headers and extract response code data
          df_sheet.columns = df_sheet.iloc[header_row_idx]
          data_df = df_sheet.iloc[header_row_idx + 1 :].copy()

          # Filter out empty rows or total summary rows
          data_df = data_df.dropna(subset=[data_df.columns[0]])
          data_df = data_df[
              ~data_df[data_df.columns[0]]
              .astype(str)
              .str.contains("Total|RC", case=False, na=False)
          ]

          all_daily_dfs.append(data_df)

    except Exception as e:
      st.error(f"Error reading file {file.name}: {e}")

  if all_daily_dfs:
    # Combine data from all uploaded daily files
    combined_raw = pd.concat(all_daily_dfs, ignore_index=True)

    # Standardize column names (Response Code column)
    rc_col = combined_raw.columns[0]
    combined_raw[rc_col] = (
        pd.to_numeric(combined_raw[rc_col], errors="coerce")
        .fillna(0)
        .astype(int)
        .astype(str)
    )

    # Filter out 0 or invalid response codes
    combined_raw = combined_raw[combined_raw[rc_col] != "0"]

    # Identify numeric bank columns
    bank_columns = [
        col
        for col in combined_raw.columns[1:]
        if pd.notna(col) and str(col).strip().upper() != "TOTAL"
    ]

    # Convert numeric values for calculations
    for col in bank_columns:
      combined_raw[col] = pd.to_numeric(
          combined_raw[col], errors="coerce"
      ).fillna(0)

    # 2. Group by Response Code (RC) and sum across all days
    monthly_summary = (
        combined_raw.groupby(rc_col)[bank_columns].sum().reset_index()
    )

    # Recalculate row totals across all banks
    monthly_summary["Total"] = monthly_summary[bank_columns].sum(axis=1)

    # Sort by total decline volume descending
    monthly_summary = monthly_summary.sort_values(
        by="Total", ascending=False
    ).reset_index(drop=True)

    # Add a Grand Total row at the bottom
    grand_total_row = {rc_col: "Total"}
    for col in bank_columns + ["Total"]:
      grand_total_row[col] = monthly_summary[col].sum()

    monthly_summary_with_total = pd.concat(
        [monthly_summary, pd.DataFrame([grand_total_row])], ignore_index=True
    )

    # Display results
    st.header("📈 Monthly Consolidated Decline Report")

    # Format numbers with commas for display
    display_df = monthly_summary_with_total.copy()
    for col in bank_columns + ["Total"]:
      display_df[col] = display_df[col].apply(lambda x: f"{int(x):,}")

    st.dataframe(display_df, use_container_width=True)

    # 3. Download option for the generated monthly report
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      monthly_summary_with_total.to_excel(
          writer, sheet_name="Monthly Decline Summary", index=False
      )

    st.download_button(
        label="📥 Download Monthly Consolidated Report (.xlsx)",
        data=output.getvalue(),
        file_name="Monthly_ATM_Decline_Consolidated_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
  st.info("👈 Please select and upload your daily report files from the sidebar.")
