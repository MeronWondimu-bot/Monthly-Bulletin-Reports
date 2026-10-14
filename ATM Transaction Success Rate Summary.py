import pandas as pd

# Load the previously uploaded success rate file
file_path = "ATM Success Rate Sep 01-30,2026.xlsx"
df_success = pd.read_excel(file_path, sheet_name="Sheet 1")

# Extract bank names (Row index 3) and monthly averages (Row index 34)
headers = df_success.iloc[3].tolist()
averages = df_success.iloc[34].tolist()

# Build summary table data
summary_data = []
for bank, avg_val in zip(headers[1:], averages[1:]):
  if pd.notna(bank):
    success_rate = float(avg_val)
    status = "Meets target" if success_rate >= 0.985 else "Below target"
    summary_data.append({
        "Bank / Participant": bank,
        "Monthly Average Success Rate": f"{success_rate * 100:.2f}%",
        "Target Status (>= 98.5%)": status,
    })

# Convert to DataFrame and sort by success rate descending
summary_df = pd.DataFrame(summary_data)

# Separate weighted average row from the banks
weighted_avg_row = summary_df[
    summary_df["Bank / Participant"] == "Weighted Average"
]
banks_df = summary_df[
    summary_df["Bank / Participant"] != "Weighted Average"
].copy()

# Sort banks by success rate
banks_df["sort_val"] = (
    df_success.iloc[34, 1:-1].astype(float).values
    if len(df_success.iloc[34, 1:-1]) == len(banks_df)
    else 0
)
banks_df = banks_df.sort_values(
    by="Monthly Average Success Rate", ascending=False
).reset_index(drop=True)
banks_df = banks_df.drop(columns=["sort_val"])

# Add rank
banks_df.index += 1
banks_df.insert(0, "Rank", banks_df.index)

# Combine back with weighted average
print("=== ATM TRANSACTION SUCCESS RATE SUMMARY ===")
print(banks_df.to_string())
print("\nOverall Monthly Weighted Average:")
print(weighted_avg_row.to_string(index=False))

# Optional: Export to a new Excel summary file
with pd.ExcelWriter(
    "ATM_Success_Rate_Summary_Report.xlsx", engine="openpyxl"
) as writer:
  banks_df.to_excel(writer, sheet_name="Success Rate Ranking", index=False)
  weighted_avg_row.to_excel(
      writer, sheet_name="Weighted Average", index=False
  )