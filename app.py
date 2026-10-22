# 5. Success Rate Processing Function (Report 2) - Enhanced Fallback for Oct 15
@st.cache_data
def process_success_rate_matrix(file_tuples_names):
  daily_success_records = []

  for idx, file in enumerate(file_tuples_names, start=1):
    filename = getattr(file, "name", f"file_{idx}.xlsx")
    date_label = extract_date_label(filename, idx)

    df_raw = pd.read_excel(file, header=None)
    success_row = None
    weighted_avg_from_file = None

    # 1. Broad row scan for success rate keywords
    for r_idx, row in df_raw.iterrows():
      row_str = " ".join([str(val) for val in row.values]).upper()
      if ("ISSU" in row_str and "SUCC" in row_str) or ("SUCCESS" in row_str and "RATE" in row_str):
        success_row = row
        break

    # 2. Fallback: If specific text isn't matched, check the last few rows of the file where summary rows usually sit
    if success_row is None and len(df_raw) > 5:
      for r_idx in range(len(df_raw) - 1, max(0, len(df_raw) - 8), -1):
        row = df_raw.iloc[r_idx]
        row_str = " ".join([str(val) for val in row.values]).upper()
        if any(term in row_str for term in ["RATE", "%", "TOTAL", "AVERAGE"]):
          success_row = row
          break

    bank_success_rates = {"Bank/Date": date_label}

    if success_row is not None:
      header_row_idx = None
      for r_idx, row in df_raw.iterrows():
        if row.astype(str).str.contains("RC/BANK NAME", case=False, na=False).any():
          header_row_idx = r_idx
          break

      if header_row_idx is not None:
        header_vals = df_raw.iloc[header_row_idx].values
        header_cols = [standardize_bank_name(c) for c in header_vals]
        
        # Look for Weighted Average / Total column in header
        for col_idx, col_name in enumerate(header_vals):
          c_str = str(col_name).upper()
          if any(term in c_str for term in ["WEIGHTED", "W.AVG", "TOTAL", "AVERAGE"]):
            w_val = success_row.iloc[col_idx]
            try:
              w_str = str(w_val).replace("%", "").strip()
              val_num = float(w_str)
              if val_num <= 1.0 and val_num > 0:
                val_num *= 100
              weighted_avg_from_file = round(val_num, 1)
              break
            except:
              pass

        # Trailing cells scan for summary/weighted average figure
        if weighted_avg_from_file is None:
          for col_idx in range(len(success_row) - 1, max(0, len(success_row) - 6), -1):
            val = success_row.iloc[col_idx]
            try:
              val_str = str(val).replace("%", "").strip()
              num = float(val_str)
              if 0 < num <= 100 or 0 < num <= 1.0:
                if num <= 1.0:
                  num *= 100
                weighted_avg_from_file = round(num, 1)
                break
            except:
              pass

        for bank in STANDARD_BANKS:
          if bank in header_cols:
            col_pos = header_cols.index(bank)
            val = success_row.iloc[col_pos]
            try:
              if pd.isna(val):
                bank_success_rates[bank] = 100.0
              else:
                val_str = str(val).replace("%", "").strip()
                if val_str.lower() == "nan" or val_str == "":
                  bank_success_rates[bank] = 100.0
                else:
                  num = float(val_str)
                  if num <= 1.0 and num > 0:
                    num = num * 100
                  bank_success_rates[bank] = round(num, 1)
            except ValueError:
              bank_success_rates[bank] = 100.0
          else:
            bank_success_rates[bank] = 100.0
    else:
      for bank in STANDARD_BANKS:
        bank_success_rates[bank] = 100.0

    if weighted_avg_from_file is not None:
      bank_success_rates["Weighted Average"] = weighted_avg_from_file
    else:
      valid_rates = [v for k, v in bank_success_rates.items() if k in STANDARD_BANKS]
      bank_success_rates["Weighted Average"] = round(sum(valid_rates) / len(valid_rates), 1) if valid_rates else 100.0

    daily_success_records.append(bank_success_rates)

  if not daily_success_records:
    return None

  success_df = pd.DataFrame(daily_success_records)
  cols = ["Bank/Date"] + STANDARD_BANKS + ["Weighted Average"]
  success_df = success_df[cols]

  avg_row = {"Bank/Date": "Average"}
  for bank in STANDARD_BANKS + ["Weighted Average"]:
    col_vals = pd.to_numeric(success_df[bank], errors="coerce")
    avg_row[bank] = round(col_vals.mean(), 1)

  success_df = pd.concat([success_df, pd.DataFrame([avg_row])], ignore_index=True)
  return success_df
