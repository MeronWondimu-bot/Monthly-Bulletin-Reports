# ==========================================
# 2. ATM TRANSACTION SUCCESS RATE SUMMARY
# ==========================================
elif report_type == "2. ATM Transaction Success Rate Summary":
  st.subheader("📈 ATM Transaction Success Rate Summary")
  if uploaded_files:
    success_rows = []
    for file in uploaded_files:
      try:
        xls = pd.ExcelFile(file)
        raw_name = file.name.split(".")[0]
        cleaned_name = (
            raw_name.replace("ATM Declined Transaction Report for", "")
            .replace("ATM Success Rate", "")
            .strip()
        )
        file_date = re.sub(r",\s*", ", ", cleaned_name)

        for sheet_name in xls.sheet_names:
          df_sheet = pd.read_excel(file, sheet_name=sheet_name)
          target_row_series = None
          for idx, row in df_sheet.iterrows():
            row_str = " ".join([str(v) for v in row.values if pd.notna(v)]).upper()
            if "SUCC" in row_str and "RATE" in row_str:
              target_row_series = row
              break

          if target_row_series is not None:
            header_idx = None
            for h_idx in range(len(df_sheet)):
              h_str = " ".join([str(v) for v in df_sheet.iloc[h_idx].values if pd.notna(v)]).upper()
              if "CBE" in h_str or "BOA" in h_str:
                header_idx = h_idx
                break

            cols = (
                df_sheet.iloc[header_idx].tolist()
                if header_idx is not None
                else df_sheet.columns.tolist()
            )
            vals = target_row_series.values.tolist()

            record = {"Bank/Date": file_date}
            for c, val in zip(cols, vals):
              c_str = str(c).strip().upper().replace("\n", " ")
              if c_str in HEADER_MAPPING:
                b_name = HEADER_MAPPING[c_str]
                try:
                  num = float(str(val).replace("%", "").strip())
                  record[b_name] = num
                except:
                  record[b_name] = None
              elif "TSEDAY" in c_str:
                try:
                  num = float(str(val).replace("%", "").strip())
                  record["Tseday"] = num
                except:
                  record["Tseday"] = None
              elif "WEIGHT" in c_str or "AVERAGE" in c_str:
                try:
                  num = float(str(val).replace("%", "").strip())
                  record["Weighted Average"] = num
                except:
                  record["Weighted Average"] = None

            if len(record) > 1:
              success_rows.append(record)
              break
      except Exception as e:
        st.error(f"Error reading {file.name}: {e}")

    if success_rows:
      df_success = pd.DataFrame(success_rows)

      # Chronological sorting
      try:
        norm_dates = df_success["Bank/Date"].str.replace(r",\s*", ", ", regex=True)
        df_success["_dt"] = pd.to_datetime(norm_dates, format="mixed", errors="coerce")
        df_success = df_success.sort_values(by="_dt").drop(columns=["_dt"]).reset_index(drop=True)
      except Exception:
        pass

      # Column ordering: Bank/Date first, standard banks, and Weighted Average last
      bank_cols = [c for c in df_success.columns if c != "Bank/Date" and c != "Weighted Average"]
      ordered_cols = ["Bank/Date"] + bank_cols
      if "Weighted Average" in df_success.columns:
        ordered_cols.append("Weighted Average")
      df_success = df_success[[c for c in ordered_cols if c in df_success.columns]]

      # Average row at bottom
      avg_dict = {"Bank/Date": "Average"}
      calc_cols = [c for c in df_success.columns if c != "Bank/Date"]
      for col in calc_cols:
        numeric_vals = pd.to_numeric(df_success[col], errors="coerce")
        avg_dict[col] = numeric_vals.mean()

      df_success = pd.concat([df_success, pd.DataFrame([avg_dict])], ignore_index=True)

      display_df = df_success.astype(object)

      for idx, row in df_success.iterrows():
        is_avg = str(row["Bank/Date"]).strip().lower() == "average"
        for col in calc_cols:
          val = row[col]
          try:
            num = float(val)
            if pd.isna(num):
              display_df.at[idx, col] = ""
              continue

            pct_val = num * 100 if num <= 1.0 else num
            pct_str = f"{pct_val:.1f}%"

            if is_avg:
              ratio = num if num <= 1.0 else num / 100.0
              if ratio >= 0.985:
                display_df.at[idx, col] = f"<span style='background-color: #2ecc71; color: white; font-weight: bold; padding: 2px 4px; border-radius: 3px;'>{pct_str}</span>"
              elif ratio >= 0.90:
                display_df.at[idx, col] = f"<span style='background-color: #f1c40f; color: black; font-weight: bold; padding: 2px 4px; border-radius: 3px;'>{pct_str}</span>"
              elif ratio >= 0.80:
                display_df.at[idx, col] = f"<span style='background-color: #fcf3cf; color: black; font-weight: bold; padding: 2px 4px; border-radius: 3px;'>{pct_str}</span>"
              else:
                display_df.at[idx, col] = f"<span style='background-color: #e74c3c; color: white; font-weight: bold; padding: 2px 4px; border-radius: 3px;'>{pct_str}</span>"
            else:
              display_df.at[idx, col] = pct_str
          except:
            pass

      st.markdown(display_df.to_html(escape=False, index=False), unsafe_allow_html=True)

      out_succ = io.BytesIO()
      with pd.ExcelWriter(out_succ, engine="openpyxl") as writer:
        df_success.to_excel(writer, sheet_name="Success Rate Summary", index=False)
      st.download_button(
          "📥 Download Success Rate Summary (.xlsx)",
          out_succ.getvalue(),
          "ATM_Success_Rate_Summary.xlsx",
          key="dl_success_sum",
      )
    else:
      st.warning("Could not locate 'SUCC. RATE' rows in the uploaded files. Please check file structure.")
  else:
    st.info("👆 Please upload your daily source report files at the top.")
