import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="E-Commerce GST Entry Generator", layout="wide"
)

st.title("🛒 E-Commerce GST Statewise Summary Generator")
st.write(
    "Upload your raw Meesho or Flipkart Sales & Return reports to generate"
    " clean statewise accounting entries."
)

platform = st.radio(
    "Select Marketplace Platform:", ["Meesho", "Flipkart"], horizontal=True
)

col1, col2 = st.columns(2)
with col1:
  sales_file = st.file_uploader(
      "Upload Raw Sales Report (.xlsx / .csv)", type=["xlsx", "xls", "csv"]
  )
with col2:
  returns_file = st.file_uploader(
      "Upload Raw Returns Report (.xlsx / .csv)", type=["xlsx", "xls", "csv"]
  )


def load_file(uploaded_file):
  if uploaded_file is None:
    return None
  if uploaded_file.name.endswith(".csv"):
    return pd.read_csv(uploaded_file)
  else:
    return pd.read_excel(uploaded_file)


def find_col(df, possible_names):
  if df is None:
    return None
  cols = [str(c).strip().lower() for c in df.columns]
  for target in possible_names:
    target_lower = target.lower()
    for idx, col in enumerate(cols):
      if target_lower in col:
        return df.columns[idx]
  return None


def process_data(df_s, df_r, plat):
  if df_s is None and df_r is None:
    return None

  ref_df = df_s if df_s is not None else df_r

  if plat == "Meesho":
    st_col = find_col(
        ref_df,
        [
            "end_customer_state_new",
            "end_customer_state",
            "customer_state",
            "state",
            "delivery_state",
        ],
    )
    qty_col = find_col(ref_df, ["quantity", "qty", "pcs", "pcs_count"])
    tax_col = find_col(
        ref_df,
        [
            "total_taxable_sale_value",
            "taxable_value",
            "taxable_amount",
            "taxable",
            "total_taxable_value",
        ],
    )
  else:  # Flipkart
    st_col = find_col(
        ref_df,
        [
            "Customer Delivery State",
            "Buyer State",
            "Delivery State",
            "State",
            "state",
        ],
    )
    qty_col = find_col(ref_df, ["Quantity", "quantity", "qty"])
    tax_col = find_col(
        ref_df,
        [
            "Taxable Value",
            "taxable_value",
            "Taxable Amount",
            "Total Taxable Value",
        ],
    )

  if not st_col or not qty_col or not tax_col:
    available_cols = ", ".join([str(c) for c in ref_df.columns[:10]])
    raise ValueError(
        f"Could not automatically detect required columns. Available columns"
        f" include: {available_cols}"
    )

  sales_map = {}

  if df_s is not None:
    s_st = find_col(df_s, [st_col]) or st_col
    s_qty = find_col(df_s, [qty_col]) or qty_col
    s_tax = find_col(df_s, [tax_col]) or tax_col
    for _, row in df_s.iterrows():
      st_name = str(row.get(s_st, "")).strip().upper()
      if not st_name or st_name == "NAN":
        continue
      try:
        q = float(row.get(s_qty, 0) or 0)
        t = float(row.get(s_tax, 0) or 0)
      except Exception:
        continue
      if st_name not in sales_map:
        sales_map[st_name] = {"pcs": 0, "tax": 0.0}
      sales_map[st_name]["pcs"] += q
      sales_map[st_name]["tax"] += t

  if df_r is not None:
    r_st = find_col(df_r, [st_col]) or st_col
    r_qty = find_col(df_r, [qty_col]) or qty_col
    r_tax = find_col(df_r, [tax_col]) or tax_col
    for _, row in df_r.iterrows():
      st_name = str(row.get(r_st, "")).strip().upper()
      if not st_name or st_name == "NAN":
        continue
      try:
        q = float(row.get(r_qty, 0) or 0)
        t = float(row.get(r_tax, 0) or 0)
      except Exception:
        continue
      if st_name not in sales_map:
        sales_map[st_name] = {"pcs": 0, "tax": 0.0}
      sales_map[st_name]["pcs"] -= q
      sales_map[st_name]["tax"] -= t

  rows = []
  for st_name in sorted(sales_map.keys()):
    pcs = int(sales_map[st_name]["pcs"])
    tax = round(sales_map[st_name]["tax"], 2)
    rate = round(tax / pcs, 2) if pcs > 0 else 0.0
    rows.append(
        {"STATE": st_name, "PCS": pcs, "RATE": rate, "TAXABLE VALUE": tax}
    )

  return pd.DataFrame(rows)


if st.button("⚡ Generate Accounting Summary", type="primary"):
  df_sales = load_file(sales_file)
  df_returns = load_file(returns_file)

  if df_sales is None and df_returns is None:
    st.error("Please upload at least one report file.")
  else:
    try:
      final_df = process_data(df_sales, df_returns, platform)
      st.success("Summary Generated Successfully!")
      st.dataframe(final_df, use_container_width=True)

      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="GST State Summary")
      excel_data = output.getvalue()

      st.download_button(
          label="📥 Download Clean Excel Format",
          data=excel_data,
          file_name=f"{platform}_GST_Statewise_Summary.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    except Exception as e:
      st.error(f"Processing details: {str(e)}")
