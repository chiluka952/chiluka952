import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="E-Commerce GST Entry Generator", layout="wide")

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


def process_data(df_s, df_r, plat):
  if df_s is None and df_r is None:
    return None

  if plat == "Meesho":
    st_col = "end_customer_state_new"
    qty_col = "quantity"
    tax_col = "total_taxable_sale_value"
  else:  # Flipkart
    target_df = df_s if df_s is not None else df_r
    st_col = (
        "Customer Delivery State"
        if "Customer Delivery State" in target_df.columns
        else "Buyer State"
    )
    qty_col = "Quantity"
    tax_col = "Taxable Value"

  sales_map = {}

  if df_s is not None:
    for _, row in df_s.iterrows():
      st_name = str(row.get(st_col, "")).strip().upper()
      if not st_name or st_name == "NAN":
        continue
      q = float(row.get(qty_col, 0) or 0)
      t = float(row.get(tax_col, 0) or 0)
      if st_name not in sales_map:
        sales_map[st_name] = {"pcs": 0, "tax": 0.0}
      sales_map[st_name]["pcs"] += q
      sales_map[st_name]["tax"] += t

  if df_r is not None:
    for _, row in df_r.iterrows():
      st_name = str(row.get(st_col, "")).strip().upper()
      if not st_name or st_name == "NAN":
        continue
      q = float(row.get(qty_col, 0) or 0)
      t = float(row.get(tax_col, 0) or 0)
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
      st.error(
          f"Error processing files: {str(e)}. Please check if file column"
          " headers match."
      )