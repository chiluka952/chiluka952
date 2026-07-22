import io
import pandas as pd
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="GSTR-1 Accounting Hub",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling for Modern Dark-Mode UI
st.markdown(
    """
    <style>
    .stApp {
        background-color: #0f172a;
        color: #f8fafc;
    }
    .main-title {
        font-size: 26px;
        font-weight: 800;
        color: #38bdf8;
        margin-bottom: 0px;
    }
    .sub-title {
        font-size: 14px;
        color: #94a3b8;
        margin-bottom: 20px;
    }
    div[data-testid="stMetricValue"] {
        font-family: 'Segoe UI', monospace;
        color: #38bdf8;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Header Section
st.markdown(
    '<div class="main-title">🛒 E-Commerce GST Statewise Entry Hub</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-title">Automated Accounting Engine for Meesho & Flipkart Sales and Returns</div>',
    unsafe_allow_html=True,
)

# Platform Selector
platform = st.radio(
    "Select Marketplace Platform:",
    ["Meesho", "Flipkart"],
    horizontal=True,
    help="Choose the platform whose reports you are processing.",
)

st.divider()

# File Uploaders
col1, col2 = st.columns(2)
with col1:
  sales_file = st.file_uploader(
      "1. Upload Raw TCS Sales Report (.xlsx / .csv)",
      type=["xlsx", "xls", "csv"],
  )
with col2:
  returns_file = st.file_uploader(
      "2. Upload Raw TCS Returns Report (.xlsx / .csv)",
      type=["xlsx", "xls", "csv"],
  )


# Fast File Reader
def load_data(file_obj):
  if file_obj is None:
    return None
  if file_obj.name.endswith(".csv"):
    return pd.read_csv(file_obj, low_memory=False)
  else:
    try:
      return pd.read_excel(file_obj, engine="calamine")
    except Exception:
      return pd.read_excel(file_obj, engine="openpyxl")


# Dynamic Column Finder
def find_col(df_cols, possible_names):
  cols_lower = {str(c).strip().lower(): c for c in df_cols}
  for name in possible_names:
    n_lower = name.lower()
    for col_l, original in cols_lower.items():
      if n_lower in col_l:
        return original
  return None


# Processing Logic
def generate_summary(df_s, df_r, plat):
  if df_s is None and df_r is None:
    return None

  ref_df = df_s if df_s is not None else df_r

  if plat == "Meesho":
    st_names = [
        "end_customer_state_new",
        "end_customer_state",
        "customer_state",
        "state",
    ]
    qty_names = ["quantity", "qty", "pcs"]
    tax_names = ["total_taxable_sale_value", "taxable_value", "taxable"]
  else:  # Flipkart
    st_names = [
        "Customer Delivery State",
        "Buyer State",
        "Delivery State",
        "State",
    ]
    qty_names = ["Quantity", "quantity", "qty"]
    tax_names = ["Taxable Value", "taxable_value", "Taxable Amount"]

  st_col = find_col(ref_df.columns, st_names)
  qty_col = find_col(ref_df.columns, qty_names)
  tax_col = find_col(ref_df.columns, tax_names)

  if not st_col or not qty_col or not tax_col:
    st.error(
        f"Could not map required columns. Found headers: {list(ref_df.columns[:8])}"
    )
    return None

  state_map = {}

  # Aggregate Sales
  if df_s is not None:
    s_st = find_col(df_s.columns, st_names) or st_col
    s_qty = find_col(df_s.columns, qty_names) or qty_col
    s_tax = find_col(df_s.columns, tax_names) or tax_col

    temp_s = df_s[[s_st, s_qty, s_tax]].dropna(subset=[s_st]).copy()
    temp_s[s_st] = temp_s[s_st].astype(str).str.strip().str.upper()
    temp_s = temp_s[temp_s[s_st] != "NAN"]
    temp_s[s_qty] = pd.to_numeric(temp_s[s_qty], errors="coerce").fillna(0)
    temp_s[s_tax] = pd.to_numeric(temp_s[s_tax], errors="coerce").fillna(0)

    grp_s = temp_s.groupby(s_st).agg({s_qty: "sum", s_tax: "sum"})
    for state, row in grp_s.iterrows():
      state_map[state] = {
          "pcs": row[s_qty],
          "tax": row[s_tax],
      }

  # Subtract Returns
  if df_r is not None:
    r_st = find_col(df_r.columns, st_names) or st_col
    r_qty = find_col(df_r.columns, qty_names) or qty_col
    r_tax = find_col(df_r.columns, tax_names) or tax_col

    temp_r = df_r[[r_st, r_qty, r_tax]].dropna(subset=[r_st]).copy()
    temp_r[r_st] = temp_r[r_st].astype(str).str.strip().str.upper()
    temp_r = temp_r[temp_r[r_st] != "NAN"]
    temp_r[r_qty] = pd.to_numeric(temp_r[r_qty], errors="coerce").fillna(0)
    temp_r[r_tax] = pd.to_numeric(temp_r[r_tax], errors="coerce").fillna(0)

    grp_r = temp_r.groupby(r_st).agg({r_qty: "sum", r_tax: "sum"})
    for state, row in grp_r.iterrows():
      if state not in state_map:
        state_map[state] = {"pcs": 0.0, "tax": 0.0}
      state_map[state]["pcs"] -= row[r_qty]
      state_map[state]["tax"] -= row[r_tax]

  rows = []
  for state in sorted(state_map.keys()):
    pcs = int(round(state_map[state]["pcs"]))
    tax = round(state_map[state]["tax"], 2)
    rate = round(tax / pcs, 2) if pcs > 0 else 0.0
    rows.append({"STATE": state, "PCS": pcs, "RATE": rate, "TAXABLE VALUE": tax})

  res_df = pd.DataFrame(rows)
  return res_df


# Action Button
if st.button("⚡ Generate Accounting Summary", type="primary"):
  if sales_file is None and returns_file is None:
    st.warning("Please upload at least one report file.")
  else:
    with st.spinner("Processing report entries..."):
      df_sales = load_data(sales_file)
      df_returns = load_data(returns_file)
      summary_df = generate_summary(df_sales, df_returns, platform)

      if summary_df is not None and not summary_df.empty:
        st.success("Summary Generated Successfully!")

        # Metrics Summary Row
        tot_pcs = summary_df["PCS"].sum()
        tot_tax = summary_df["TAXABLE VALUE"].sum()

        m1, m2 = st.columns(2)
        m1.metric("Total Pieces Sold (Net)", f"{tot_pcs:,} PCS")
        m2.metric("Total Taxable Value", f"₹{tot_tax:,.2f}")

        # Interactive Output Table
        st.dataframe(
            summary_df.style.format(
                {"PCS": "{:,}", "RATE": "₹{:,.2f}", "TAXABLE VALUE": "₹{:,.2f}"}
            ),
            use_container_width=True,
        )

        # Excel Export Button
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          summary_df.to_excel(
              writer, index=False, sheet_name="GST Statewise Summary"
          )
        excel_bytes = output.getvalue()

        st.download_button(
            label="📥 Download Clean Excel Format",
            data=excel_bytes,
            file_name=f"{platform}_GST_Statewise_Summary.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
