import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd


def process_meesho_gst_files(
    sales_filepath, returns_filepath, output_filepath
):
  print("Processing files and cleaning numeric errors...")

  # 1. Read Sales File
  df_sales = (
      pd.read_excel(sales_filepath)
      if sales_filepath.endswith((".xlsx", ".xls"))
      else pd.read_csv(sales_filepath)
  )

  # 2. Read Returns File
  df_returns = None
  if returns_filepath:
    df_returns = (
        pd.read_excel(returns_filepath)
        if returns_filepath.endswith((".xlsx", ".xls"))
        else pd.read_csv(returns_filepath)
    )

  # Explicit Column Mappings specified by you:
  # State -> end_customer_state_new
  # PCS -> quantity
  # Taxable Value -> total_taxable_sale_value

  s_st_col = "end_customer_state_new"
  s_qty_col = "quantity"
  s_tax_col = "total_taxable_sale_value"

  state_map = {}

  # Clean & Aggregate Sales Data
  if s_st_col in df_sales.columns:
    # Fix "Number Stored as Text" errors by forcing numeric conversion
    df_sales[s_qty_col] = pd.to_numeric(
        df_sales[s_qty_col], errors="coerce"
    ).fillna(0)
    df_sales[s_tax_col] = pd.to_numeric(
        df_sales[s_tax_col], errors="coerce"
    ).fillna(0)
    df_sales[s_st_col] = (
        df_sales[s_st_col].astype(str).str.strip().str.upper()
    )

    for _, row in df_sales.iterrows():
      st_name = row[s_st_col]
      if not st_name or st_name == "NAN":
        continue
      if st_name not in state_map:
        state_map[st_name] = {"pcs": 0.0, "tax": 0.0}
      state_map[st_name]["pcs"] += row[s_qty_col]
      state_map[st_name]["tax"] += row[s_tax_col]

  # Clean & Subtract Returns Data
  if df_returns is not None and s_st_col in df_returns.columns:
    df_returns[s_qty_col] = pd.to_numeric(
        df_returns[s_qty_col], errors="coerce"
    ).fillna(0)
    df_returns[s_tax_col] = pd.to_numeric(
        df_returns[s_tax_col], errors="coerce"
    ).fillna(0)
    df_returns[s_st_col] = (
        df_returns[s_st_col].astype(str).str.strip().str.upper()
    )

    for _, row in df_returns.iterrows():
      st_name = row[s_st_col]
      if not st_name or st_name == "NAN":
        continue
      if st_name not in state_map:
        state_map[st_name] = {"pcs": 0.0, "tax": 0.0}
      state_map[st_name]["pcs"] -= row[s_qty_col]
      state_map[st_name]["tax"] -= row[s_tax_col]

  # Build Statewise Summary Rows
  summary_rows = []
  for st_name in sorted(state_map.keys()):
    pcs = int(round(state_map[st_name]["pcs"]))
    tax = round(state_map[st_name]["tax"], 2)
    rate = round(tax / pcs, 2) if pcs > 0 else 0.0
    rows = {"STATE": st_name, "PCS": pcs, "RATE": rate, "TAXABLE VALUE": tax}
    summary_rows.append(rows)

  # Generate Formatted Excel Workbook without errors
  wb = openpyxl.Workbook()
  ws = wb.active
  ws.title = "GSTR1 Statewise Summary"
  ws.views.sheetView[0].showGridLines = True

  # Styling & Formats
  fill_navy = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  fill_zebra = PatternFill(
      start_color="F9FAFB", end_color="F9FAFB", fill_type="solid"
  )
  fill_total = PatternFill(
      start_color="E9EDF4", end_color="E9EDF4", fill_type="solid"
  )

  font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
  font_data = Font(name="Segoe UI", size=10)
  font_total = Font(name="Segoe UI", size=10, bold=True)

  border_side = Side(style="thin", color="E0E0E0")
  border_data = Border(
      left=border_side, right=border_side, top=border_side, bottom=border_side
  )
  border_total = Border(
      top=Side(style="thin", color="000000"),
      bottom=Side(style="double", color="000000"),
  )

  # Indian Number Formatting rule (₹15,74,455.94)
  indian_fmt = (
      "[>=10000000]₹##\\,##\\,##\\,##0.00;[>=100000]₹##\\,##\\,##0.00;₹##,##0.00"
  )

  # Headers
  headers = ["STATE", "PCS", "RATE", "TAXABLE VALUE"]
  for col_idx, h_text in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=col_idx, value=h_text)
    cell.font = font_header
    cell.fill = fill_navy
    cell.alignment = Alignment(horizontal="center", vertical="center")
  ws.row_dimensions[1].height = 24

  # Data Rows
  for idx, r_data in enumerate(summary_rows, start=2):
    is_even = idx % 2 == 0

    ws.cell(row=idx, column=1, value=r_data["STATE"]).alignment = Alignment(
        horizontal="left", vertical="center"
    )

    c_pcs = ws.cell(row=idx, column=2, value=r_data["PCS"])
    c_pcs.number_format = "#,##0"
    c_pcs.alignment = Alignment(horizontal="right", vertical="center")

    c_rate = ws.cell(row=idx, column=3, value=r_data["RATE"])
    c_rate.number_format = indian_fmt
    c_rate.alignment = Alignment(horizontal="right", vertical="center")

    c_tax = ws.cell(row=idx, column=4, value=r_data["TAXABLE VALUE"])
    c_tax.number_format = indian_fmt
    c_tax.alignment = Alignment(horizontal="right", vertical="center")

    ws.row_dimensions[idx].height = 20
    for col_c in range(1, 5):
      cell_obj = ws.cell(row=idx, column=col_c)
      cell_obj.font = font_data
      cell_obj.border = border_data
      if is_even:
        cell_obj.fill = fill_zebra

  # Total Row
  tot_row = len(summary_rows) + 2
  ws.row_dimensions[tot_row].height = 22
  ws.cell(row=tot_row, column=1, value="Total").alignment = Alignment(
      horizontal="left", vertical="center"
  )

  c_tot_pcs = ws.cell(
      row=tot_row, column=2, value=f"=SUM(B2:B{tot_row-1})"
  )
  c_tot_pcs.number_format = "#,##0"
  c_tot_pcs.alignment = Alignment(horizontal="right", vertical="center")

  c_tot_tax = ws.cell(
      row=tot_row, column=4, value=f"=SUM(D2:D{tot_row-1})"
  )
  c_tot_tax.number_format = indian_fmt
  c_tot_tax.alignment = Alignment(horizontal="right", vertical="center")

  for col_c in range(1, 5):
    cell_obj = ws.cell(row=tot_row, column=col_c)
    cell_obj.font = font_total
    cell_obj.fill = fill_total
    cell_obj.border = border_total

  ws.column_dimensions["A"].width = 38
  ws.column_dimensions["B"].width = 16
  ws.column_dimensions["C"].width = 18
  ws.column_dimensions["D"].width = 22

  wb.save(output_filepath)
  print(f"File successfully created: {output_filepath}")


# Run the function with your file paths:
process_meesho_gst_files(
    "tcs_sales.xlsx", "tcs_sales_return.xlsx", "Meesho_Clean_GSTR1_Report.xlsx"
)

 
