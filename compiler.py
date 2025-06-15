# compiler.py

import openpyxl

def map_to_excel(json_data, excel_path, output_path=None):
    """
    Updates the existing Excel template (preserving formulas, formatting).
    If output_path is None, will overwrite excel_path in place.
    """
    # Mapping from JSON keys to Excel template category labels
    category_mapping = {
        "Revenue": "Reveneue",  # typo in Excel
        "Cost of Goods Sold (COGS)": "Cost of Goods Sold (COGS)",
        "Operating Expenses": "Less Operating Expenses",  # ← correct!
        "Depreciation & Amortization": "Plus Depreciation & Amortization",
        "Plus Interest": "Plus Interest",
        "Owner Salary+Super": "Plus Owner Salary+Super etc",
        "Owner Benefits": "Plus Owner Benefits",
        "Manager Salary": "Manager Salary",
        "Investor Salary": "Investor Salary",
        "One off Revenue Adjustments": "One off Revenue Adjustments",
        "One off Expenses Adjustments": "One off Expenses Adjustments",
        "Taxes": "Taxes",
        # These are missing in your template, so will be skipped:
        # "Assets": "",
        # "Liabilities": "",
        # "Equity": "",
        # "Other Income": "",
    }

    if isinstance(json_data, dict) and "year" in json_data:
        data = [json_data]
    elif isinstance(json_data, list):
        data = json_data
    else:
        raise ValueError("Invalid JSON data format.")

    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    year_row_idx = 3  # Third row: years
    cat_col_idx = 2   # Column B: categories

    # Map years to column numbers (accept both number and string keys)
    years = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=year_row_idx, column=col).value
        if val is not None:
            key_str = str(val).strip()
            years[key_str] = col
            try:
                key_int = str(int(float(val)))
                years[key_int] = col
            except Exception:
                pass

    # Map categories to row numbers (accept both number and string keys)
    cats = {}
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row=row, column=cat_col_idx).value
        if val is not None:
            key_str = str(val).strip()
            cats[key_str] = row

    print("Year columns found:", years)
    print("Categories found in template:", list(cats.keys()))

    for entry in data:
        yr = str(entry["year"]).strip()
        if yr not in years:
            print(f"Year {yr} not found, skipping.")
            continue
        for category, value in entry.items():
            if category == "year":
                continue
            excel_label = category_mapping.get(category, category)
            if excel_label in cats:
                ws.cell(row=cats[excel_label], column=years[yr]).value = value
            else:
                print(f"Category '{category}' (Excel: '{excel_label}') not found, skipping.")

    save_path = output_path if output_path else excel_path
    wb.save(save_path)
    print(f"Saved to {save_path}")

# Optional CLI usage
if __name__ == "__main__":
    import argparse, json
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", required=True, help="Extracted JSON file from GPT")
    parser.add_argument("--template", required=True, help="Excel template file")
    parser.add_argument("--output", help="Output Excel file (leave blank to overwrite)")
    args = parser.parse_args()

    with open(args.json) as f:
        json_data = json.load(f)
    map_to_excel(json_data, args.template, args.output)