# compiler.py

import openpyxl

def map_to_excel(json_data, excel_path, output_path=None):
    """
    Updates the existing Excel template (preserving formulas, formatting).
    If output_path is None, will overwrite excel_path in place.
    """
    if isinstance(json_data, dict) and "year" in json_data:
        data = [json_data]
    elif isinstance(json_data, list):
        data = json_data
    else:
        raise ValueError("Invalid JSON data format.")

    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    # Find where year columns and category rows are
    year_row_idx = 1  # First row: years
    cat_col_idx = 1   # First column: categories

    # Map years to column numbers
    years = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=year_row_idx, column=col).value
        if val is not None:
            years[str(val).strip()] = col

    # Map categories to row numbers
    cats = {}
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row=row, column=cat_col_idx).value
        if val is not None:
            cats[str(val).strip()] = row

    for entry in data:
        yr = str(entry["year"])
        if yr not in years:
            print(f"Year {yr} not found, skipping.")
            continue
        for category, value in entry.items():
            if category == "year":
                continue
            if category in cats:
                ws.cell(row=cats[category], column=years[yr]).value = value
            else:
                print(f"Category '{category}' not found, skipping.")

    # Save (overwrite if output_path is None)
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