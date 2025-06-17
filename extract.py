#!/usr/bin/env python3
import sys, os, argparse, json
from openpyxl import load_workbook

# Try PDF support
try:
    from PyPDF2 import PdfReader
    _PDF_OK = True
except ImportError:
    _PDF_OK = False

def extract_excel(path):
    wb = load_workbook(path, data_only=True)
    out = {}
    for name in wb.sheetnames:
        ws = wb[name]
        out[name] = [list(row) for row in ws.iter_rows(values_only=True)]
    return out

def extract_pdf(path):
    if not _PDF_OK:
        sys.exit("Install PyPDF2 (`pip install PyPDF2`) to extract PDF")
    reader = PdfReader(path)
    all_text = "\n\n".join((p.extract_text() or "").strip() for p in reader.pages)
    instruction = (
        "Below is a financial summary. Extract the numbers for each year and output them as a JSON array. "
        "Each array element should have these keys, in this order: year, Revenue, Cost of Goods Sold (COGS), "
        "Operating Expenses, Taxes, Depreciation & Amortization, Plus Interest, Owner Salary+Super, "
        "Owner Benefits, Manager Salary, Investor Salary, One off Revenue Adjustments, One off Expenses Adjustments, "
        "Other Adjustments 1, Other Adjustments 2, Assets, Liabilities, Equity, Other Income, Total add backs. "
        "If a value is missing for a year, use 0. Do not add, remove, or infer categories. "
        "Output ONLY the JSON array—no explanation."
    )
    return {
        "instructions": instruction,
        "data": all_text
    }

def main():
    p = argparse.ArgumentParser(description="Extract data from PDF or Excel")
    p.add_argument("file", help="Path to .pdf or .xlsx file")
    args = p.parse_args()

    if not os.path.isfile(args.file):
        sys.exit(f"File not found: {args.file}")

    ext = os.path.splitext(args.file)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        data = extract_excel(args.file)
    elif ext == ".pdf":
        data = extract_pdf(args.file)
    else:
        sys.exit("Unsupported file type. Use .pdf or .xlsx")

    # Write the extracted data to extracted_data.json in the current working directory
    out_path = os.path.join(os.getcwd(), "extracted_data.json")
    with open(out_path, "w") as f:
        json.dump(data, f, indent=2, default=str)
    print(f"Extracted data saved to {out_path}")

    # Also print to console
    print(json.dumps(data, indent=2, default=str))

if __name__ == "__main__":
    main()