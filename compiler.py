# compiler.py

import pandas as pd

def map_to_excel(json_data, excel_path, output_path):
    """
    json_data: dict (single year) or list of dicts (multiple years)
    excel_path: path to input template
    output_path: path to save filled Excel file
    """
    if isinstance(json_data, dict) and "year" in json_data:
        data = [json_data]
    elif isinstance(json_data, list):
        data = json_data
    else:
        raise ValueError("Invalid JSON data format.")

    df = pd.read_excel(excel_path, engine='openpyxl')
    header = df.columns.tolist()
    label_col = df.columns[0]

    for entry in data:
        yr = entry["year"]
        if str(yr) not in header:
            print(f"Year {yr} not found in columns, skipping.")
            continue
        for category, value in entry.items():
            if category == "year":
                continue
            match = df[df[label_col].astype(str).str.strip() == category]
            if not match.empty:
                row_idx = match.index[0]
                df.at[row_idx, str(yr)] = value
            else:
                print(f"Category '{category}' not found, skipping.")

    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"Saved to {output_path}")

# Optional: CLI for testing (but not needed for GUI)
if __name__ == "__main__":
    import argparse, json
    parser = argparse.ArgumentParser()
    parser.add_argument("--json", required=True, help="Extracted JSON file from GPT")
    parser.add_argument("--template", required=True, help="Excel template file")
    parser.add_argument("--output", required=True, help="Output Excel file")
    args = parser.parse_args()

    with open(args.json) as f:
        json_data = json.load(f)
    map_to_excel(json_data, args.template, args.output)