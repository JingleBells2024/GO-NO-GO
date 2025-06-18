#!/usr/bin/env python3
import os
import sys
import json
import argparse
from compiler import map_to_excel

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", required=True, help="Path to extracted JSON file (e.g., gpt4o_extracted.json)")
    parser.add_argument("--template", required=True, help="Excel template file to fill")
    parser.add_argument("--output", help="Output Excel file (leave blank to overwrite template)")
    args = parser.parse_args()

    # Load the extracted data (already in JSON format)
    with open(args.data) as f:
        data = json.load(f)

    # Accept both single dict or list of dicts
    if isinstance(data, dict):
        records = [data]
    elif isinstance(data, list):
        records = data
    else:
        raise Exception("Extracted JSON is not dict or list.")

    # Map to Excel using your mapping function
    map_to_excel(records, args.template, args.output)

    print("Done: Excel file updated.")

if __name__ == "__main__":
    main()