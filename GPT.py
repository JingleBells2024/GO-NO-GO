#!/usr/bin/env python3
import os
import json
import openai
from openpyxl import load_workbook


def extract_with_gpt(extracted_data: dict, user_prompt: str, api_key: str) -> dict:
    """
    Calls OpenAI to turn raw extracted_data + user_prompt 
    into a clean JSON of field→value.
    """
    openai.api_key = api_key
    messages = [
        {"role": "system",
         "content": "You are a financial data extractor.  "
                    "Given raw P&L data, return ONLY a JSON object "
                    "mapping each field name to its numeric value."},
        {"role": "user", "content": json.dumps(extracted_data)},
        {"role": "user", "content": user_prompt}
    ]
    resp = openai.ChatCompletion.create(
        model="gpt-4",
        messages=messages,
        temperature=0
    )
    body = resp.choices[0].message.content.strip()
    return json.loads(body)


def fill_template(template_path: str,
                  output_path: str,
                  values: dict,
                  cell_map: dict):
    """
    Writes each value into its mapped cell in the template,
    then saves to output_path.
    """
    wb = load_workbook(template_path)
    ws = wb.active

    for field, cell in cell_map.items():
        if field in values:
            ws[cell] = values[field]

    wb.save(output_path)


if __name__ == "__main__":
    import argparse
    import sys
    import shlex
    import subprocess

    parser = argparse.ArgumentParser()
    parser.add_argument("--template", required=True, help="Path to .xlsx template")
    parser.add_argument("--data",     required=True, help="Path to raw-extracted JSON file")
    parser.add_argument("--prompt",   required=True, help="The user’s GPT prompt")
    parser.add_argument("--year",     type=int, required=True, help="Year column to fill")
    parser.add_argument("--key",      required=True, help="OpenAI API key")
    args = parser.parse_args()

    # If not already running inside a Terminal window, spawn one and exit.
    # Detect via an environment variable to avoid recursion.
    if os.environ.get("AI_PROC_IN_TERMINAL") is None:
        # Build command to re-run this script in new Terminal
        script = os.path.abspath(__file__)
        cmd_args = ' '.join(shlex.quote(arg) for arg in sys.argv[1:])
        cmd = f'python3 {shlex.quote(script)} {cmd_args}'
        # Launch in new Terminal window (macOS)
        subprocess.Popen([
            'osascript', '-e',
            f'tell application "Terminal" to do script ' + shlex.quote(cmd)
        ], env={**os.environ, "AI_PROC_IN_TERMINAL": "1"})
        sys.exit(0)

    # 1) load extracted data
    extracted = json.load(open(args.data))

    # 2) ask GPT to structure it
    values = extract_with_gpt(extracted, args.prompt, args.key)

    # 3) pick the right cell_map for that year
    maps = {
        2023: {"Revenue":"D10","COGS":"D11","NetProfit":"D12"},
        2024: {"Revenue":"E10","COGS":"E11","NetProfit":"E12"},
        2025: {"Revenue":"F10","COGS":"F11","NetProfit":"F12"},
    }
    cell_map = maps.get(args.year, {})

    # 4) fill and save
    out = os.path.splitext(args.template)[0] + f"_filled_{args.year}.xlsx"
    fill_template(args.template, out, values, cell_map)
    print(f"Finished! Output saved to {out}")