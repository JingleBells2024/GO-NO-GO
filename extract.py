#!/usr/bin/env python3
import argparse
import openai
import os
import json
import re

# --- Prompt ---
VISION_PROMPT = """
Instructions:
I am uploading several financial documents.
Please extract all relevant yearly financial data and summarize it in this exact text format, for every year present in the documents:

2022
• Revenue: 
• Cost of Goods Sold (COGS):
• Less Operating Expenses:
• Other Income:
• Plus Owner Salary+Super etc:
• Plus Owner Benefits:
• Total add backs:
2023
(and so on for each year...)

If any value is missing, set it to 0.
Use exactly these categories and wording for each year.
Do not add or remove any fields.
Do not provide explanations, just the formatted text.
"""

def parse_gpt_output(text):
    # Parse the text block into a JSON array
    years = re.split(r'(?=\d{4}\n)', text)
    categories = [
        "Revenue",
        "Cost of Goods Sold (COGS)",
        "Less Operating Expenses",
        "Other Income",
        "Plus Owner Salary+Super etc",
        "Plus Owner Benefits",
        "Total add backs"
    ]
    data = []
    for year_block in years:
        lines = year_block.strip().split('\n')
        entry = {}
        for line in lines:
            if line.strip().isdigit():
                entry['year'] = int(line.strip())
            else:
                match = re.match(r'• (.*?):\s*([0-9,\.]*)', line.strip())
                if match:
                    k, v = match.groups()
                    if k in categories:
                        entry[k] = float(v.replace(',', '')) if v else 0
        if entry.get('year'):
            # Make sure all categories are present
            for cat in categories:
                if cat not in entry:
                    entry[cat] = 0
            data.append(entry)
    return data

def call_gpt4o_vision(api_key, filepaths, prompt=VISION_PROMPT):
    client = openai.OpenAI(api_key=api_key)
    files = []
    try:
        # Upload all files to OpenAI
        for path in filepaths:
            files.append({"type": "file", "file": open(path, "rb")})
        messages = [
            {"role": "system", "content": "You are a financial document data extractor."},
            {"role": "user", "content": [
                {"type": "text", "text": prompt},
                *files
            ]}
        ]
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=messages,
            max_tokens=2048,
            temperature=0
        )
        text = response.choices[0].message.content.strip()
        return text
    finally:
        for f in files:
            if hasattr(f['file'], 'close'):
                f['file'].close()

def main():
    parser = argparse.ArgumentParser(description="Extracts yearly financial data using GPT-4o Vision.")
    parser.add_argument('--key', required=True, help="OpenAI API key")
    parser.add_argument('--files', nargs='+', required=True, help="Paths to PDF, image, or Excel files")
    parser.add_argument('--output', default='gpt4o_extracted.txt', help="Text output file")
    parser.add_argument('--json', default='gpt4o_extracted.json', help="JSON output file")
    args = parser.parse_args()

    # Run GPT-4o Vision
    formatted_text = call_gpt4o_vision(args.key, args.files)
    print("\n--- Extracted formatted text ---\n")
    print(formatted_text)
    with open(args.output, 'w') as f:
        f.write(formatted_text)

    # Parse and save JSON
    data = parse_gpt_output(formatted_text)
    print("\n--- Parsed JSON ---\n")
    print(json.dumps(data, indent=2))
    with open(args.json, 'w') as f:
        json.dump(data, f, indent=2)

if __name__ == "__main__":
    main()