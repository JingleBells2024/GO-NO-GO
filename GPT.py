#!/usr/bin/env python3
import os
import json
import openai
import argparse

def extract_with_gpt(extracted_data: dict, user_prompt: str, api_key: str) -> dict:
    client = openai.OpenAI(api_key=api_key)

    fixed_categories_example = {
        "year": 2024,
        "Revenue": 964021.78,
        "Cost of Goods Sold (COGS)": 156873.40,
        "Operating Expenses": 311169.35,
        "Taxes": 30000,
        "Depreciation & Amortization": 20000,
        "Plus Interest": 0,
        "Owner Salary+Super": 50000,
        "Owner Benefits": 200000,
        "Manager Salary": 105000,
        "Investor Salary": 70000,
        "One off Revenue Adjustments": 0,
        "One off Expenses Adjustments": 0,
        "Other Adjustments 1": 0,
        "Other Adjustments 2": 0,
        "Assets": 1234567.89,
        "Liabilities": 234567.89,
        "Equity": 1000000,
        "Other Income": 0
    }

    system_content = (
        "You are a financial data extractor. "
        "Given raw financial data in various forms, map all values into the following fixed categories exactly as named: \n"
        + json.dumps(fixed_categories_example, indent=2) + "\n"
        "You must interpret synonyms and variations in labels to fit these categories. "
        "Return ONLY a JSON object containing these categories and their numeric values. "
        "Include the year as the 'year' key. "
        "Do not include any percentages or fields not listed. "
        "If a category is missing in the input, set its value to 0."
    )

    messages = [
        {"role": "system", "content": system_content},
        {"role": "user", "content": json.dumps(extracted_data)},
        {"role": "user", "content": user_prompt}
    ]

    resp = client.chat.completions.create(
        model="gpt-4",
        messages=messages,
        temperature=0
    )
    body = resp.choices[0].message.content.strip()
    values = json.loads(body)
    return values

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", required=True, help="Path to raw-extracted JSON file")
    parser.add_argument("--prompt", required=True, help="User's GPT prompt")
    parser.add_argument("--key", required=True, help="OpenAI API key")
    args = parser.parse_args()

    with open(args.data) as f:
        extracted = json.load(f)
    structured_data = extract_with_gpt(extracted, args.prompt, args.key)
    # Print to stdout so GUI can capture directly
    print(json.dumps(structured_data, indent=2))