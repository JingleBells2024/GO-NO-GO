#!/usr/bin/env python3
import os
import json
import openai
import argparse
import sys

def extract_with_gpt(extracted_data, user_prompt, api_key):
    client = openai.OpenAI(api_key=api_key)

    fixed_categories_example = {
        "year": 2024,
        "Revenue": 964021.78,
        "Cost of Goods Sold (COGS)": 156873.40,
        "Less Operating Expenses": 311169.35,
        "Other Income": 47252.79,
        "Plus Owner Salary+Super etc": 116923,
        "Plus Owner Benefits": 7251,
        "Total add backs": 142230
    }
    system_content = (
        "You are a financial data extractor. "
        "Given structured financial data, return ONLY a JSON array of yearly records, each using these exact keys:\n"
        + json.dumps(fixed_categories_example, indent=2) +
        "\nInclude the 'year' field. "
        "For any missing value, use 0. "
        "Do not add, remove, or infer categories. "
        "Output ONLY a JSON array."
    )

    print("\n--- INPUT DATA TO GPT ---")
    print(json.dumps(extracted_data, indent=2))
    print("\n--- SYSTEM PROMPT (Categories) ---")
    print(json.dumps(fixed_categories_example, indent=2))
    print("\n--- USER PROMPT ---")
    print(user_prompt)
    print("\n--- SENDING TO GPT... ---\n")

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
    print("--- RAW GPT RESPONSE ---")
    print(body)

    # Robust: try loading JSON output, error if invalid
    try:
        values = json.loads(body)
    except Exception as e:
        print("ERROR: GPT response is not valid JSON.")
        sys.exit(1)
    return values

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", required=True, help="Path to extracted JSON file (e.g., extracted_data.json)")
    parser.add_argument("--prompt", required=True, help="Prompt to send to GPT")
    parser.add_argument("--key", required=True, help="OpenAI API key")
    parser.add_argument("--output", default="GPT_output.json", help="Path for output file (default: GPT_output.json)")
    args = parser.parse_args()

    if not os.path.isfile(args.data):
        print(f"ERROR: File not found: {args.data}")
        sys.exit(1)

    with open(args.data) as f:
        try:
            extracted = json.load(f)
        except Exception as e:
            print("ERROR: Input file is not valid JSON.")
            sys.exit(1)

    structured_data = extract_with_gpt(extracted, args.prompt, args.key)
    print("\n--- FINAL STRUCTURED OUTPUT ---")
    print(json.dumps(structured_data, indent=2))

    # Save the GPT output
    with open(args.output, "w") as f:
        json.dump(structured_data, f, indent=2)
    print(f"\nGPT output saved to {args.output}")