#!/usr/bin/env python3
import os
import json
import openai
import argparse

def extract_with_gpt(extracted_data: dict, user_prompt: str, api_key: str) -> dict:
    openai.api_key = api_key
    system_message = (
        "You are a financial data extractor. Given raw P&L data, "
        "return ONLY a JSON object mapping each field name to its numeric value. "
        "Your response must include the year from the user's prompt as '_year' in the output JSON."
    )
    messages = [
        {"role": "system", "content": system_message},
        {"role": "user", "content": json.dumps(extracted_data)},
        {"role": "user", "content": user_prompt}
    ]
    resp = openai.ChatCompletion.create(
        model="gpt-4",
        messages=messages,
        temperature=0
    )
    body = resp.choices[0].message.content.strip()
    values = json.loads(body)
    return values

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--template", required=True, help="Path to .xlsx template (passed for reference)")
    parser.add_argument("--data",     required=True, help="Path to raw-extracted JSON file")
    parser.add_argument("--prompt",   required=True, help="The user’s GPT prompt")
    parser.add_argument("--key",      required=True, help="OpenAI API key")
    args = parser.parse_args()

    extracted = json.load(open(args.data))
    structured_data = extract_with_gpt(extracted, args.prompt, args.key)

    # Output structured JSON only, including the _year field
    print(json.dumps(structured_data, indent=2))