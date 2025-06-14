#!/usr/bin/env python3
import os
import json
import openai
import argparse

def extract_with_gpt(extracted_data: dict, user_prompt: str, api_key: str) -> dict:
    openai.api_key = api_key
    messages = [
        {"role": "system",
         "content": "You are a financial data extractor. "
                    "Given raw P&L data, return ONLY a JSON object "
                    "mapping each field name to its numeric value."},
        {"role": "user", "content": json.dumps(extracted_data)},
        {"role": "user", "content": user_prompt}
    ]
    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=messages,
        temperature=0
    )
    body = response.choices[0].message.content.strip()
    return json.loads(body)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", required=True, help="Path to raw-extracted JSON file")
    parser.add_argument("--prompt", required=True, help="The user’s GPT prompt")
    parser.add_argument("--key", required=True, help="OpenAI API key")
    args = parser.parse_args()

    extracted = json.load(open(args.data))
    values = extract_with_gpt(extracted, args.prompt, args.key)

    # Output structured JSON to stdout
    print(json.dumps(values, indent=2))