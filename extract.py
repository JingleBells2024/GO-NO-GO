import sys
import os
import json
import argparse
import openai
from pdf2image import convert_from_path
from io import BytesIO
import base64

HARDCODED_PROMPT = """
Read the financial data below. For each year, copy the numbers into this exact schema—one record per year.
If a number is missing, use 0.
Return only valid JSON—no extra text.
[
  {
    "year": 2022,
    "Revenue": ...,
    "Cost of Goods Sold (COGS)": ...,
    "Less Operating Expenses": ...,
    "Other Income": ...,
    "Taxes": ...,
    "Plus Depreciation & Amortization": ...,
    "Plus Interest": ...,
    "Plus Taxes": ...,
    "Plus Owner Salary+Super etc": ...,
    "Plus Owner Benefits": ...,
    "Manager Salary": ...,
    "Investor Salary": ...,
    "One off Revenue Adjustments": ...,
    "One off Expenses Adjustments": ...,
    "Other Adjustments 1": ...,
    "Other Adjustments 2": ...,
    "Total add backs": ...,
    "Total SDE Adjustments": ...,
    "Total Adjustments": ...
  },
  {
    "year": 2023,
    "Revenue": ...,
    "Cost of Goods Sold (COGS)": ...,
    "Less Operating Expenses": ...,
    "Other Income": ...,
    "Taxes": ...,
    "Plus Depreciation & Amortization": ...,
    "Plus Interest": ...,
    "Plus Taxes": ...,
    "Plus Owner Salary+Super etc": ...,
    "Plus Owner Benefits": ...,
    "Manager Salary": ...,
    "Investor Salary": ...,
    "One off Revenue Adjustments": ...,
    "One off Expenses Adjustments": ...,
    "Other Adjustments 1": ...,
    "Other Adjustments 2": ...,
    "Total add backs": ...,
    "Total SDE Adjustments": ...,
    "Total Adjustments": ...
  },
  {
    "year": 2024,
    "Revenue": ...,
    "Cost of Goods Sold (COGS)": ...,
    "Less Operating Expenses": ...,
    "Other Income": ...,
    "Taxes": ...,
    "Plus Depreciation & Amortization": ...,
    "Plus Interest": ...,
    "Plus Taxes": ...,
    "Plus Owner Salary+Super etc": ...,
    "Plus Owner Benefits": ...,
    "Manager Salary": ...,
    "Investor Salary": ...,
    "One off Revenue Adjustments": ...,
    "One off Expenses Adjustments": ...,
    "Other Adjustments 1": ...,
    "Other Adjustments 2": ...,
    "Total add backs": ...,
    "Total SDE Adjustments": ...,
    "Total Adjustments": ...
  },
  {
    "year": 2025,
    "Revenue": ...,
    "Cost of Goods Sold (COGS)": ...,
    "Less Operating Expenses": ...,
    "Other Income": ...,
    "Taxes": ...,
    "Plus Depreciation & Amortization": ...,
    "Plus Interest": ...,
    "Plus Taxes": ...,
    "Plus Owner Salary+Super etc": ...,
    "Plus Owner Benefits": ...,
    "Manager Salary": ...,
    "Investor Salary": ...,
    "One off Revenue Adjustments": ...,
    "One off Expenses Adjustments": ...,
    "Other Adjustments 1": ...,
    "Other Adjustments 2": ...,
    "Total add backs": ...,
    "Total SDE Adjustments": ...,
    "Total Adjustments": ...
  }
]
Do not add or remove fields. Do not estimate or infer values. Only copy numbers as written.
If a field is missing for a year, use 0.
Return only the JSON array—no commentary, no formatting, no markdown.
"""

def pdfs_to_images(pdf_paths):
    images = []
    for pdf_path in pdf_paths:
        imgs = convert_from_path(pdf_path)
        for img in imgs:
            buf = BytesIO()
            img.save(buf, format='PNG')
            buf.seek(0)
            images.append(buf.read())
    return images

def vision_extract(api_key, images, user_prompt):
    client = openai.OpenAI(api_key=api_key)
    # Combine user prompt and hardcoded prompt
    combined_prompt = (user_prompt.strip() + "\n\n" + HARDCODED_PROMPT.strip()).strip()
    # Prepare message for GPT-4o vision
    contents = [{"type": "text", "text": combined_prompt}]
    for img in images:
        b64_img = base64.b64encode(img).decode('utf-8')
        contents.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{b64_img}"
            }
        })
    messages = [
        {"role": "user", "content": contents}
    ]
    resp = client.chat.completions.create(
        model="gpt-4o",
        messages=messages,
        max_tokens=4096,
        temperature=0
    )
    return resp.choices[0].message.content.strip()

def extract_json_from_output(output):
    # Try to find the first JSON array in the output
    try:
        start = output.index('[')
        end = output.rindex(']') + 1
        json_str = output[start:end]
        return json.loads(json_str)
    except Exception as e:
        raise RuntimeError(f"Could not parse JSON from GPT output.\nGPT output was:\n{output}")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--key', required=True, help='OpenAI API key')
    parser.add_argument('--prompt', required=True, help='User extraction prompt')
    parser.add_argument('files', nargs='+', help='List of PDF files')
    parser.add_argument('-o', '--output', default='gpt4o_extracted.json')
    args = parser.parse_args()

    print(f"Extracting {sum(convert_from_path(f).__len__() for f in args.files)} page(s) from {len(args.files)} file(s)...")
    images = pdfs_to_images(args.files)
    gpt_output = vision_extract(args.key, images, args.prompt)
    try:
        extracted = extract_json_from_output(gpt_output)
    except Exception as e:
        print(str(e))
        sys.exit(1)
    with open(args.output, 'w') as f:
        json.dump(extracted, f, indent=2)
    print(f"Extracted data saved to {args.output}")

if __name__ == "__main__":
    main()