import sys
import os
import json
import argparse
import openai
from pdf2image import convert_from_path
from io import BytesIO
import base64

# -- YOUR FULL HARDCODED PROMPT --
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
    "Plus Owner Salary+Super etc": ...,
    "Plus Owner Benefits": ...,
    "Total add backs": ...
  },
  {
    "year": 2023,
    "Revenue": ...,
    "Cost of Goods Sold (COGS)": ...,
    "Less Operating Expenses": ...,
    "Other Income": ...,
    "Plus Owner Salary+Super etc": ...,
    "Plus Owner Benefits": ...,
    "Total add backs": ...
  },
  {
    "year": 2024,
    "Revenue": ...,
    "Cost of Goods Sold (COGS)": ...,
    "Less Operating Expenses": ...,
    "Other Income": ...,
    "Plus Owner Salary+Super etc": ...,
    "Plus Owner Benefits": ...,
    "Total add backs": ...
  },
  {
    "year": 2025,
    "Revenue": ...,
    "Cost of Goods Sold (COGS)": ...,
    "Less Operating Expenses": ...,
    "Other Income": ...,
    "Plus Owner Salary+Super etc": ...,
    "Plus Owner Benefits": ...,
    "Total add backs": ...
  }
]
Do not add or remove fields. Do not estimate or infer values. Only copy numbers as written.
If a field is missing for a year, use 0.
Return only the JSON array—no commentary, no formatting, no markdown.
"""

def build_final_prompt(user_prompt):
    # Always combine the hardcoded prompt with the user's, if provided
    if user_prompt:
        return HARDCODED_PROMPT.strip() + "\n\n" + user_prompt.strip()
    else:
        return HARDCODED_PROMPT.strip()

def pdf_to_images(pdf_path):
    images = convert_from_path(pdf_path)
    img_buffers = []
    for img in images:
        buf = BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        img_buffers.append(buf)
    return img_buffers

def encode_image_buf(img_buf):
    return base64.b64encode(img_buf.read()).decode("utf-8")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--key', required=True, help="OpenAI API key")
    parser.add_argument('--prompt', required=False, help="User extraction prompt (optional)")
    parser.add_argument('files', nargs='+', help="Financial files (PDF, images, etc.)")
    parser.add_argument('-o', '--output', default="gpt4o_extracted.json", help="JSON output path")
    args = parser.parse_args()

    api_key = args.key
    user_prompt = args.prompt or ""
    all_images = []

    # Convert each file to image(s) in memory
    for file in args.files:
        ext = os.path.splitext(file)[1].lower()
        if ext == '.pdf':
            all_images += pdf_to_images(file)
        elif ext in ['.png', '.jpg', '.jpeg']:
            with open(file, "rb") as f:
                buf = BytesIO(f.read())
                img_buffers = [buf]
                all_images += img_buffers
        else:
            print(f"Unsupported file type: {file}")
            continue

    if not all_images:
        print("No valid images to process.")
        sys.exit(1)

    # Prepare the OpenAI Vision API call
    client = openai.OpenAI(api_key=api_key)
    final_prompt = build_final_prompt(user_prompt)

    messages = [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": final_prompt}
            ] + [
                {
                    "type": "image_url",
                    "image_url": {
                        "url": "data:image/png;base64," + encode_image_buf(img_buf)
                    }
                } for img_buf in all_images
            ]
        }
    ]

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=messages,
        max_tokens=4096,
        temperature=0
    )
    gpt_output = response.choices[0].message.content.strip()
    try:
        extracted_data = json.loads(gpt_output)
    except Exception:
        print("Could not extract JSON from GPT output: Could not parse JSON from GPT output.")
        print("GPT output was:\n", gpt_output)
        sys.exit(1)

    # Write the extracted JSON to file
    with open(args.output, "w") as f:
        json.dump(extracted_data, f, indent=2)
    print(f"Extracted data saved to {args.output}")

if __name__ == '__main__':
    main()