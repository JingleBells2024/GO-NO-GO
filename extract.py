import sys
import os
import argparse
import openai
import json
from pdf2image import convert_from_path
from io import BytesIO

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

def pdf_to_images(pdf_path):
    # Convert all pages to images (kept in memory, not saved to disk)
    images = convert_from_path(pdf_path)
    img_bytes_list = []
    for img in images:
        buf = BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        img_bytes_list.append(buf.read())
    return img_bytes_list

def encode_image(image_bytes):
    import base64
    return base64.b64encode(image_bytes).decode("utf-8")

def vision_extract(api_key, images, prompt):
    client = openai.OpenAI(api_key=api_key)
    # Prepare images as base64-encoded PNGs for OpenAI Vision API
    vision_inputs = []
    for img_bytes in images:
        b64img = encode_image(img_bytes)
        vision_inputs.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{b64img}"
            }
        })

    # Compose messages for GPT-4o Vision API
    messages = [
        {"role": "system", "content": "You are a helpful assistant."},
        {"role": "user", "content": [
            {"type": "text", "text": prompt}
        ] + vision_inputs}
    ]
    resp = client.chat.completions.create(
        model="gpt-4o",
        messages=messages,
        max_tokens=2048,
        temperature=0,
    )
    output = resp.choices[0].message.content.strip()
    return output

def extract_to_json(formatted_text, output_path):
    # Parse the extracted text to a list of yearly dicts (strict format)
    import re
    years = []
    current = None
    for line in formatted_text.splitlines():
        line = line.strip()
        m = re.match(r"^(\d{4})$", line)
        if m:
            if current:
                years.append(current)
            current = {"year": int(m.group(1))}
        elif current and line.startswith("•"):
            k, v = line[1:].split(":", 1)
            k = k.strip()
            v = v.strip().replace(",", "")
            # Try to coerce to float, else set to 0
            try:
                val = float(v)
            except Exception:
                val = 0.0
            current[k] = val
    if current:
        years.append(current)
    with open(output_path, "w") as f:
        json.dump(years, f, indent=2)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--key', required=True, help="OpenAI API key")
    parser.add_argument('-o', '--output', default="gpt4o_extracted.json", help="Output JSON file")
    parser.add_argument('files', nargs='+', help="One or more PDF files to extract from")
    args = parser.parse_args()

    all_images = []
    for file in args.files:
        print(f"Converting {os.path.basename(file)}...")
        all_images.extend(pdf_to_images(file))
    print(f"Extracting {len(all_images)} pages from {len(args.files)} PDF(s)...")

    formatted_text = vision_extract(args.key, all_images, VISION_PROMPT)
    # Save JSON only (no txt summary)
    extract_to_json(formatted_text, args.output)
    print(f"Extracted data saved to {args.output}")

if __name__ == "__main__":
    main()