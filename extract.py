#!/usr/bin/env python3
import os
import sys
import base64
import argparse
from pdf2image import convert_from_path
import openai

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
    images = convert_from_path(pdf_path)
    img_b64_list = []
    for img in images:
        from io import BytesIO
        buf = BytesIO()
        img.save(buf, format="PNG")
        b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
        img_b64_list.append(f"data:image/png;base64,{b64}")
    return img_b64_list

def vision_extract(api_key, all_images, prompt):
    client = openai.OpenAI(api_key=api_key)
    # Build the messages list
    messages = [
        {"role": "system", "content": "You are an expert financial data extractor."},
        {
            "role": "user",
            "content": [
                {"type": "text", "text": prompt}
            ] + [
                {"type": "image_url", "image_url": {"url": img_url}}
                for img_url in all_images
            ]
        }
    ]
    # Call GPT-4o with vision
    resp = client.chat.completions.create(
        model="gpt-4o",
        messages=messages,
        max_tokens=2048,
        temperature=0
    )
    return resp.choices[0].message.content.strip()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--key', required=True, help='OpenAI API key')
    parser.add_argument('--files', nargs='+', required=True, help='PDF file paths')
    parser.add_argument('-o', '--output', default='extracted_data.txt', help='Output text file')
    args = parser.parse_args()

    all_images = []
    print(f"Extracting pages from {len(args.files)} PDF(s)...")
    for pdf in args.files:
        imgs = pdf_to_images(pdf)
        all_images.extend(imgs)
    print(f"Total images: {len(all_images)}")

    result_text = vision_extract(args.key, all_images, VISION_PROMPT)

    with open(args.output, "w") as f:
        f.write(result_text)
    print(f"\nExtraction complete. Output saved to {args.output}")

if __name__ == "__main__":
    main()