#!/usr/bin/env python3
import sys
import os
import argparse
import openai
import io
from pdf2image import convert_from_path
from PIL import Image

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

def pdf_to_images_in_memory(pdf_path):
    images = convert_from_path(pdf_path)
    image_bytes_list = []
    for img in images:
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        image_bytes_list.append(buf)
    return image_bytes_list

def call_gpt4o_vision(api_key, images, prompt):
    client = openai.OpenAI(api_key=api_key)
    content = [{"type": "text", "text": prompt}]
    for img_buf in images:
        img_buf.seek(0)
        b64_img = openai._utils._to_base64(img_buf.read())
        content.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{b64_img}"
            }
        })
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "user", "content": content}
        ],
        max_tokens=4096,
        temperature=0,
    )
    return response.choices[0].message.content

def main():
    parser = argparse.ArgumentParser(description="Extracts formatted text from PDFs using GPT-4o Vision.")
    parser.add_argument('--key', required=True, help="OpenAI API key")
    parser.add_argument('files', nargs='+', help="Path(s) to one or more PDF files")
    parser.add_argument("-o", "--output", help="Output file for formatted text", default="formatted_output.txt")
    args = parser.parse_args()

    all_images = []
    for fpath in args.files:
        print(f"Converting {os.path.basename(fpath)}...")
        all_images.extend(pdf_to_images_in_memory(fpath))

    print(f"Extracting {len(all_images)} pages from {len(args.files)} PDF(s)...")
    formatted_text = call_gpt4o_vision(args.key, all_images, VISION_PROMPT)

    with open(args.output, "w") as f:
        f.write(formatted_text.strip())
    print(f"Formatted output saved to {args.output}")

if __name__ == "__main__":
    main()