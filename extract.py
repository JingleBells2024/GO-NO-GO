#!/usr/bin/env python3
import sys
import os
import argparse
import openai
import base64
from pdf2image import convert_from_path
from PIL import Image

# --- Hard-coded prompt (edit as needed) ---
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
    """Converts a PDF into a list of images (one per page)."""
    images = convert_from_path(pdf_path)
    img_paths = []
    for i, img in enumerate(images):
        img_path = f"{os.path.splitext(pdf_path)[0]}_page{i+1}.png"
        img.save(img_path, "PNG")
        img_paths.append(img_path)
    return img_paths

def encode_image(image_path):
    """Read image and base64 encode it for Vision API."""
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

def vision_extract(api_key, image_paths, prompt):
    """Call GPT-4o Vision with each image, concatenate results."""
    client = openai.OpenAI(api_key=api_key)
    all_text = ""
    for img_path in image_paths:
        encoded = encode_image(img_path)
        messages = [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": f"data:image/png;base64,{encoded}"}
                ]
            }
        ]
        resp = client.chat.completions.create(
            model="gpt-4o",
            messages=messages,
            max_tokens=2048,
            temperature=0
        )
        page_text = resp.choices[0].message.content.strip()
        all_text += "\n" + page_text
        print(f"\n--- Extracted text from {img_path} ---\n{page_text}\n")
    return all_text.strip()

def main():
    parser = argparse.ArgumentParser(description="Extract financial data from PDF(s) using GPT-4o Vision.")
    parser.add_argument("--key", required=True, help="OpenAI API key")
    parser.add_argument("--files", nargs="+", required=True, help="PDF files to extract from")
    parser.add_argument("-o", "--output", default="extracted_data.txt", help="Output text file")
    args = parser.parse_args()

    all_images = []
    for pdf in args.files:
        imgs = pdf_to_images(pdf)
        all_images.extend(imgs)

    print(f"Extracting {len(all_images)} pages from {len(args.files)} PDF(s)...")

    result_text = vision_extract(args.key, all_images, VISION_PROMPT)

    with open(args.output, "w") as f:
        f.write(result_text)
    print(f"\nAll extracted text saved to {args.output}")

    # Optional: If you want to convert formatted text to JSON automatically, you could do it here.

if __name__ == "__main__":
    main()