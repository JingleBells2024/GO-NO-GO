import sys
import os
import argparse
import json
import openai
from pdf2image import convert_from_path
import io
import re

def extract_json_from_gpt_output(gpt_output):
    # Remove triple backticks and optional language tags
    cleaned = re.sub(r"```(?:json)?", "", gpt_output, flags=re.IGNORECASE).strip()
    # Extract JSON array if present
    match = re.search(r"($begin:math:display$\\s*{.*?}\\s*$end:math:display$)", cleaned, re.DOTALL)
    if match:
        return match.group(1)
    return cleaned

def pdf_to_images(pdf_path):
    # Convert PDF to images (in-memory)
    images = convert_from_path(pdf_path)
    bufs = []
    for img in images:
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        bufs.append(buf)
    return bufs

def vision_extract(api_key, image_buffers, prompt):
    client = openai.OpenAI(api_key=api_key)
    messages = [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
            ] + [
                {"type": "image_url", "image_url": {"data": buf.getvalue(), "mime_type": "image/png"}}
                for buf in image_buffers
            ]
        }
    ]
    resp = client.chat.completions.create(
        model="gpt-4o",
        messages=messages,
        max_tokens=4096,
        temperature=0
    )
    return resp.choices[0].message.content.strip()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--key', required=True, help='OpenAI API key')
    parser.add_argument('--prompt', required=True, help='User/custom prompt (additional instructions)')
    parser.add_argument('files', nargs='+', help='List of PDF/image files')
    args = parser.parse_args()

    # Hardcoded technical requirements prompt (always appended)
    base_prompt = (
        "Return the extracted data in this exact JSON array format, one object per year. "
        "If a value is missing, use 0. Do not add or remove any fields. "
        "Do not include any explanation or markdown—only return the JSON array:\n\n"
        "[\n"
        "  {\n"
        "    \"year\": 2022,\n"
        "    \"Revenue\": ..., \n"
        "    \"Cost of Goods Sold (COGS)\": ..., \n"
        "    \"Less Operating Expenses\": ..., \n"
        "    \"Other Income\": ..., \n"
        "    \"Plus Owner Salary+Super etc\": ..., \n"
        "    \"Plus Owner Benefits\": ..., \n"
        "    \"Total add backs\": ...\n"
        "  },\n"
        "  ...\n"
        "]\n"
        "Do not include triple backticks or any markdown formatting. Output only the JSON array."
    )

    # Compose full prompt
    prompt = args.prompt.strip() + "\n\n" + base_prompt

    # Convert all PDFs to images (in memory)
    all_imgs = []
    for fpath in args.files:
        ext = os.path.splitext(fpath)[-1].lower()
        if ext == ".pdf":
            all_imgs += pdf_to_images(fpath)
        else:
            # Assume it's an image file
            with open(fpath, "rb") as imgf:
                img_buf = io.BytesIO(imgf.read())
                img_buf.seek(0)
                all_imgs.append(img_buf)

    print(f"Extracting {len(all_imgs)} page(s) from {len(args.files)} file(s)...")

    # Send all images and prompt to GPT-4o Vision
    gpt_output = vision_extract(args.key, all_imgs, prompt)
    print("--- RAW GPT OUTPUT ---\n", gpt_output)

    # Clean and extract JSON
    json_str = extract_json_from_gpt_output(gpt_output)
    try:
        data = json.loads(json_str)
    except Exception as e:
        print("Could not extract JSON from GPT output: Could not parse JSON from GPT output.")
        print("GPT output was:\n", gpt_output)
        sys.exit(1)

    # Save result to file
    with open('gpt4o_extracted.json', 'w') as f:
        json.dump(data, f, indent=2)
    print("Extracted data saved to gpt4o_extracted.json")

if __name__ == "__main__":
    main()