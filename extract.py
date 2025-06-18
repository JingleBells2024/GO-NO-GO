import sys
import os
import argparse
import openai
import json
from pdf2image import convert_from_path
from io import BytesIO

def pdf_to_images(pdf_path):
    # Convert PDF to PIL Images, in-memory
    images = convert_from_path(pdf_path)
    buffers = []
    for img in images:
        buf = BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        buffers.append(buf)
    return buffers

def all_files_to_images(files):
    # Accept PDF/JPG/PNG/Excel—just handle PDFs for vision
    all_images = []
    for file_path in files:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.pdf':
            print(f"Converting {os.path.basename(file_path)}...")
            all_images.extend(pdf_to_images(file_path))
        elif ext in ('.png', '.jpg', '.jpeg'):
            with open(file_path, "rb") as f:
                all_images.append(BytesIO(f.read()))
        # else: ignore others (Excel handled elsewhere)
    return all_images

def vision_extract(api_key, image_buffers, prompt):
    client = openai.OpenAI(api_key=api_key)

    # Prepare vision message (split into chunks of max 10 images per request)
    vision_results = []
    for i in range(0, len(image_buffers), 10):
        imgs = image_buffers[i:i+10]
        content = [
            {"type": "text", "text": prompt},
        ]
        for img_buf in imgs:
            # Encode to base64 using built-in Python, not openai._utils
            import base64
            b64_img = base64.b64encode(img_buf.read()).decode("utf-8")
            content.append({
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/png;base64,{b64_img}"
                }
            })
            img_buf.seek(0)
        resp = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "user", "content": content}
            ],
            max_tokens=4096,
            temperature=0
        )
        vision_results.append(resp.choices[0].message.content)
    # Combine multi-chunk results
    return "\n".join(vision_results)

def extract_json_from_text(text):
    # Use a simple regex to find JSON array in the output if present
    import re
    match = re.search(r'(\[\s*{.*}\s*\])', text, re.DOTALL)
    if match:
        return json.loads(match.group(1))
    # else: attempt full parse (if user prompt tells GPT to output only JSON)
    try:
        return json.loads(text)
    except Exception:
        # Not pure JSON—raise error
        raise ValueError("Could not parse JSON from GPT output.")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--key', required=True, help='OpenAI API key')
    parser.add_argument('--prompt', required=True, help='Prompt for GPT extraction')
    parser.add_argument('-o', '--output', default='gpt4o_extracted.json', help='Output JSON file')
    parser.add_argument('files', nargs='+', help='List of PDF/image files to extract from')
    args = parser.parse_args()

    # 1. Convert all files to images (in memory)
    all_images = all_files_to_images(args.files)
    print(f"Extracting {len(all_images)} pages from {len(args.files)} file(s)...")
    if not all_images:
        print("No images found in provided files.")
        sys.exit(1)

    # 2. Call GPT-4o Vision with prompt + images
    result_text = vision_extract(args.key, all_images, args.prompt)

    # 3. Parse out JSON array from result (if present)
    try:
        data = extract_json_from_text(result_text)
    except Exception as e:
        print("Could not extract JSON from GPT output:", e)
        print("GPT output was:\n", result_text)
        sys.exit(1)

    # 4. Save JSON output
    with open(args.output, "w") as f:
        json.dump(data, f, indent=2)
    print(f"Extracted data saved to {args.output}")

if __name__ == "__main__":
    main()