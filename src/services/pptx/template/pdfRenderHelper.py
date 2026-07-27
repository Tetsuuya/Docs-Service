import sys
import os
import json
import pypdfium2 as pdfium
from PIL import Image, ImageDraw

def render_and_clean_pdf_slides(pdf_path, slide_requests, output_dir, scale=2.5):
    """
    Renders PDF slide pages into high-res images and patches text container regions
    to erase old template dummy text while retaining 100% of Canva background graphics & framing.
    
    slide_requests: list of dicts e.g. [{"slideIndex": 1, "category": "title_slide"}, ...]
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    pdf = pdfium.PdfDocument(pdf_path)
    total_pages = len(pdf)
    rendered_files = {}

    for req in slide_requests:
        idx = req.get("slideIndex", 1)
        category = req.get("category", "bullets")
        
        page_num = idx - 1
        if 0 <= page_num < total_pages:
            img = pdf[page_num].render(scale=scale).to_pil()
            w, h = img.size
            draw = ImageDraw.Draw(img)

            # Clean text regions based on slide layout category
            if category == "title_slide" or idx == 1:
                # Patch central title & subtitle box (retain outer Canva branding/margins)
                draw.rectangle([int(w * 0.08), int(h * 0.20), int(w * 0.92), int(h * 0.80)], fill=(255, 255, 255))
            elif category == "3_column_cards":
                # Patch header title area
                draw.rectangle([int(w * 0.05), int(h * 0.04), int(w * 0.95), int(h * 0.16)], fill=(255, 255, 255))
                # Patch 3 card container text areas
                draw.rectangle([int(w * 0.06), int(h * 0.20), int(w * 0.33), int(h * 0.88)], fill=(248, 250, 252))
                draw.rectangle([int(w * 0.37), int(h * 0.20), int(w * 0.64), int(h * 0.88)], fill=(248, 250, 252))
                draw.rectangle([int(w * 0.68), int(h * 0.20), int(w * 0.95), int(h * 0.88)], fill=(248, 250, 252))
            elif category == "2_column_comparison":
                # Patch header title area
                draw.rectangle([int(w * 0.05), int(h * 0.04), int(w * 0.95), int(h * 0.16)], fill=(255, 255, 255))
                # Patch 2 column text areas
                draw.rectangle([int(w * 0.06), int(h * 0.20), int(w * 0.48), int(h * 0.88)], fill=(248, 250, 252))
                draw.rectangle([int(w * 0.52), int(h * 0.20), int(w * 0.94), int(h * 0.88)], fill=(248, 250, 252))
            elif category == "stat_callout":
                # Patch header title area
                draw.rectangle([int(w * 0.05), int(h * 0.04), int(w * 0.95), int(h * 0.16)], fill=(255, 255, 255))
                # Patch stat box & right text box
                draw.rectangle([int(w * 0.06), int(h * 0.20), int(w * 0.42), int(h * 0.88)], fill=(15, 23, 42)) # Dark stat card
                draw.rectangle([int(w * 0.46), int(h * 0.20), int(w * 0.94), int(h * 0.88)], fill=(248, 250, 252))
            else:
                # General header & body text container patch
                draw.rectangle([int(w * 0.05), int(h * 0.04), int(w * 0.95), int(h * 0.16)], fill=(255, 255, 255))
                draw.rectangle([int(w * 0.06), int(h * 0.20), int(w * 0.94), int(h * 0.88)], fill=(248, 250, 252))

            output_filename = f"clean_slide_page_{idx}.png"
            output_path = os.path.join(output_dir, output_filename)
            img.save(output_path)
            rendered_files[idx] = output_path.replace("\\", "/")

    return rendered_files

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(json.dumps({"error": "Missing parameters"}))
        sys.exit(1)
        
    pdf_path = sys.argv[1]
    slide_requests_json = sys.argv[2] # JSON string of slide requests e.g. '[{"slideIndex":1,"category":"title_slide"}]'
    output_dir = sys.argv[3] if len(sys.argv) > 3 else "temp/pdf_slides"
    
    requests = json.loads(slide_requests_json)
    results = render_and_clean_pdf_slides(pdf_path, requests, output_dir)
    print(json.dumps(results))
