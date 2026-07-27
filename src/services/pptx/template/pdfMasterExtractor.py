import fitz  # PyMuPDF
import os
import json
from PIL import Image

def extract_full_pdf_template(pdf_path, output_dir="temp/template_extracted"):
    images_dir = os.path.join(output_dir, "slide_images")
    assets_dir = os.path.join(output_dir, "extracted_assets")
    os.makedirs(images_dir, exist_ok=True)
    os.makedirs(assets_dir, exist_ok=True)

    print(f"--- EXTRACTING ALL CONTENT FROM MASTER PDF: '{pdf_path}' ---")
    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    print(f"Total Pages: {total_pages}")

    extracted_slides = []

    for i in range(total_pages):
        page = doc[i]
        page_num = i + 1

        # 1. Render High-Res Page Pixmap (300 DPI)
        pix = page.get_pixmap(dpi=300)
        img_filename = f"slide_{page_num}.png"
        img_path = os.path.join(images_dir, img_filename)
        pix.save(img_path)

        # 2. Extract Text Blocks & Bounding Boxes
        blocks = page.get_text("blocks")
        text_blocks = []
        for b in blocks:
            if len(b) >= 5 and b[4].strip():
                text_blocks.append({
                    "bbox": [round(b[0], 2), round(b[1], 2), round(b[2], 2), round(b[3], 2)],
                    "text": b[4].strip()
                })

        # 3. Extract Embedded Image Assets
        image_list = page.get_images()
        page_assets = []
        for img_idx, img_info in enumerate(image_list):
            xref = img_info[0]
            try:
                base_img = doc.extract_image(xref)
                asset_filename = f"p{page_num}_asset_{img_idx+1}_{xref}.{base_img['ext']}"
                asset_path = os.path.join(assets_dir, asset_filename)
                with open(asset_path, "wb") as f:
                    f.write(base_img["image"])
                page_assets.append(asset_path.replace("\\", "/"))
            except Exception as e:
                pass

        # 4. Extract Drawings & Shape Rectangles
        drawings = page.get_drawings()
        shape_rects = []
        for d in drawings[:15]:
            r = d.get("rect")
            fill_col = d.get("fill")
            if r:
                shape_rects.append({
                    "rect": [round(r[0], 2), round(r[1], 2), round(r[2], 2), round(r[3], 2)],
                    "fill": fill_col
                })

        slide_data = {
            "slideIndex": page_num,
            "width": page.rect.width,
            "height": page.rect.height,
            "imagePath": img_path.replace("\\", "/"),
            "textBlocks": text_blocks,
            "embeddedAssets": page_assets,
            "shapeCount": len(drawings),
            "shapes": shape_rects
        }
        extracted_slides.append(slide_data)

    master_metadata = {
        "templateName": os.path.basename(pdf_path),
        "totalSlides": total_pages,
        "extractionTimestamp": os.stat(pdf_path).st_mtime,
        "slides": extracted_slides
    }

    metadata_path = os.path.join(output_dir, "master_template_data.json")
    with open(metadata_path, "w", encoding="utf-8") as f:
        json.dump(master_metadata, f, indent=2)

    print("\nCOMPLETE EXTRACTION SUCCESSFUL!")
    print(f"Extracted Master Data Saved to: {metadata_path}")
    print(f"Slide Background Images Saved to: {images_dir}")
    print(f"Embedded Assets Saved to: {assets_dir}")

    return master_metadata

if __name__ == "__main__":
    pdf_file = "template i-avantage.pdf"
    extract_full_pdf_template(pdf_file)
