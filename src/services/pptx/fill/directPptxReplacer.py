import sys
import os
import json
import pptx

def replace_text_preserve_runs(shape, new_text):
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    if not tf.paragraphs:
        return
    
    first_p = tf.paragraphs[0]
    if first_p.runs:
        first_p.runs[0].text = new_text
        for r in first_p.runs[1:]:
            r.text = ""
    else:
        first_p.text = new_text

def fill_master_pptx_deck(master_pptx_path, fill_plan_json_path, output_pptx_path):
    print(f"Executing Direct Master Template Replacement on: '{master_pptx_path}'")
    prs = pptx.Presentation(master_pptx_path)

    with open(fill_plan_json_path, "r", encoding="utf-8") as f:
        fill_plan = json.load(f)

    selected_slides_data = fill_plan.get("selectedSlides", [])
    print(f"Total Selected Slides in Plan: {len(selected_slides_data)}")

    template_keywords = [
        "copilot", "i-avantage", "acculturation", "démystifier", "demystifier",
        "biais", "intelligences", "question", "aujourd’hui", "aujourd'hui",
        "module 1", "module 2", "présenté par", "presente par"
    ]

    selected_indices = set()
    for slide_data in selected_slides_data:
        slide_idx = slide_data.get("slideIndex", 1) - 1 # 0-indexed
        fill_content = slide_data.get("fillContent", {})

        if 0 <= slide_idx < len(prs.slides):
            selected_indices.add(slide_idx)
            target_slide = prs.slides[slide_idx]
            
            for shape in target_slide.shapes:
                if shape.has_text_frame:
                    if shape.name in fill_content:
                        new_text = str(fill_content[shape.name])
                        if new_text.strip():
                            replace_text_preserve_runs(shape, new_text)
                    else:
                        # Check for unmapped residual template text and sanitize
                        current_txt = shape.text.strip().lower()
                        if any(kw in current_txt for kw in template_keywords):
                            replace_text_preserve_runs(shape, "")

    # Prune unselected slides from the presentation
    total_slides = len(prs.slides)
    for i in range(total_slides - 1, -1, -1):
        if i not in selected_indices:
            rId = prs.slides._sldIdLst[i].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[i]

    prs.save(output_pptx_path)
    print(f"Saved Perfect Master-Filled PPTX File: {output_pptx_path} (Final Slides: {len(prs.slides)}, Size: {os.stat(output_pptx_path).st_size} bytes)")

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: python directPptxReplacer.py <master_pptx> <fill_plan_json> <output_pptx>")
        sys.exit(1)

    fill_master_pptx_deck(sys.argv[1], sys.argv[2], sys.argv[3])
