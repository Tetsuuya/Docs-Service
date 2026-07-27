import sys
import os
import json
import pptx
from pptx.dml.color import RGBColor

def rgb_to_hex(rgb):
    if isinstance(rgb, RGBColor):
        return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    return None

def extract_deep_pptx_metadata(file_path):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"PPTX file not found: {file_path}")

    prs = pptx.Presentation(file_path)
    slide_width_in = round(prs.slide_width / 914400.0, 2)
    slide_height_in = round(prs.slide_height / 914400.0, 2)
    
    fonts_found = set()
    colors_found = set()
    slides_data = []

    for idx, slide in enumerate(prs.slides):
        slide_info = {
            "slideIndex": idx + 1,
            "layoutName": slide.slide_layout.name if slide.slide_layout else "Custom",
            "elements": []
        }
        
        text_elements_summary = []
        shape_names = []

        for shape in slide.shapes:
            shape_names.append(shape.name)
            elem = {
                "shapeId": shape.shape_id,
                "shapeName": shape.name,
                "shapeType": str(shape.shape_type),
                "position": {
                    "left_in": round(shape.left / 914400.0, 2),
                    "top_in": round(shape.top / 914400.0, 2),
                    "width_in": round(shape.width / 914400.0, 2),
                    "height_in": round(shape.height / 914400.0, 2)
                }
            }

            try:
                if shape.fill.type == 1:
                    rgb = shape.fill.fore_color.rgb
                    if rgb:
                        hex_c = rgb_to_hex(rgb)
                        elem["fillColor"] = hex_c
                        colors_found.add(hex_c)
            except Exception:
                pass

            if shape.has_text_frame:
                tf = shape.text_frame
                elem["text"] = tf.text.strip()
                paragraphs_info = []

                for p in tf.paragraphs:
                    p_text = p.text.strip()
                    if not p_text:
                        continue

                    font_names = set()
                    font_sizes = set()
                    font_colors = set()

                    for r in p.runs:
                        if r.font.name:
                            font_names.add(r.font.name)
                            fonts_found.add(r.font.name)
                        if r.font.size:
                            font_sizes.add(round(r.font.size / 12700.0, 1))
                        if r.font.color and hasattr(r.font.color, 'rgb') and r.font.color.rgb:
                            hex_c = rgb_to_hex(r.font.color.rgb)
                            font_colors.add(hex_c)
                            colors_found.add(hex_c)

                    paragraphs_info.append({
                        "textSample": p_text[:80],
                        "alignment": str(p.alignment) if p.alignment else "LEFT",
                        "fontNames": list(font_names),
                        "fontSizesPt": list(font_sizes),
                        "colors": list(font_colors)
                    })

                elem["paragraphs"] = paragraphs_info
                if tf.text.strip():
                    text_elements_summary.append(f"{shape.name}: {tf.text.strip()[:50]}")

            slide_info["elements"].append(elem)

        if idx == 0:
            category = "title_slide"
        elif len(text_elements_summary) >= 3:
            category = "3_column_cards"
        else:
            category = "content_slide"

        slide_info["layoutCategory"] = category
        slide_info["shapeNames"] = shape_names
        slide_info["sampleTexts"] = text_elements_summary[:5]
        slides_data.append(slide_info)

    font_list = list(fonts_found)
    color_list = list(colors_found)

    blueprint = {
        "templateName": os.path.basename(file_path),
        "totalSlides": len(prs.slides),
        "dimensions": {
            "widthInches": slide_width_in,
            "heightInches": slide_height_in,
            "aspectRatio": f"{slide_width_in}:{slide_height_in}"
        },
        "brandTheme": {
            "primaryColor": color_list[0] if color_list else "071E3D",
            "secondaryColor": color_list[1] if len(color_list) > 1 else "1E293B",
            "accentColor": color_list[2] if len(color_list) > 2 else "38B6FF",
            "fontFamily": font_list[0] if font_list else "Poppins",
            "detectedFonts": font_list,
            "detectedColors": color_list
        },
        "slides": slides_data
    }

    return blueprint

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python pptxDeepExtractor.py <pptx_file_path> [output_json_path]")
        sys.exit(1)

    pptx_path = sys.argv[1]
    blueprint_data = extract_deep_pptx_metadata(pptx_path)

    if len(sys.argv) >= 3:
        out_json_path = sys.argv[2]
        with open(out_json_path, "w", encoding="utf-8") as f:
            json.dump(blueprint_data, f, indent=2, ensure_ascii=False)
        print(f"Blueprint JSON saved to: {out_json_path}")
    else:
        print(json.dumps(blueprint_data, indent=2, ensure_ascii=False))
