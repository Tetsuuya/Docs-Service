"""
pptxDeepExtractor.py
====================
Extracts a complete structural blueprint from a PPTX master template.
"""

import sys
import os
import json

try:
    import pptx
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
except ImportError:
    print("ERROR: python-pptx is not installed. Run: pip install python-pptx", file=sys.stderr)
    sys.exit(1)

EMU = 914400.0


def emu_to_in(emu_val):
    try:
        return round(emu_val / EMU, 3)
    except Exception:
        return 0.0


def rgb_to_hex(rgb):
    try:
        if isinstance(rgb, RGBColor):
            return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    except Exception:
        pass
    return None


def get_shape_fill_hex(shape):
    try:
        fill = shape.fill
        if fill and fill.type == 1:
            rgb = fill.fore_color.rgb
            return rgb_to_hex(rgb)
    except Exception:
        pass
    return None


def get_alignment_label(para_alignment):
    mapping = {
        PP_ALIGN.LEFT: "LEFT",
        PP_ALIGN.CENTER: "CENTER",
        PP_ALIGN.RIGHT: "RIGHT",
        PP_ALIGN.JUSTIFY: "JUSTIFY",
    }
    return mapping.get(para_alignment, "LEFT")


def extract_text_frame(tf):
    paragraphs = []
    all_text_parts = []

    for para in tf.paragraphs:
        p_text = para.text.strip()
        if not p_text:
            continue

        runs_info = []
        for run in para.runs:
            r_info = {}
            try:
                if run.font.name:
                    r_info["fontName"] = run.font.name
            except Exception:
                pass
            try:
                if run.font.size:
                    r_info["fontSizePt"] = round(run.font.size / 12700.0, 1)
            except Exception:
                pass
            try:
                if run.font.bold:
                    r_info["bold"] = True
            except Exception:
                pass
            try:
                rgb = run.font.color.rgb
                if rgb:
                    r_info["color"] = rgb_to_hex(rgb)
            except Exception:
                pass
            if run.text.strip():
                r_info["text"] = run.text
            if r_info:
                runs_info.append(r_info)

        paragraphs.append({
            "text": p_text,
            "alignment": get_alignment_label(para.alignment),
            "runs": runs_info
        })
        all_text_parts.append(p_text)

    return paragraphs, " | ".join(all_text_parts)


def classify_layout(slide_idx, elements, has_blips=False):
    if slide_idx == 0:
        return "title_slide"

    text_shapes = [e for e in elements if e.get("hasText")]
    n_text = len(text_shapes)

    # Check for standalone section number badge (01, 02, 03...)
    all_texts = [e.get("fullText", "").strip() for e in text_shapes if e.get("fullText")]
    badge_found = any(t in ["01", "02", "03", "04", "05", "06", "07", "08", "09"] for t in all_texts)

    if badge_found and n_text <= 4:
        return "section_header"

    if has_blips and n_text >= 2:
        return "split_image_text"
    if n_text >= 4:
        return "multi_column_content"
    if n_text == 3:
        return "three_column_cards"
    if n_text == 2:
        return "two_column_layout"
    if has_blips and n_text == 1:
        return "full_image_caption"
    if n_text == 1:
        all_text = text_shapes[0].get("fullText", "")
        if len(all_text) > 120:
            return "content_body"
        return "section_header"
    if has_blips:
        return "full_bleed_image"
    return "content_slide"


def describe_slide(layout_category, elements, has_blips=False):
    text_samples = [e.get("fullText", "")[:60] for e in elements if e.get("hasText") and e.get("fullText")]
    desc_parts = []
    if text_samples:
        desc_parts.append(f"Text slots: {'; '.join(text_samples[:3])}")
    if has_blips:
        desc_parts.append("Has main photo image slot")
    desc_parts.insert(0, f"Layout: {layout_category.replace('_', ' ').title()}")
    return ". ".join(desc_parts) + "."


def extract_deep_pptx_metadata(pptx_path):
    prs = pptx.Presentation(pptx_path)

    slide_w = emu_to_in(prs.slide_width)
    slide_h = emu_to_in(prs.slide_height)

    all_fonts = set()
    all_colors = set()
    slides_data = []

    for slide_idx, slide in enumerate(prs.slides):
        layout_name = ""
        try:
            layout_name = slide.slide_layout.name or ""
        except Exception:
            pass

        elements = []

        # Check for blips (image fills / picture elements) in slide XML
        blips = slide.shapes._spTree.xpath('.//a:blip')
        has_blips = len(blips) > 0

        for shape in slide.shapes:
            elem = {
                "shapeId": shape.shape_id,
                "shapeName": shape.name,
                "shapeType": str(shape.shape_type),
                "position": {
                    "left": emu_to_in(shape.left),
                    "top": emu_to_in(shape.top),
                    "width": emu_to_in(shape.width),
                    "height": emu_to_in(shape.height),
                },
                "hasText": False,
                "isImage": False,
            }

            fill_hex = get_shape_fill_hex(shape)
            if fill_hex:
                elem["fillColor"] = fill_hex
                all_colors.add(fill_hex)

            if shape.has_text_frame:
                try:
                    paragraphs, full_text = extract_text_frame(shape.text_frame)
                    if full_text:
                        elem["hasText"] = True
                        elem["paragraphs"] = paragraphs
                        elem["fullText"] = full_text[:200]

                        for para in paragraphs:
                            for run in para.get("runs", []):
                                if run.get("fontName"):
                                    all_fonts.add(run["fontName"])
                                if run.get("color"):
                                    all_colors.add(run["color"])
                except Exception:
                    pass

            elements.append(elem)

        layout_category = classify_layout(slide_idx, elements, has_blips)
        description = describe_slide(layout_category, elements, has_blips)

        shape_names = [e["shapeName"] for e in elements]
        text_shape_map = {
            e["shapeName"]: e.get("fullText", "")[:80]
            for e in elements if e.get("hasText")
        }
        image_shape_names = ["Main Photo Placeholder"] if has_blips else []

        slides_data.append({
            "slideIndex": slide_idx + 1,
            "layoutName": layout_name,
            "layoutCategory": layout_category,
            "description": description,
            "shapeNames": shape_names,
            "textShapes": text_shape_map,
            "imageShapes": image_shape_names,
            "elements": elements,
        })

    font_list = sorted(all_fonts)
    color_list = sorted(all_colors)

    def is_dark_enough(hex_val):
        try:
            h = hex_val.lstrip("#")
            r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
            brightness = (r * 299 + g * 587 + b * 114) / 1000
            return brightness < 220
        except Exception:
            return False

    dark_colors = [c for c in color_list if is_dark_enough(c)]

    brand_theme = {
        "primaryColor": dark_colors[0].lstrip("#") if dark_colors else "071E3D",
        "secondaryColor": dark_colors[1].lstrip("#") if len(dark_colors) > 1 else "1E293B",
        "accentColor": dark_colors[2].lstrip("#") if len(dark_colors) > 2 else "38B6FF",
        "fontFamily": font_list[0] if font_list else "Calibri",
        "detectedFonts": font_list,
        "detectedColors": color_list,
    }

    blueprint = {
        "templateName": os.path.basename(pptx_path),
        "totalSlides": len(prs.slides),
        "dimensions": {
            "widthInches": slide_w,
            "heightInches": slide_h,
        },
        "brandTheme": brand_theme,
        "slides": slides_data,
    }

    return blueprint


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python pptxDeepExtractor.py <pptx_path> [output_json_path]", file=sys.stderr)
        sys.exit(1)

    pptx_path = sys.argv[1]
    blueprint = extract_deep_pptx_metadata(pptx_path)

    if len(sys.argv) >= 3:
        out_path = sys.argv[2]
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(blueprint, f, indent=2, ensure_ascii=False)
        print(f"Blueprint saved to: {out_path}", file=sys.stderr)
    else:
        print(json.dumps(blueprint, indent=2, ensure_ascii=False))
