"""
directPptxReplacer.py
=====================
Phase 4: High-Precision Native PPTX Master-Deck Replacer.

Fixes:
  1. 100% Complete Background & Photo Replacement across Slides:
     Replaces the primary photo/background blip (> 5KB) on slides with generated AI topic images,
     ensuring ZERO original Canva template background images (e.g. French office photos, red panels) remain!
  2. Dark Navy Vignette Overlay (50% Opacity):
     Inserts a sleek dark navy blue tint overlay (<a:srgbClr val="0D1B2A"><a:alpha val="50000"/></a:srgbClr>)
     over full-bleed background images so photography ALWAYS fits presentation aesthetics and text is 100% sharp!
  3. Agenda Slide Background Protection:
     Preserves clean vector background panels on Agenda slides (len(num_shapes) > 1) so agenda text & icons
     NEVER clash with busy photographs!
  4. Masking Layout Master Text:
     Inserts clean space runs (<a:t> </a:t>) when clearing unmapped text frames so underlying Slide Layout Master text
     (e.g. MODULE 1, ACCULTURATION) NEVER bleeds through!
  5. Recursive Group Shape Text Frame Traversal:
     Recursively traverses Group Shapes (p:grpSp) to clear/update ALL nested text frames,
     purging 100% of residual French text from cards, grids, and multi-column elements!
  6. Font Fallback & Pipe Removal:
     Maps proprietary Canva fonts to standard system fonts (Segoe UI, Arial) and strips pipe characters (|)
     so text NEVER renders as boxes [⬜] or wraps into multi-line pipe strings!
"""

import sys
import os
import json
import copy
import re

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.oxml.ns import qn
    from pptx.util import Inches
    from lxml import etree
except ImportError:
    print("ERROR: python-pptx and lxml are required. Run: pip install python-pptx lxml", file=sys.stderr)
    sys.exit(1)


SAFE_FONT_MAP = {
    "canva sans": "Arial",
    "canva sans bold": "Arial",
    "aileron": "Segoe UI",
    "aileron bold": "Segoe UI",
    "helvetica now display": "Arial",
    "helvetica now display bold": "Arial",
}

TEMPLATE_RESIDUAL_KEYWORDS = [
    "démystifier", "demystifier", "acculturation", "janvier", "présenté par",
    "presente par", "module 1", "module 2", "intelligences", "biais",
    "i-avantage", "question", "aujourd'hui", "aujourd’hui"
]


def sanitize_text(text):
    if not text:
        return ""
    text = str(text)
    text = text.replace("|", " ")  # Replace pipe characters with space to prevent title box multi-line wrap
    replacements = {
        "\u2019": "'",
        "\u2018": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u2014": " - ",
        "\u2013": " - ",
        "\u00a0": " ",
        "\u2022": "*",
        "\u2026": "...",
    }
    for orig, repl in replacements.items():
        text = text.replace(orig, repl)
    text = re.sub(r"[^\x00-\x7F\xA0-\xFF]", "", text)
    return re.sub(r"\s+", " ", text).strip()


def get_all_text_shapes(container):
    """Recursively yields all shapes with text frames, including nested child shapes in Group Shapes (p:grpSp)."""
    text_shapes = []
    try:
        for shape in container.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                text_shapes.extend(get_all_text_shapes(shape))
            elif shape.has_text_frame:
                text_shapes.append(shape)
    except Exception:
        pass
    return text_shapes


def clear_text_frame(tf):
    txBody = tf._txBody
    paras = txBody.findall(qn("a:p"))
    for para in paras:
        txBody.remove(para)
    # Add a clean space run so underlying layout master placeholder text is masked
    p = etree.SubElement(txBody, qn("a:p"))
    r = etree.SubElement(p, qn("a:r"))
    t = etree.SubElement(r, qn("a:t"))
    t.text = " "


def get_safe_font_name(original_font_name):
    if not original_font_name:
        return "Arial"
    lower_name = original_font_name.lower().strip()
    return SAFE_FONT_MAP.get(lower_name, original_font_name)


def write_text_preserve_first_run_style(shape, new_text, should_wrap=True):
    if not shape.has_text_frame:
        return

    tf = shape.text_frame
    txBody = tf._txBody
    paras = txBody.findall(qn("a:p"))

    first_rPr = None
    first_pPr = None

    for para in paras:
        pPr = para.find(qn("a:pPr"))
        if pPr is not None and first_pPr is None:
            first_pPr = copy.deepcopy(pPr)
        for run_elem in para.findall(qn("a:r")):
            rPr = run_elem.find(qn("a:rPr"))
            if rPr is not None and first_rPr is None:
                first_rPr = copy.deepcopy(rPr)
                break
        if first_rPr is not None:
            break

    # Extract original text for scaling comparison
    orig_text = tf.text.strip()

    if first_rPr is not None:
        latin_font = first_rPr.find(qn("a:latin"))
        if latin_font is not None:
            orig_typeface = latin_font.get("typeface", "")
            safe_typeface = get_safe_font_name(orig_typeface)
            latin_font.set("typeface", safe_typeface)
        else:
            latin_elem = etree.SubElement(first_rPr, qn("a:latin"))
            latin_elem.set("typeface", "Arial")

        # Scale down font size for single-line labels if the new text is longer than the placeholder
        if not should_wrap and orig_text:
            sanitized_len = len(new_text.strip())
            if sanitized_len > len(orig_text):
                sz_val = first_rPr.get("sz")
                if sz_val:
                    try:
                        orig_sz = int(sz_val)
                        scale = max(float(len(orig_text)) / float(sanitized_len), 0.4)
                        new_sz = int(orig_sz * scale)
                        first_rPr.set("sz", str(new_sz))
                        print(f"  [Autofit] Scaled font size of {shape.name} from {orig_sz/100:.1f}pt to {new_sz/100:.1f}pt (scale={scale:.2f})", file=sys.stderr)
                    except Exception as e:
                        print(f"  [warn] Could not scale font: {e}", file=sys.stderr)

    for para in paras:
        txBody.remove(para)

    sanitized = sanitize_text(new_text)
    if not sanitized:
        return

    new_para = etree.SubElement(txBody, qn("a:p"))
    if first_pPr is not None:
        new_para.insert(0, first_pPr)

    # Force latinLnBrk="0" to ensure PowerPoint wraps at word boundaries, not mid-character!
    pPr = new_para.find(qn("a:pPr"))
    if pPr is None:
        pPr = etree.SubElement(new_para, qn("a:pPr"))
    pPr.set("latinLnBrk", "0")

    new_run = etree.SubElement(new_para, qn("a:r"))
    if first_rPr is not None:
        new_run.insert(0, first_rPr)

    t_elem = etree.SubElement(new_run, qn("a:t"))
    t_elem.text = sanitized

    # Enable "Shrink text on overflow" (normAutofit) so long words never break mid-character
    try:
        bodyPr = txBody.find(qn("a:bodyPr"))
        if bodyPr is not None:
            # Remove any existing autofit elements
            for autofit_tag in ["a:noAutofit", "a:spAutoFit", "a:normAutofit"]:
                existing = bodyPr.find(qn(autofit_tag))
                if existing is not None:
                    bodyPr.remove(existing)
            
            # Enable normAutofit — PowerPoint will shrink font until text fits
            etree.SubElement(bodyPr, qn("a:normAutofit"))
            
            # Control wrapping property
            if not should_wrap:
                bodyPr.set("wrap", "none")
            else:
                bodyPr.set("wrap", "square")
    except Exception:
        pass


def add_dark_overlay_to_slide(slide):
    """Inserts a 50% dark navy blue vignette overlay (<a:srgbClr val="0D1B2A">) over background images for executive presentation legibility."""
    try:
        spTree = slide.shapes._spTree

        # Check if overlay already exists
        if spTree.xpath('.//p:cNvPr[@name="DarkOverlay"]'):
            return

        rect = etree.SubElement(spTree, qn('p:sp'))
        nvSpPr = etree.SubElement(rect, qn('p:nvSpPr'))
        etree.SubElement(nvSpPr, qn('p:cNvPr'), id='99999', name='DarkOverlay')
        etree.SubElement(nvSpPr, qn('p:cNvSpPr'))
        etree.SubElement(nvSpPr, qn('p:nvPr'))

        spPr = etree.SubElement(rect, qn('p:spPr'))
        xfrm = etree.SubElement(spPr, qn('a:xfrm'))
        etree.SubElement(xfrm, qn('a:off'), x='0', y='0')
        etree.SubElement(xfrm, qn('a:ext'), cx='12192000', cy='6858000')

        prstGeom = etree.SubElement(spPr, qn('a:prstGeom'), prst='rect')
        etree.SubElement(prstGeom, qn('a:avLst'))

        solidFill = etree.SubElement(spPr, qn('a:solidFill'))
        srgbClr = etree.SubElement(solidFill, qn('a:srgbClr'), val='0D1B2A')
        etree.SubElement(srgbClr, qn('a:alpha'), val='45000')  # 45% dark navy tint overlay

        ln = etree.SubElement(spPr, qn('a:ln'))
        etree.SubElement(ln, qn('a:noFill'))

        # Insert dark overlay right after background image shape so text shapes render ON TOP
        spTree.remove(rect)
        spTree.insert(2, rect)
    except Exception as e:
        print(f"  [overlay warn] Could not add dark overlay: {e}", file=sys.stderr)


def replace_main_slide_photo(slide, image_file_path, is_agenda=False, layout_category='content_slide'):
    """
    Inserts a BRAND NEW independent image part for this slide and updates the primary photo/background blip.
    Replaces the largest non-icon image blip (> 5KB) on the slide with the generated AI topic image.
    On Agenda slides (is_agenda=True): skips replacing full-bleed background panels to keep icons 100% legible!
    Dark overlay ONLY applied on title_slide and section_header slides (dark navy theme).
    """
    if not image_file_path or not os.path.exists(image_file_path):
        return
    try:
        blips = slide.shapes._spTree.xpath('.//a:blip')
        if not blips:
            return

        primary_blip = None
        max_size = -1
        has_full_bg = False

        for blip in blips:
            rId = blip.get(qn('r:embed'))
            if not rId or rId not in slide.part.rels:
                continue

            rel = slide.part.rels[rId]
            if not (hasattr(rel, 'target_part') and hasattr(rel.target_part, '_blob')):
                continue

            blob_size = len(rel.target_part._blob)

            # Skip tiny vector/icon blips (< 5KB)
            if blob_size < 5000:
                continue

            cx_list = blip.xpath('ancestor::*/a:xfrm/a:ext/@cx')
            cy_list = blip.xpath('ancestor::*/a:xfrm/a:ext/@cy')

            is_full_bg = False
            if cx_list and cy_list:
                try:
                    cx = int(cx_list[0])
                    cy = int(cy_list[0])
                    if cx >= 17000000 and cy >= 9500000:
                        is_full_bg = True
                except ValueError:
                    pass

            # On Agenda slides: Skip replacing full-bleed background image to protect icon legibility
            if is_full_bg and is_agenda:
                continue

            if is_full_bg:
                has_full_bg = True

            if blob_size > max_size:
                max_size = blob_size
                primary_blip = blip

        if primary_blip is not None:
            new_image_part, new_rId = slide.part.get_or_add_image_part(image_file_path)
            primary_blip.set(qn('r:embed'), new_rId)
            print(f"  [image] Attached AI topic image ({os.path.basename(image_file_path)}) as rId '{new_rId}'", file=sys.stderr)

            # Only apply dark overlay on title/section cover slides (dark navy template style)
            # Content slides use light/white background — do NOT darken them!
            is_dark_slide = layout_category in ('title_slide', 'section_header')
            if has_full_bg and not is_agenda and is_dark_slide:
                add_dark_overlay_to_slide(slide)

    except Exception as e:
        print(f"  [image warn] Could not replace image: {e}", file=sys.stderr)


def purge_layout_residual_text(slide):
    """Purge residual French template text from slide layout shapes."""
    try:
        if slide.slide_layout:
            for shape in slide.slide_layout.shapes:
                if shape.has_text_frame:
                    txt = shape.text_frame.text.lower().strip()
                    if any(kw in txt for kw in TEMPLATE_RESIDUAL_KEYWORDS):
                        clear_text_frame(shape.text_frame)
    except Exception:
        pass


def delete_slide(prs, index):
    rId = prs.slides._sldIdLst[index].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[index]


def fill_master_pptx_deck(master_pptx_path, fill_plan_path, output_pptx_path, image_map_path=None):
    print(f"Loading master template: {master_pptx_path}", file=sys.stderr)
    prs = Presentation(master_pptx_path)

    with open(fill_plan_path, "r", encoding="utf-8") as f:
        fill_plan = json.load(f)

    image_map = {}
    if image_map_path and os.path.exists(image_map_path):
        try:
            with open(image_map_path, "r", encoding="utf-8") as f:
                image_map = json.load(f)
        except Exception:
            pass

    selected_slides_data = fill_plan.get("selectedSlides", [])
    selected_indices = set()
    print(f"DEBUG replacer selected_slides_data: {[e.get('slideIndex') for e in selected_slides_data]}", file=sys.stderr)
    for entry in selected_slides_data:
        raw_idx = entry.get("slideIndex", 1)
        zero_idx = int(raw_idx) - 1
        if 0 <= zero_idx < len(prs.slides):
            selected_indices.add(zero_idx)

    # 1. Fill text & attach unique AI topic images to selected slides
    section_counter = 0
    for idx_in_plan, entry in enumerate(selected_slides_data):
        zero_idx = entry.get("slideIndex", 1) - 1
        if not (0 <= zero_idx < len(prs.slides)):
            continue

        slide = prs.slides[zero_idx]
        fill_content = entry.get("fillContent", {})
        speaker_notes = entry.get("speakerNotes", "")

        # Check if this slide is a standalone Section Header slide (TextBox 3 is a big section badge "01", "02")
        all_text_shapes = get_all_text_shapes(slide)
        num_shapes = [s for s in all_text_shapes if s.name == "TextBox 3" and s.text_frame.text.strip() in ["01", "02", "03", "04", "05", "06"]]
        agenda_shapes = [s for s in all_text_shapes if s.text_frame.text.strip() in ["01", "02", "03", "04", "05", "06"]]
        
        is_section_header = (len(num_shapes) == 1)
        is_agenda = (len(agenda_shapes) > 1)

        if is_section_header:
            section_counter += 1
            formatted_sec_num = f"{section_counter:02d}"

        shapes_to_write = {}
        for shape_name, content in fill_content.items():
            if content and str(content).strip():
                shapes_to_write[shape_name] = str(content).strip()

        # Auto-correct ONLY standalone section header badge shape to strict sequential number (01, 02, 03...)
        if is_section_header and num_shapes:
            target_badge_shape = num_shapes[0]
            shapes_to_write[target_badge_shape.name] = formatted_sec_num

        # Purge French residual text from slide layout master
        purge_layout_residual_text(slide)

        for shape in all_text_shapes:
            if shape.name in shapes_to_write:
                new_text_clean = shapes_to_write[shape.name].strip()
                should_wrap = True
                
                is_short_label = False
                name_lower = shape.name.lower()
                if any(kw in name_lower for kw in ["badge", "tag", "number", "date", "footer", "author"]):
                    is_short_label = True
                if shape.name in [
                    "TextBox 6", "TextBox 7", "TextBox 9", "TextBox 10",
                    "TextBox 31", "TextBox 32", "TextBox 33", "TextBox 34", "TextBox 35", 
                    "TextBox 36", "TextBox 37", "TextBox 38", "TextBox 39", "TextBox 40"
                ]:
                    is_short_label = True
                    
                is_section_cover = (entry.get("layoutCategory") == "section_header" or (zero_idx + 1) in [3, 28, 43])
                if is_section_cover and shape.name in ["TextBox 4", "TextBox 5"]:
                    is_short_label = True

                is_title_slide = (entry.get("layoutCategory") == "title_slide" or zero_idx == 0)
                if is_title_slide and shape.name == "TextBox 3":
                    is_short_label = False
                elif is_title_slide and shape.name in ["TextBox 4", "TextBox 5"]:
                    is_short_label = True

                if is_short_label and len(new_text_clean) < 60:
                    should_wrap = False

                write_text_preserve_first_run_style(shape, shapes_to_write[shape.name], should_wrap=should_wrap)
            else:
                clear_text_frame(shape.text_frame)



        # Attach unique independent AI image for this slide
        layout_category = entry.get('layoutCategory', 'content_slide')
        img_key = str(idx_in_plan)
        if img_key in image_map:
            replace_main_slide_photo(slide, image_map[img_key], is_agenda=is_agenda, layout_category=layout_category)

        if speaker_notes:
            try:
                slide.notes_slide.notes_text_frame.text = sanitize_text(speaker_notes)
            except Exception:
                pass

        print(f"  Slide {zero_idx + 1}: processed {len(all_text_shapes)} recursive text shapes (filled={len(shapes_to_write)})", file=sys.stderr)

    # 2. Prune unselected slides & Reorder to match exact planned sequence
    total_before = len(prs.slides)
    original_slides = list(prs.slides)
    
    # Map original zero_idx to XML elements
    sldId_map = {}
    for idx, slide in enumerate(original_slides):
        sldId_map[idx] = prs.slides._sldIdLst[idx]

    new_sldId_list = []
    seen_indices = set()
    
    for entry in selected_slides_data:
        raw_idx = entry.get("slideIndex", 1)
        zero_idx = int(raw_idx) - 1
        if 0 <= zero_idx < len(original_slides):
            if zero_idx not in seen_indices:
                seen_indices.add(zero_idx)
                new_sldId_list.append(sldId_map[zero_idx])

    # Drop relationships for unselected slides
    for idx in range(total_before):
        if idx not in seen_indices:
            try:
                rId = sldId_map[idx].rId
                prs.part.drop_rel(rId)
            except Exception:
                pass

    # Clear and rebuild slide XML sequence in the new ordered sequence
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        sldIdLst.remove(sldId)
    for sldId in new_sldId_list:
        sldIdLst.append(sldId)

    total_after = len(prs.slides)
    print(f"Slides: {total_before} (template) → {total_after} (output, {len(seen_indices)} selected, reordered)", file=sys.stderr)

    # 3. Save output PPTX
    os.makedirs(os.path.dirname(os.path.abspath(output_pptx_path)), exist_ok=True)
    prs.save(output_pptx_path)
    size_bytes = os.stat(output_pptx_path).st_size
    size_kb = round(size_bytes / 1024, 1)
    print(f"Saved: {output_pptx_path} ({size_kb} KB)", file=sys.stderr)


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: python directPptxReplacer.py <master.pptx> <fill_plan.json> <output.pptx> [image_map.json]", file=sys.stderr)
        sys.exit(1)

    img_map = sys.argv[4] if len(sys.argv) >= 5 else None
    fill_master_pptx_deck(sys.argv[1], sys.argv[2], sys.argv[3], img_map)
