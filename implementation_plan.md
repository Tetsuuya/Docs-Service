# Implementation Plan: PowerPoint (PPTX) Generation & Template Reuse System

This plan outlines the multi-phase implementation for adding **PowerPoint (.pptx) generation** to `Docs-Service`. It covers both **Scratch Mode** (creating slides from scratch) and **Template Analysis & Fill Mode** (scanning existing master presentation decks, saving layout metadata, and filling them with new AI content).

All code changes will be strictly isolated inside `src/services/pptx/` to prevent any conflicts with existing DOCX, PDF, or XLSX services.

---

## Layman's Executive Summary (What This Entire System Does)

Imagine your client has a beautiful master PowerPoint or Canva PDF template with 43 pre-designed slides (headers, 3-column feature cards, quote slides, bullet points).

Instead of manually copying and editing slides every time:
1. **Scanner:** The AI "scans" the master template and remembers the exact layout, text boxes, and styling of every single slide.
2. **Brain (Planner):** When asked to make a guide (e.g. *"Create a 10-slide guide on Marketing Tricks"*), the AI picks the best matching master slides for each part of the guide.
3. **Builder:** The AI replaces the dummy filler text and images in those slides with brand-new, topic-specific text while keeping the exact design, fonts, and colors of the original template intact.
4. **Scratch Fallback:** If no template is provided, the AI can invent its own presentation layout and build a PowerPoint file from scratch.

---

## User Review Required

> [!IMPORTANT]
> - **Dependency Check:** PPTX generation will utilize `pptxgenjs` (for scratch generation) and XML/zip manipulation (`jszip` / `xml2js`) for template parsing and replacement.
> - **Template Formats:** Template parsing will natively support `.pptx` master files. PDF templates (like the Canva export) can be processed either by converting slide pages to background images or extracting text overlay regions.
> - **Zero Conflict Guarantee:** All files sit inside `src/services/pptx/`, ensuring no changes or risks to existing `src/services/docx/` rules.

---

## Phased Implementation Roadmap

### Phase 1: PPTX Core Generator (Scratch Mode)

#### **What it does in Layman's Terms:**
You give the API a topic (e.g., *"5 Steps to Scale a Business"*). The AI plans the slide deck structure, writes titles and content for each slide, and builds a clean PowerPoint file `.pptx` from scratch.

#### **Technical Details:**
- **Location:** `src/services/pptx/scratch/`
- **New Files:**
  - `src/services/pptx/scratch/scratchPptxService.js`: Calls AI model to generate slide outline JSON (title, layout type, text boxes, bullet points, speaker notes).
  - `src/services/pptx/scratch/scratchPptxBuilder.js`: Uses `pptxgenjs` to programmatically build slide objects and export an `.pptx` Buffer.

---

### Phase 2: PPTX Template Inspector & Metadata Extractor

#### **What it does in Layman's Terms:**
This is the "Scanner." You upload a master PowerPoint file (`.pptx`). The system opens the file, looks at every slide, and creates a "blueprint map" of where the title, subtitles, body text paragraphs, icon boxes, and image placeholders are located.

#### **Technical Details:**
- **Location:** `src/services/pptx/template/`
- **New Files:**
  - `src/services/pptx/template/pptxTemplateParser.js`: Parses `.pptx` files by inspecting slide XML structure (`jszip`). Extracts:
    - Slide index & layout type (e.g., Title Slide, 2-Column Comparison, Grid Card, Quote).
    - Placeholder IDs, bounding boxes, text styles (font size, color, alignment).
    - Background styles & shapes.
  - Returns a clean JSON metadata map representing the template's structure.

---

### Phase 3: AI Presentation Planner & Smart Slide Selection

#### **What it does in Layman's Terms:**
This is the "Brain." When a user asks for a presentation based on a master template, the AI compares the user's request against the template's blueprint map. It selects which template slides to use for each step of the presentation and writes the exact text to put into each text box.

#### **Technical Details:**
- **Location:** `src/services/pptx/fill/`
- **New Files:**
  - `src/services/pptx/fill/pptxPlanGenerator.js`: Sends template metadata map + user prompt to the AI model.
  - AI returns a detailed plan:
    - Slide 1: Use Template Slide #1 (Title) -> Replace Header with "AI Marketing Guide".
    - Slide 2: Use Template Slide #5 (3-Column Layout) -> Replace Column 1, 2, 3 text.
    - Slide 3: Use Template Slide #12 (Call to Action) -> Replace CTA text.

---

### Phase 4: PPTX Template Fill & Assembly Builder

#### **What it does in Layman's Terms:**
This is the "Assembly Line." It takes the chosen template slides from the master presentation, inserts the AI-written content into the text boxes, replaces placeholder images if needed, and packages everything into a final, ready-to-download `.pptx` file.

#### **Technical Details:**
- **Location:** `src/services/pptx/fill/`
- **New Files:**
  - `src/services/pptx/fill/fillPptxBuilder.js`: 
    - Copies selected slide XMLs from master PPTX template.
    - Swaps text nodes inside `<a:t>` tags for matching placeholders.
    - Re-bundles the zip archive to output a clean, valid `.pptx` presentation buffer.

---

### Phase 5: Controller & Route Integration

#### **What it does in Layman's Terms:**
Wires the new PowerPoint features into your main backend API so your frontend/clients can request PowerPoint files using the exact same API endpoints they currently use for Word documents.

#### **Technical Details:**
- **Files to Modify:**
  - `src/controllers/documentController.js`: Route `format === 'pptx'` to PPTX scratch or fill handlers.
  - `src/routes/documentRoutes.js`: Ensure file uploads (master templates) work seamlessly for PPTX.

---

## Verification & Testing Plan

### Automated / API Tests
- `temp/test_pptx_scratch.js`: Test generating a 5-slide PPTX from a prompt.
- `temp/test_pptx_fill.js`: Test parsing a master `.pptx` template file, supplying a prompt, and validating that the output PPTX retains template styling with new text.

### Manual Verification
- Open generated `.pptx` files in Microsoft PowerPoint / Google Slides / Keynote to verify typography, slide layouts, shapes, and alignment.
