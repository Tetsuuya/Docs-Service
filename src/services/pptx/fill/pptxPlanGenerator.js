import { GoogleGenerativeAI } from '@google/generative-ai';
import { config } from '../../../config/env.js';
import { logger } from '../../../utils/logger.js';

const genAI = new GoogleGenerativeAI(config.geminiApiKey);

const candidateModels = [
  config.geminiModel,
  'gemini-flash-lite-latest',
  'gemini-2.0-flash-lite-preview-02-05',
  'gemini-1.5-flash-latest'
];

async function callGemini(promptInput, maxTokens = 16384) {
  let lastError = null;

  for (const modelName of candidateModels) {
    if (!modelName) continue;
    try {
      const model = genAI.getGenerativeModel({
        model: modelName,
        generationConfig: {
          temperature: 0.2,
          maxOutputTokens: maxTokens,
          responseMimeType: 'application/json'
        }
      });

      const result = await model.generateContent(promptInput);
      const responseText = result.response.text();
      return JSON.parse(responseText);
    } catch (err) {
      logger.warn(`PPTX Plan Generator — model "${modelName}" failed: ${err.message}`);
      lastError = err;
    }
  }
  throw lastError || new Error('Gemini API calls failed for PPTX plan generation');
}

function buildSlideCatalog(templateBlueprint) {
  const slides = templateBlueprint.slides || [];

  return slides.map(s => ({
    slideIndex: s.slideIndex,
    layoutCategory: s.layoutCategory || 'content_slide',
    description: s.description || '',
    textShapes: s.textShapes || {},     // { shapeName: sampleText }
    imageShapes: s.imageShapes || [],   // picture / image placeholder slots
  }));
}

const SECTION_HEADER_INDICES = new Set([3, 28, 43]);

/**
 * Deterministically enforces layout rules on Gemini's selected slides:
 * 1. Zero consecutive section cover slides (replaces consecutive section covers with content card layouts).
 * 2. Strict layout diversity across all slides.
 */
function sanitizeSelectedSlides(selectedSlides, slideCatalog) {
  if (!Array.isArray(selectedSlides) || selectedSlides.length === 0) return selectedSlides;

  // Find sample content card slides in template catalog to use as replacement
  const contentCardCatalogEntry = slideCatalog.find(s => s.layoutCategory === 'split_image_text' || s.layoutCategory === 'three_column_cards' || s.layoutCategory === 'multi_column_grid' || (s.slideIndex !== 1 && s.slideIndex !== 2 && !SECTION_HEADER_INDICES.has(s.slideIndex))) || { slideIndex: 5, layoutCategory: 'split_image_text' };

  const sanitized = [];

  for (let i = 0; i < selectedSlides.length; i++) {
    const current = selectedSlides[i];
    const prev = sanitized.length > 0 ? sanitized[sanitized.length - 1] : null;

    const isCurrentSection = SECTION_HEADER_INDICES.has(current.slideIndex) || current.layoutCategory === 'section_header';
    const isPrevSection = prev ? (SECTION_HEADER_INDICES.has(prev.slideIndex) || prev.layoutCategory === 'section_header') : false;

    if (isCurrentSection && isPrevSection) {
      logger.info(`  [Layout Guard] Converted consecutive Section Cover (slideIndex ${current.slideIndex}) to Content Card layout (slideIndex ${contentCardCatalogEntry.slideIndex})`);
      
      sanitized.push({
        ...current,
        slideIndex: contentCardCatalogEntry.slideIndex,
        layoutCategory: contentCardCatalogEntry.layoutCategory,
        fillContent: {
          "TextBox 6": current.fillContent?.["TextBox 4"] || current.fillContent?.["TextBox 3"] || "Key Takeaways",
          "TextBox 8": current.fillContent?.["TextBox 5"] || "Strategic overview and operational insights."
        }
      });
    } else {
      sanitized.push(current);
    }
  }

  return sanitized;
}

export const generatePptxFillPlan = async (userPrompt, templateBlueprint) => {
  if (!templateBlueprint || !templateBlueprint.slides?.length) {
    throw new Error('Template blueprint is empty or missing slides — cannot generate a fill plan.');
  }

  const templateName = templateBlueprint.templateName || 'Master Deck';
  const totalSlides = templateBlueprint.totalSlides || templateBlueprint.slides.length;
  const slideCatalog = buildSlideCatalog(templateBlueprint);

  logger.info(`Phase 3: Generating fill plan for "${userPrompt}" using template "${templateName}" (${totalSlides} slides available)`);

  const slideCountMatch = userPrompt.match(/\b(\d+)\s*(?:-| )?slide/i);
  const requestedSlideCount = slideCountMatch
    ? Math.min(Math.max(parseInt(slideCountMatch[1], 10), 3), Math.min(totalSlides, 20))
    : null;

  const slideCatalogJson = JSON.stringify(slideCatalog, null, 2);

  const systemPrompt = `
You are an executive presentation director and copywriter.
You are given a MASTER TEMPLATE BLUEPRINT containing layout placeholders (text slots & image slots).
Your task is to plan a complete PowerPoint presentation based on the user's prompt topic: "${userPrompt}".
You must replace ALL text placeholders with fresh, topic-specific 100% ENGLISH executive content.

USER REQUEST: "${userPrompt}"
${requestedSlideCount ? `REQUESTED SLIDE COUNT: ${requestedSlideCount} slides` : 'SLIDE COUNT: Choose 6–8 slides that best tell the story'}

MASTER TEMPLATE SLIDE PLACEHOLDER CATALOG:
${slideCatalogJson}

CRITICAL COPYWRITING & TEMPLATE STRUCTURAL RULES:
1. 100% ENGLISH ONLY: Write all presentation titles, subtitles, bullet points, and speaker notes in clean, professional English.
2. DO NOT COPY TEMPLATE FRENCH TEXT:
   - NEVER copy French template strings from the catalog (e.g. "MODULE 1", "ACCULTURATION", "PRÉSENTÉ PAR", "DÉMYSTIFIER", "BIAIS", "INTELLIGENCES").
   - Replace category headers with clean English (e.g. "SECTION 1", "OVERVIEW", "KEY TAKEAWAYS", "CLINICAL RESEARCH").
3. MAXIMUM 2 SECTION COVER SLIDES FOR SHORT DECKS (MANDATORY):
   - For presentations with 6 slides, include AT MOST 2 Section Cover slides (e.g. Section 01 and Section 02 only).
   - NEVER place two section cover slides (slides with layoutCategory: "section_header" like 01, 02, 03) right next to each other!
   - Every Section Cover slide MUST be followed by 1 or 2 CONTENT slides (e.g. Cards grid slides, Split Photo slides, Stat/Fact slides) before another section cover slide can appear!
4. STRICT SECTION NUMBER SEQUENCING (01, 02, 03...):
   - If you include section header slides (slides with numbers like "01", "02", "03"), select them in STRICT NUMERICAL ORDER!
   - Slide 3 is Section 01, Slide 28 is Section 02, Slide 43 is Section 03.
   - NEVER skip Section 02 to jump to 03! Always plan sections sequentially: Section 1 (01) -> Section 2 (02) -> Section 3 (03).
5. NO PIPE SEPARATORS OR LONG PREFIXES:
   - Do NOT use pipe characters "|" in titles (e.g. write "Cellular Biology Overview", NOT "MODULE 1 | DEMYSTIFYING | CELLULAR BIOLOGY").
   - Keep titles short (3–6 words max) so text NEVER overflows or overlaps adjacent shapes!
6. STACKED SHAPE DISCIPLINE: On slides with multiple title shapes (e.g. TextBox 3 and TextBox 4):
   - Put the main short headline in TextBox 3 (e.g. "CANCER BIOLOGY")
   - Put a short subtitle or LEAVE BLANK in TextBox 4 (do NOT repeat the headline in TextBox 4!)
7. TEMPLATE-MATCHED IMAGE PROMPTS PER SLIDE TYPE (CRITICAL):
    - The template has TWO distinct slide visual styles. You MUST match the correct image style to each slide type:

    STYLE A — Title Cover & Section Header slides (dark background slides):
    - These slides have a DARK NAVY BLUE full-bleed background with a glowing neon-lit photo clipped into a teardrop/lens shape on the RIGHT side.
    - imagePrompt format: "[specific topic subject], dramatic dark navy blue background, glowing electric blue neon rim lighting, photorealistic, cinematic studio shot, sharp focus, ultra high resolution, no text"
    - Example: "Modern skyscraper at night, dramatic dark navy blue background, glowing electric blue neon rim lighting, cinematic studio shot, ultra high resolution, no text"

    STYLE B — Content slides (light background slides):
    - These slides have a LIGHT/WHITE background. The generated image fills a DARK NAVY ROUNDED CARD PANEL on the RIGHT half of the slide.
    - imagePrompt format: "[specific topic subject], clean white studio background, soft blue accent lighting, professional product/concept shot, photorealistic, sharp focus, no text, no watermark"
    - Example: "Glowing blue neural network sphere, clean white studio background, soft blue accent lighting, professional concept shot, no text"

    MATCHING RULE:
    - layoutCategory = "title_slide" or "section_header" → USE STYLE A (dark navy neon)
    - layoutCategory = "split_image_text", "three_column_cards", "multi_column_grid", or any other content type → USE STYLE B (white studio)
    - NEVER use random stock photography or busy scenes. ALWAYS generate abstract/concept visuals matching the topic.
    - NEVER add text, logos, or watermarks in image prompts.

OUTPUT FORMAT (strict JSON, no markdown wrappers):
{
  "presentationTitle": "Full Presentation Title in English",
  "topic": "${userPrompt}",
  "selectedSlides": [
    {
      "slideIndex": 1,
      "layoutCategory": "title_slide",
      "fillContent": {
        "TextBox 3": "EXECUTIVE NETWORKING",
        "TextBox 4": "Building Strategic Professional Relationships"
      },
      "imagePrompt": "Corporate professionals handshake, dramatic dark navy blue background, glowing electric blue neon rim lighting, photorealistic, cinematic studio shot, sharp focus, ultra high resolution, no text",
      "speakerNotes": "Introduce key concepts of professional networking."
    },
    {
      "slideIndex": 5,
      "layoutCategory": "split_image_text",
      "fillContent": {
        "TextBox 6": "KEY STRATEGIES",
        "TextBox 8": "Insights and frameworks for building executive connections."
      },
      "imagePrompt": "Professional business network graph nodes, clean white studio background, soft blue accent lighting, professional concept shot, photorealistic, sharp focus, no text, no watermark",
      "speakerNotes": "Detail the core strategies for professional networking success."
    }
  ]
}
`;

  const fillPlan = await callGemini(systemPrompt, 16384);

  if (!fillPlan.selectedSlides || !Array.isArray(fillPlan.selectedSlides)) {
    throw new Error('Gemini returned an invalid fill plan — missing selectedSlides array');
  }

  // Deterministically enforce zero consecutive section cover slides
  fillPlan.selectedSlides = sanitizeSelectedSlides(fillPlan.selectedSlides, slideCatalog);

  logger.info(`Phase 3 Complete → "${fillPlan.presentationTitle}" — ${fillPlan.selectedSlides.length} slides planned`);
  return fillPlan;
};
