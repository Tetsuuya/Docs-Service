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
   - Perfect 6-slide deck structure:
     Slide 1: Title Cover (slideIndex 1)
     Slide 2: Agenda Overview (slideIndex 2)
     Slide 3: Section 01 Cover (slideIndex 3)
     Slide 4: Section 01 Content Card (slideIndex 4 or 5)
     Slide 5: Section 02 Cover (slideIndex 28)
     Slide 6: Section 02 Content Card (slideIndex 38 or 40)
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
7. IMAGE PROMPTS FOR BACKGROUND PHOTOS: Generate hyper-specific, high-quality image prompts for slides with image slots.
   - ALWAYS include the main topic ("${userPrompt}", "chinese culture", "traditions", etc.) in EVERY imagePrompt.
   - ALWAYS include background aesthetic styling terms: "dark navy blue background, sleek modern studio lighting, cinematic aesthetic, high resolution photograph, matching presentation background".

OUTPUT FORMAT (strict JSON, no markdown wrappers):
{
  "presentationTitle": "Full Presentation Title in English",
  "topic": "${userPrompt}",
  "selectedSlides": [
    {
      "slideIndex": 1,
      "layoutCategory": "title_slide",
      "fillContent": {
        "TextBox 3": "CHINESE CULTURAL HERITAGE",
        "TextBox 4": "Traditions, Innovations, and Global Impact"
      },
      "imagePrompt": "A serene Chinese traditional architectural courtyard at night, dark navy blue background, sleek lighting, cinematic high resolution photograph",
      "speakerNotes": "Introduce Chinese culture, history, and global influence."
    }
  ]
}
`;

  const fillPlan = await callGemini(systemPrompt, 16384);

  if (!fillPlan.selectedSlides || !Array.isArray(fillPlan.selectedSlides)) {
    throw new Error('Gemini returned an invalid fill plan — missing selectedSlides array');
  }

  logger.info(`Phase 3 Complete → "${fillPlan.presentationTitle}" — ${fillPlan.selectedSlides.length} slides planned`);
  return fillPlan;
};
