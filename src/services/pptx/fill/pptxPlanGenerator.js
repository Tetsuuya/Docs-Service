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

async function callGemini(promptInput, maxTokens = 8192) {
  let lastError = null;

  for (const modelName of candidateModels) {
    if (!modelName) continue;
    try {
      const model = genAI.getGenerativeModel({
        model: modelName,
        generationConfig: {
          temperature: 0.7,
          maxOutputTokens: maxTokens,
          responseMimeType: 'application/json'
        }
      });

      const result = await model.generateContent(promptInput);
      const responseText = result.response.text();
      return JSON.parse(responseText);
    } catch (err) {
      logger.warn(`PPTX Plan Generator Gemini call with model "${modelName}" failed: ${err.message}`);
      lastError = err;
    }
  }
  throw lastError || new Error('Gemini API calls failed for PPTX plan generation');
}

/**
 * Phase 3: AI Presentation Planner & Matcher
 * Matches user prompt topics in ANY language to master slide layouts in template blueprint
 * Generates exact text content for each slide container shape in ENGLISH by default.
 *
 * @param {string} userPrompt - Topic prompt requested by user (in any language)
 * @param {Object} templateBlueprint - Structured metadata blueprint extracted in Phase 2
 * @returns {Promise<Object>} - Structured AI Fill Plan JSON
 */
export const generatePptxFillPlan = async (userPrompt, templateBlueprint) => {
  logger.info(`Phase 3: Generating Presentation Plan for prompt: "${userPrompt}"`);

  const totalAvailableSlides = templateBlueprint.slides?.length || 0;
  const blueprintSummary = (templateBlueprint.slides || []).slice(0, 15).map(s => ({
    slideIndex: s.slideIndex,
    category: s.layoutCategory || 'layout',
    shapeNames: s.shapeNames || (s.containers || []).map(c => c.id || c),
    sampleTexts: (s.sampleTexts || []).slice(0, 5)
  }));

  const systemPrompt = `
You are an executive presentation director and copywriter.
A user wants to create a PowerPoint presentation based on a topic prompt.
We have a master presentation template with pre-designed slide layouts.

USER PROMPT: "${userPrompt}"
MASTER TEMPLATE NAME: "${templateBlueprint.templateName || 'Master Deck'}"
TOTAL TEMPLATE SLIDES AVAILABLE: ${totalAvailableSlides}

MASTER SLIDE LAYOUT BLUEPRINT INDEX:
${JSON.stringify(blueprintSummary, null, 2)}

LANGUAGE & DIRECTIVE RULES:
1. INPUT LANGUAGE FLEXIBILITY: You can receive user prompts in ANY language (French, Spanish, German, Tagalog, Chinese, Japanese, etc.).
2. OUTPUT LANGUAGE: ALL generated presentation titles, subtitles, card headers, body text descriptions, and speaker notes MUST BE WRITTEN IN HIGH-QUALITY EXECUTIVE ENGLISH.
3. Overwrite all template placeholder text with fresh, topic-specific English content.

YOUR TASK:
1. Select the 4 to 8 best matching slide layouts from the master template to form a coherent presentation deck.
2. Slide 1 MUST be a cover/title slide.
3. Write topic-specific, highly professional, realistic ENGLISH content for EVERY SINGLE SHAPE listed in "shapeNames" for each selected slide.
4. Set keys in "fillContent" matching EVERY shape name listed in "shapeNames". If a shape is a footer, module tag, subtitle, or author label, overwrite it with appropriate topic text. NEVER leave any shape unmapped.

JSON SCHEMA TO RETURN (Strict JSON, no markdown wrappers):
{
  "presentationTitle": "Main Presentation Title in English",
  "topic": "${userPrompt}",
  "selectedSlides": [
    {
      "slideIndex": 1,
      "layoutCategory": "title_slide",
      "fillContent": {
        "TextBox 3": "AI-POWERED LINKEDIN STRATEGY",
        "TextBox 4": "EXECUTIVE AUDIENCE SCALING ENGINE",
        "TextBox 5": "PRESENTED BY DOCS-SERVICE"
      },
      "speakerNotes": "Welcome audience and introduce presentation overview."
    }
  ]
}

REQUISITE RULES:
1. Provide rich, highly informative, topic-specific English text. No placeholders like "Lorem Ipsum".
2. Match container keys in "fillContent" directly to the exact shape names provided in the blueprint (e.g. "TextBox 3", "TextBox 4", "main_title").
`;

  const fillPlan = await callGemini(systemPrompt);
  return fillPlan;
};
