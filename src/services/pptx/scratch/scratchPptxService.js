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

/**
 * Call Gemini API with model fallback
 */
async function callGemini(promptInput, isJson = true, maxTokens = 8192) {
  let lastError = null;

  for (const modelName of candidateModels) {
    if (!modelName) continue;
    try {
      const model = genAI.getGenerativeModel({
        model: modelName,
        generationConfig: {
          temperature: 0.7,
          maxOutputTokens: maxTokens,
          ...(isJson ? { responseMimeType: 'application/json' } : {})
        }
      });

      const result = await model.generateContent(promptInput);
      const responseText = result.response.text();
      return isJson ? JSON.parse(responseText) : responseText;
    } catch (err) {
      logger.warn(`PPTX Gemini call with model "${modelName}" failed: ${err.message}`);
      lastError = err;
    }
  }
  throw lastError || new Error('Gemini API calls failed on all candidate models');
}

/**
 * Generate structured presentation JSON content using Gemini AI
 */
export const generatePptxContent = async (userPrompt) => {
  if (!config.geminiApiKey || config.geminiApiKey === 'your_gemini_api_key_here') {
    throw new Error('GEMINI_API_KEY is not configured in .env file');
  }

  // Detect slide count request or default to 6 slides
  const slideMatch = userPrompt.match(/\b(\d+)\s*(?:-| )?slide/i);
  const targetSlideCount = slideMatch ? Math.min(Math.max(parseInt(slideMatch[1], 10), 3), 15) : 6;

  logger.info(`Generating PPTX Content for prompt: "${userPrompt}" (Target Slides: ${targetSlideCount})`);

  const systemPrompt = `
You are an expert executive presentation designer and copywriter.
Generate a structured JSON definition for a high-impact, professional PowerPoint presentation based on the user request.

USER PROMPT: "${userPrompt}"
TARGET SLIDE COUNT: ${targetSlideCount} slides (Slide 1 must be Title, last slide must be Conclusion/Call-to-Action).

JSON SCHEMA TO RETURN (Strict JSON, no markdown formatting):
{
  "title": "Presentation Title",
  "subtitle": "Catchy Subtitle",
  "theme": {
    "primaryColor": "071E3D",     // Dark text color (navy/charcoal) - ADAPT to topic theme
    "secondaryColor": "0070FF",   // Accent highlight color (blue/teal/etc) - ADAPT to topic theme
    "accentColor": "489EF9",      // Secondary accent (lighter version) - ADAPT to topic theme
    "backgroundColor": "FFFFFF",  // Clean white/light background
    "cardBgColor": "F8F9FA",      // Subtle card background (light gray)
    "textColor": "032853",        // Body text color (dark)
    "fontFamily": "Inter"         // Clean modern sans-serif
  },
  "slides": [
    {
      "slideNumber": 1,
      "type": "title",
      "title": "Presentation Title",
      "subtitle": "Subtitle or presenter details",
      "hasImage": true,
      "imagePrompt": "Detailed highly specific 3D cinematic cover art prompt matching the presentation title topic",
      "speakerNotes": "Introduction speaker notes..."
    },
    {
      "slideNumber": 2,
      "type": "content",
      "layout": "split",          // Available layouts: "split", "cards", "stat"
      "title": "Section Title 1",
      "subtitle": "Section subtitle",
      "hasImage": true,
      "imagePrompt": "Detailed 3D graphic prompt specifically illustrating Section 1 concept",
      "bullets": [
        "First key takeaway bullet point",
        "Second key takeaway bullet point"
      ],
      "speakerNotes": "Speaker notes for slide 2..."
    },
    {
      "slideNumber": 3,
      "type": "content",
      "layout": "cards",
      "title": "Key Pillars",
      "hasImage": false,
      "cards": [
        { "title": "Pillar 1", "description": "Clear explanation paragraph for point 1." },
        { "title": "Pillar 2", "description": "Clear explanation paragraph for point 2." },
        { "title": "Pillar 3", "description": "Clear explanation paragraph for point 3." }
      ],
      "speakerNotes": "Speaker notes..."
    },
    {
      "slideNumber": 4,
      "type": "content",
      "layout": "split",
      "title": "Section Title 2",
      "hasImage": true,
      "imagePrompt": "Detailed 3D graphic prompt specifically illustrating Section 2 topic concept",
      "bullets": [
        "Key observation detail",
        "Supporting analysis detail"
      ]
    }
  ]
}

REQUISITE RULES:
1. Provide rich, highly informative, and realistic business text. No placeholders like "Lorem Ipsum" or "Add text here".
2. Set "hasImage": true and write a unique, highly specific "imagePrompt" string for slides that use visual split layouts (such as Slide 1, Slide 2, Slide 4).
3. Make each "imagePrompt" highly descriptive and unique to the slide's specific sub-topic (e.g. for robots: "3D surgical robotic arm performing micro-surgery", for planets: "3D render of Jupiter and Saturn with glowing icy rings").
4. CRITICAL: Adapt the theme colors to match the presentation topic:
   - For technology/AI/robotics: Use cool blues (#071E3D primary, #0070FF accent, #489EF9 light accent)
   - For nature/environment: Use greens (#0A3622 primary, #00A86B accent, #7FE5B0 light)
   - For finance/business: Use navy/gold (#1A2332 primary, #FFD700 accent, #FFF4CC light)
   - For health/medical: Use blue/green (#003D5B primary, #00A7B5 accent, #A0E7E5 light)
   - For energy/fire: Use reds/oranges (#4A1C00 primary, #FF5722 accent, #FFB39C light)
   - For luxury/premium: Use purple/gold (#2D1B69 primary, #9B51E0 accent, #D4B3FF light)
   - For education: Use teal/yellow (#0D3B4F primary, #00BCD4 accent, #FFD54F light)
   Always keep backgroundColor FFFFFF (white) and cardBgColor F8F9FA (light gray) for readability.
5. Ensure slide 5 (Challenges/Stats) has complete content with statNumber, statLabel, and bullets filled out properly.
`;

  const presentationData = await callGemini(systemPrompt, true);
  return presentationData;
};
