import { GoogleGenerativeAI } from '@google/generative-ai';
import { config } from '../config/env.js';
import { logger } from '../utils/logger.js';

// Initialize Gemini API client
const genAI = new GoogleGenerativeAI(config.geminiApiKey);

// Candidate models list
const candidateModels = [
  config.geminiModel,
  'gemini-flash-lite-latest',
  'gemini-2.0-flash-lite-preview-02-05',
  'gemini-1.5-flash-latest'
];

/**
 * Helper to call Gemini model with candidate fallback
 */
async function callGemini(promptText, isJson = true, maxTokens = 8192) {
  let lastError = null;

  for (const modelName of candidateModels) {
    if (!modelName) continue;
    try {
      const model = genAI.getGenerativeModel({
        model: modelName,
        generationConfig: {
          temperature: 0.8,
          maxOutputTokens: maxTokens,
          ...(isJson ? { responseMimeType: 'application/json' } : {})
        }
      });

      const result = await model.generateContent(promptText);
      const responseText = result.response.text();
      return isJson ? JSON.parse(responseText) : responseText;
    } catch (err) {
      logger.warn(`Model "${modelName}" failed: ${err.message}`);
      lastError = err;
    }
  }
  throw lastError || new Error('Gemini API calls failed on all candidate models');
}

/**
 * Multi-Pass Document Generation Engine with:
 * 1. Embedded Prompt tracking inside JSON AST
 * 2. Clean Page Breaks per section in Word
 * 3. Universal Minimum Target Page-Count Scaling Rule (>= N pages, NEVER < N)
 * 4. Automatic Typo & Misspelling Correction
 * 5. Dynamic Domain-Aware Hex Theme Color Generation
 */
export const generateDocumentContent = async (userPrompt) => {
  if (!config.geminiApiKey || config.geminiApiKey === 'your_gemini_api_key_here') {
    throw new Error('GEMINI_API_KEY is not configured in .env file');
  }

  logger.info(`Starting Multi-Pass Document Generation for prompt: "${userPrompt}"`);

  // Extract requested target page count from user prompt (e.g., "5 pages", "10 page reviewer")
  const pageMatch = userPrompt.match(/\b(\d+)\s*(?:-| )?page/i);
  const requestedPages = pageMatch ? parseInt(pageMatch[1], 10) : 5;
  const minSectionsTarget = Math.max(requestedPages, 5);

  logger.info(`Target Page Count Detected: ${requestedPages} -> Planning at least ${minSectionsTarget} major sections/pages.`);

  // STEP 1: Plan Outline, Theme, and Table with Target Page-Count Scaling
  const plannerPrompt = `
  You are an elite master document architect.
  Analyze the user's prompt request: "${userPrompt}"

  UNIVERSAL SCALING & QUALITY RULES:
  1. MINIMUM TARGET PAGE RULE: The user has requested a document requiring AT LEAST ${minSectionsTarget} PAGES. In Microsoft Word, EVERY major section will be formatted cleanly onto its OWN DEDICATED NEW PAGE! Therefore, you MUST plan AT LEAST ${minSectionsTarget} (or more) distinct, major section headings to guarantee the generated document meets or exceeds ${minSectionsTarget} full pages (NEVER generate fewer than ${minSectionsTarget} sections).
  2. TYPO & MISSPELLING CORRECTION: Automatically fix all spelling, grammar, and typos in the user's prompt (e.g. "algorthim" -> "Algorithm", "scract" -> "Scratch").
  3. DOMAIN-AWARE THEME: Automatically infer the domain and invent an appropriate, harmonized 6-character Hex color palette (primaryColor, secondaryColor, accentColor, lightBgColor, textColor).

  Tasks:
  1. Create a high-level Document Title, Subtitle, and Running Header text (using perfect spelling).
  2. Generate a Table of Contents containing AT LEAST ${minSectionsTarget} DISTINCT MAJOR SECTION HEADINGS covering every aspect of "${userPrompt}" thoroughly.
  3. Generate a relevant summary data table with headers and rows.

  Return ONLY JSON matching this structure:
  {
    "theme": {
      "primaryColor": "1E3A8A",
      "secondaryColor": "2563EB",
      "accentColor": "0284C7",
      "lightBgColor": "F8FAFC",
      "textColor": "334155"
    },
    "title": "Document Title with Perfect Spelling",
    "subtitle": "Document Subtitle",
    "headerText": "DOCS SERVICE | COMPREHENSIVE DOCUMENTATION",
    "sectionHeadings": [
      "1. Section Heading One",
      "2. Section Heading Two"
      // You MUST provide AT LEAST ${minSectionsTarget} section headings here!
    ],
    "table": {
      "title": "Summary Matrix",
      "headers": ["Header 1", "Header 2", "Header 3", "Header 4"],
      "rows": [
        ["Row 1 Item", "Category / Detail", "Status / Spec", "Metric / Value"]
      ]
    }
  }
  `;

  logger.info('Step 1: Generating Document Outline & Scaling Sections to Target Page Count...');
  const outlinePlan = await callGemini(plannerPrompt, true, 4096);
  logger.info(`Step 1 Complete -> Title: "${outlinePlan.title}", Planned Sections: ${outlinePlan.sectionHeadings?.length || 0}`);

  let sectionHeadings = outlinePlan.sectionHeadings || [];

  // Fallback section scaling if LLM returned fewer than target
  if (sectionHeadings.length < minSectionsTarget) {
    logger.warn(`LLM returned ${sectionHeadings.length} sections, scaling up to target ${minSectionsTarget}`);
    while (sectionHeadings.length < minSectionsTarget) {
      const num = sectionHeadings.length + 1;
      sectionHeadings.push(`${num}. Detailed Module Analysis & Strategic Insights Part ${num}`);
    }
  }

  // STEP 2: Deep Content Generation per Section
  logger.info(`Step 2: Generating Deep Content for all ${sectionHeadings.length} planned sections...`);

  const fullSections = await Promise.all(
    sectionHeadings.map(async (heading, idx) => {
      const sectionPrompt = `
      You are an expert technical author writing Section ${idx + 1} of a document titled "${outlinePlan.title}".
      Section Topic: "${heading}"
      Overall Document Request: "${userPrompt}"

      Task: Write extensive, highly detailed, multi-paragraph content for this section with 100% perfect spelling and grammar.
      Requirements:
      - Write AT LEAST 3 to 5 long, thorough paragraphs (4-6 sentences per paragraph) covering this section topic exhaustively.
      - Provide a strategic Callout Box highlight summarizing key takeaways or critical notes for this section.

      Return ONLY JSON matching:
      {
        "heading": "${heading}",
        "paragraphs": [
          "Detailed paragraph 1...",
          "Detailed paragraph 2...",
          "Detailed paragraph 3..."
        ],
        "calloutBox": "Key strategic takeaway or formula/note for this section"
      }
      `;

      try {
        const secContent = await callGemini(sectionPrompt, true, 4096);
        return secContent;
      } catch (err) {
        logger.warn(`Failed section generation for "${heading}", using fallback content`);
        return {
          heading,
          paragraphs: [`Comprehensive details and deep technical breakdown regarding ${heading} for request "${userPrompt}".`],
          calloutBox: `Key strategic takeaway for ${heading}.`
        };
      }
    })
  );

  logger.info(`Step 2 Complete -> Generated ${fullSections.length} rich sections.`);

  // STEP 3: Assemble Full Document JSON AST (Including userPrompt)
  const finalDocumentJSON = {
    prompt: userPrompt, // 👈 Included directly inside the document AST!
    theme: outlinePlan.theme,
    title: outlinePlan.title,
    subtitle: outlinePlan.subtitle,
    headerText: outlinePlan.headerText,
    sections: fullSections,
    table: outlinePlan.table
  };

  logger.info(`Multi-Pass Document Generation Complete -> Total Sections: ${fullSections.length}`);
  return finalDocumentJSON;
};
