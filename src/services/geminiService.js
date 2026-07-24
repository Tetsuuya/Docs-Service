import fs from 'fs';
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
 * Accepts string OR array of content parts (strings & inlineData objects)
 */
async function callGemini(promptInput, isJson = true, maxTokens = 8192) {
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

      const result = await model.generateContent(promptInput);
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
 * Multi-Pass Document Generation Engine with EXACT Page-Count Calibration:
 * Formula: 1 Executive Cover Page (Page 1) + (N - 1) Content Sections = EXACTLY N TOTAL PAGES in Word.
 */
export const generateDocumentContent = async (userPrompt, file = null) => {
  if (!config.geminiApiKey || config.geminiApiKey === 'your_gemini_api_key_here') {
    throw new Error('GEMINI_API_KEY is not configured in .env file');
  }

  logger.info(`Starting Calibrated Document Generation for prompt: "${userPrompt}" ${file ? `with file: ${file.originalname}` : ''}`);

  let inlineDataPart = null;
  let fileTextContext = '';
  let uploadedImageBuffer = null;
  let uploadedImageMimeType = null;

  if (file && fs.existsSync(file.path)) {
    try {
      const mimeType = file.mimetype || '';
      if (mimeType.startsWith('image/') || mimeType === 'application/pdf') {
        const fileBuffer = fs.readFileSync(file.path);
        inlineDataPart = {
          inlineData: {
            data: fileBuffer.toString('base64'),
            mimeType: mimeType
          }
        };
        if (mimeType.startsWith('image/')) {
          uploadedImageBuffer = fileBuffer;
          uploadedImageMimeType = mimeType;
        }
        logger.info(`Attached Multimodal Inline File -> Type: ${mimeType}, Size: ${fileBuffer.length} bytes`);
      } else {
        fileTextContext = fs.readFileSync(file.path, 'utf8');
        logger.info(`Extracted Text Context File -> Length: ${fileTextContext.length} chars`);
      }
    } catch (err) {
      logger.warn(`Failed reading uploaded file context: ${err.message}`);
    } finally {
      try {
        if (fs.existsSync(file.path)) fs.unlinkSync(file.path);
      } catch (e) {
        // ignore
      }
    }
  }

  // Extract requested target page count from user prompt (e.g., "exactly 5 pages", "10 page guide")
  const pageMatch = userPrompt.match(/\b(\d+)\s*(?:-| )?page/i);
  const requestedTotalPages = pageMatch ? parseInt(pageMatch[1], 10) : 5;

  // EXACT PAGE COUNT RULE:
  // Page 1 is the Executive Cover Page.
  // The remaining (requestedTotalPages - 1) pages are dedicated content sections.
  // Example: 5 Pages Total = 1 Cover Page + 4 Content Sections.
  const targetSectionsCount = Math.max(1, requestedTotalPages - 1);

  logger.info(`Target Page Count: ${requestedTotalPages} Total Pages -> 1 Cover Page + ${targetSectionsCount} Content Sections.`);

  // STEP 1: Plan Outline, Theme, and Summary Matrix
  let plannerPrompt = `
  You are an elite master document architect designing a Claude-level executive publication document.
  Analyze the user's prompt request: "${userPrompt}"
  ${fileTextContext ? `\n\nATTACHED REFERENCE CONTEXT FILE CONTENT:\n"""\n${fileTextContext.substring(0, 10000)}\n"""\n` : ''}

  CRITICAL PAGE COUNT RULE:
  The document must contain EXACTLY ${requestedTotalPages} TOTAL PAGES in Microsoft Word:
  - Page 1: Executive Cover Page
  - Pages 2 to ${requestedTotalPages}: EXACTLY ${targetSectionsCount} CONTENT SECTIONS.
  You MUST provide EXACTLY ${targetSectionsCount} distinct major section headings in the "sectionHeadings" array. No more, no less!

  QUALITY & FORMATTING RULES:
  1. INLINE MARKDOWN FORMATTING: Use inline markdown formatting in text (e.g., **bold**, *italics*, \`code\`).
  2. DOMAIN-AWARE THEME & COLOR PALETTE: Generate a harmonized 6-character Hex color palette (primaryColor, secondaryColor, accentColor, lightBgColor, textColor).
  3. ACCURATE TITLE & BADGE OCR: If an attached screenshot or PDF contains title text, use that exact text string.
  4. CLASSIFICATION TAG: Invent an uppercase classification tag for "docTypeTag" (e.g., "PRODUCT SPECIFICATION GUIDE", "EXECUTIVE BRAND AUDIT").

  Tasks:
  1. Create Document Title, Subtitle, docTypeTag, and Running Header text.
  2. Generate a Table of Contents containing EXACTLY ${targetSectionsCount} MAJOR SECTION HEADINGS.
  3. Generate a relevant summary data table (4 columns, 3-4 rows).

  Return ONLY JSON matching:
  {
    "theme": {
      "primaryColor": "1E3A8A",
      "secondaryColor": "2563EB",
      "accentColor": "0284C7",
      "lightBgColor": "F8FAFC",
      "textColor": "334155"
    },
    "docTypeTag": "PRODUCT SPECIFICATION GUIDE",
    "title": "Exact Title or Topic with Perfect Spelling",
    "subtitle": "Comprehensive Strategic & Technical Documentation",
    "executiveOverview": "Executive overview providing strategic context with **key metrics** and *quality assurance standards*.",
    "headerText": "DOCS SERVICE | EXECUTIVE PUBLICATION",
    "sectionHeadings": [
      // You MUST provide EXACTLY ${targetSectionsCount} section headings here!
    ],
    "table": {
      "title": "Summary Performance Matrix",
      "headers": ["Category / Metric", "Standard", "Observed", "Status"],
      "rows": [
        ["Nutritional Value", "Standard", "**Optimized**", "✅ Certified"]
      ]
    }
  }
  `;

  const plannerInput = inlineDataPart ? [inlineDataPart, plannerPrompt] : plannerPrompt;

  logger.info(`Step 1: Generating Document Outline for EXACTLY ${targetSectionsCount} content sections...`);
  const outlinePlan = await callGemini(plannerInput, true, 4096);
  logger.info(`Step 1 Complete -> Title: "${outlinePlan.title}", Sections Planned: ${outlinePlan.sectionHeadings?.length || 0}`);

  let sectionHeadings = outlinePlan.sectionHeadings || [];

  // Enforce EXACT target section count
  if (sectionHeadings.length > targetSectionsCount) {
    sectionHeadings = sectionHeadings.slice(0, targetSectionsCount);
  } else if (sectionHeadings.length < targetSectionsCount) {
    logger.warn(`LLM returned ${sectionHeadings.length} sections, calibrating to exact target ${targetSectionsCount}`);
    while (sectionHeadings.length < targetSectionsCount) {
      const num = sectionHeadings.length + 1;
      sectionHeadings.push(`${num}. Detailed Module Analysis & Strategic Roadmap Part ${num}`);
    }
  }

  // STEP 2: Deep Content Generation per Section with Rich AST Blocks
  logger.info(`Step 2: Generating Calibrated Content for all ${sectionHeadings.length} planned sections...`);

  const fullSections = await Promise.all(
    sectionHeadings.map(async (heading, idx) => {
      const sectionPrompt = `
      You are an expert technical author writing Section ${idx + 1} of ${sectionHeadings.length} for document "${outlinePlan.title}".
      Section Topic: "${heading}"
      Overall Document Request: "${userPrompt}"
      ${fileTextContext ? `\nReference Context Material:\n"""\n${fileTextContext.substring(0, 4000)}\n"""\n` : ''}

      CRITICAL DENSITY REQUIREMENT:
      This content section will occupy EXACTLY Page ${idx + 2} of the Word document.
      Write clear, medium-length content so it fits cleanly on 1 physical Word page without overflowing to the next page:
      - 2 concise paragraphs (3-4 sentences each) with **bold key terms** and *italicized emphasis*.
      - 3 bullet points in "bulletList".
      - 1 "calloutBox" object: { "type": "info" | "warning" | "success" | "tip", "title": "...", "text": "..." }.
      - (Optional) "statCards": 2-3 key stat cards OR "codeBlock": { "language": "...", "code": "..." }.

      Return ONLY JSON matching:
      {
        "heading": "${heading}",
        "paragraphs": [
          "Detailed paragraph 1 with **bold metrics**...",
          "Detailed paragraph 2 with *key insights*..."
        ],
        "bulletList": [
          "**Key Objective 1**: Description",
          "**Key Objective 2**: Description",
          "**Key Objective 3**: Description"
        ],
        "calloutBox": {
          "type": "tip",
          "title": "Key Strategic Takeaway",
          "text": "Critical summary note for this section topic."
        }
      }
      `;

      try {
        const secInput = inlineDataPart ? [inlineDataPart, sectionPrompt] : sectionPrompt;
        const secContent = await callGemini(secInput, true, 4096);
        return secContent;
      } catch (err) {
        logger.warn(`Failed section generation for "${heading}", using fallback content`);
        return {
          heading,
          paragraphs: [`Comprehensive details and strategic breakdown regarding **${heading}** for request "${userPrompt}".`],
          bulletList: [`**Core specification**: Specification details for ${heading}`, `**Metric target**: Benchmark metrics for ${heading}`],
          calloutBox: {
            type: 'info',
            title: 'Section Summary',
            text: `Key strategic takeaway and analysis regarding ${heading}.`
          }
        };
      }
    })
  );

  logger.info(`Step 2 Complete -> Generated ${fullSections.length} calibrated sections.`);

  // STEP 3: Assemble Full Document JSON AST
  const finalDocumentJSON = {
    prompt: userPrompt,
    requestedTotalPages: requestedTotalPages,
    theme: outlinePlan.theme,
    docTypeTag: outlinePlan.docTypeTag || 'PRODUCT SPECIFICATION GUIDE',
    title: outlinePlan.title,
    subtitle: outlinePlan.subtitle,
    executiveOverview: outlinePlan.executiveOverview || `This document outlines comprehensive technical analysis and strategic roadmap details regarding **${outlinePlan.title}**.`,
    sectionHeadings: sectionHeadings,
    headerText: outlinePlan.headerText,
    pages: fullSections,
    sections: fullSections,
    table: outlinePlan.table,
    imageBuffer: uploadedImageBuffer,
    imageMimeType: uploadedImageMimeType
  };

  logger.info(`Calibrated Document Generation Complete -> Target Total Pages: ${requestedTotalPages} (1 Cover + ${fullSections.length} Sections)`);
  return finalDocumentJSON;
};
