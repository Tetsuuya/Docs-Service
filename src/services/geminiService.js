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
 * Multi-Pass Document Generation Engine with Diverse Human-Level Editorial Layouts:
 * 1. Multimodal File & Design Context Extraction
 * 2. Human Editorial Variety Detector per Section (Ensures varied visual types across pages)
 * 3. Structured Image Placeholder Cards (description & figure caption)
 * 4. Page Breaks per section in Word with titlePage header/footer suppression
 * 5. Target Page-Count Scaling Rule (1 Cover + (N-1) Content Sections = N Total Pages)
 */
export const generateDocumentContent = async (userPrompt, file = null) => {
  if (!config.geminiApiKey || config.geminiApiKey === 'your_gemini_api_key_here') {
    throw new Error('GEMINI_API_KEY is not configured in .env file');
  }

  logger.info(`Starting Human-Level AI Document Generation for prompt: "${userPrompt}" ${file ? `with file: ${file.originalname}` : ''}`);

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

  // Extract requested target page count from user prompt
  const pageMatch = userPrompt.match(/\b(\d+)\s*(?:-| )?page/i);
  const requestedTotalPages = pageMatch ? parseInt(pageMatch[1], 10) : 5;
  const targetSectionsCount = Math.max(1, requestedTotalPages - 1);

  logger.info(`Target Page Count: ${requestedTotalPages} Total Pages -> 1 Cover Page + ${targetSectionsCount} Content Sections.`);

  // STEP 1: Plan Outline, Theme, and Summary Matrix
  let plannerPrompt = `
  You are an elite master document architect designing a publication-grade document.
  Analyze the user's prompt request: "${userPrompt}"
  ${fileTextContext ? `\n\nATTACHED REFERENCE CONTEXT FILE CONTENT:\n"""\n${fileTextContext.substring(0, 10000)}\n"""\n` : ''}

  CRITICAL PAGE COUNT RULE:
  The document must contain EXACTLY ${requestedTotalPages} TOTAL PAGES in Microsoft Word:
  - Page 1: Executive Cover Page
  - Pages 2 to ${requestedTotalPages}: EXACTLY ${targetSectionsCount} CONTENT SECTIONS.
  Provide EXACTLY ${targetSectionsCount} distinct major section headings in "sectionHeadings".

  FORMATTING & THEME RULES:
  1. INLINE MARKDOWN FORMATTING: Use inline markdown formatting in text (e.g., **bold**, *italics*, \`code\`).
  2. DOMAIN-AWARE COLOR PALETTE: Generate a harmonized 6-character Hex color palette (primaryColor, secondaryColor, accentColor, lightBgColor, textColor).
  3. CLASSIFICATION TAG: Invent an uppercase classification tag for "docTypeTag" (e.g., "ENTERPRISE TECHNICAL REPORT", "CULTURAL HISTORICAL STUDY", "PRODUCT SPECIFICATION GUIDE").

  Tasks:
  1. Create Document Title, Subtitle, docTypeTag, and Running Header text.
  2. Generate a Table of Contents containing EXACTLY ${targetSectionsCount} MAJOR SECTION HEADINGS.
  3. Generate a relevant summary data table.

  Return ONLY JSON matching:
  {
    "theme": {
      "primaryColor": "1E3A8A",
      "secondaryColor": "2563EB",
      "accentColor": "0284C7",
      "lightBgColor": "F8FAFC",
      "textColor": "334155"
    },
    "docTypeTag": "ENTERPRISE TECHNICAL REPORT",
    "title": "Exact Title or Topic with Perfect Spelling",
    "subtitle": "Comprehensive Strategic & Technical Documentation",
    "executiveOverview": "Executive overview providing strategic context with **key metrics** and *quality assurance standards*.",
    "headerText": "DOCS SERVICE | EXECUTIVE PUBLICATION",
    "sectionHeadings": [
      // You MUST provide EXACTLY ${targetSectionsCount} section headings here!
    ],
    "table": {
      "title": "Summary Matrix",
      "headers": ["Category / Metric", "Baseline", "Observed", "Status"],
      "rows": [
        ["Core Parameter", "Standard", "**Optimized**", "✅ Verified"]
      ]
    }
  }
  `;

  const plannerInput = inlineDataPart ? [inlineDataPart, plannerPrompt] : plannerPrompt;

  logger.info(`Step 1: Generating Document Outline for EXACTLY ${targetSectionsCount} content sections...`);
  const outlinePlan = await callGemini(plannerInput, true, 4096);
  logger.info(`Step 1 Complete -> Title: "${outlinePlan.title}", Sections Planned: ${outlinePlan.sectionHeadings?.length || 0}`);

  let sectionHeadings = outlinePlan.sectionHeadings || [];

  if (sectionHeadings.length > targetSectionsCount) {
    sectionHeadings = sectionHeadings.slice(0, targetSectionsCount);
  } else if (sectionHeadings.length < targetSectionsCount) {
    while (sectionHeadings.length < targetSectionsCount) {
      const num = sectionHeadings.length + 1;
      sectionHeadings.push(`${num}. Detailed Module Analysis & Strategic Roadmap Part ${num}`);
    }
  }

  // STEP 2: Deep Content Generation per Section with Human-Level Editorial Variety
  logger.info(`Step 2: Generating Calibrated Content with Varied Editorial Layouts for all ${sectionHeadings.length} sections...`);

  const fullSections = await Promise.all(
    sectionHeadings.map(async (heading, idx) => {
      const sectionPrompt = `
      You are an expert human publishing director writing Section ${idx + 1} of ${sectionHeadings.length} for document "${outlinePlan.title}".
      Section Topic: "${heading}"
      Overall Document Request: "${userPrompt}"
      ${fileTextContext ? `\nReference Context Material:\n"""\n${fileTextContext.substring(0, 4000)}\n"""\n` : ''}

      STRICT HUMAN EDITORIAL BALANCE & VARIETY RULES:
      Act like a human editor designing a magazine or publication. Vary the visual components across pages so the document NEVER feels like a repetitive template:
      - Section ${idx + 1} of ${sectionHeadings.length}: Pick the single best visualNeed ("diagram" | "image" | "table" | "none").
      - "diagram": Best for step-by-step workflows, timelines, historical accord sequences, or architecture pipelines.
      - "image": Best for vivid physical scenes, landscapes, species, or cultural photos. Limit images to AT MOST 1 or 2 sections across the entire document!
      - "table": Best for numerical metrics, legal/demographic breakdowns, or specs.
      - "none": Best for narrative, legal background, policy, or governance text (uses clean paragraphs + KPI stat cards or callout box).

      Return ONLY JSON matching:
      {
        "heading": "${heading}",
        "paragraphs": [
          "Paragraph 1 with **bold metrics**...",
          "Paragraph 2 with *key insights*..."
        ],
        "bulletList": [
          "**Key Spec 1**: Detailed explanation",
          "**Key Spec 2**: Detailed explanation"
        ],
        "calloutBox": {
          "type": "tip",
          "title": "Strategic Highlight",
          "text": "Key takeaway note for this topic."
        },
        "visualNeed": {
          "type": "diagram" | "image" | "table" | "none",
          "diagram": {
            "title": "Process / Workflow Title",
            "steps": ["Step 1 Title", "Step 2 Title", "Step 3 Title", "Step 4 Title"]
          },
          "image": {
            "imagePrompt": "Detailed description of the image/photo recommended for this section",
            "caption": "Figure ${idx + 1}.1: Descriptive caption of the recommended visual"
          },
          "table": {
            "title": "Section Data Table",
            "headers": ["Header 1", "Header 2", "Header 3"],
            "rows": [
              ["Row 1 Data", "Value", "Status"]
            ]
          }
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
          },
          visualNeed: { type: 'none' }
        };
      }
    })
  );

  // Post-Processing: Normalize visual types to prevent repetitive "all images" or "all diagrams" template feel!
  let imageCount = 0;
  fullSections.forEach((sec, idx) => {
    const vType = sec.visualNeed ? sec.visualNeed.type : 'none';
    if (vType === 'image') {
      imageCount++;
      // If more than 2 sections have images, convert excess images to diagram or none for editorial variety
      if (imageCount > 2) {
        logger.info(`Normalizing section ${idx + 1} ("${sec.heading}") from "image" to "none" for editorial variety.`);
        sec.visualNeed.type = 'none';
      }
    }
  });

  logger.info(`Step 2 Complete -> Generated ${fullSections.length} calibrated sections with Varied Editorial Layouts.`);

  // STEP 3: Assemble Full Document JSON AST
  const finalDocumentJSON = {
    prompt: userPrompt,
    requestedTotalPages: requestedTotalPages,
    theme: outlinePlan.theme,
    docTypeTag: outlinePlan.docTypeTag || 'ENTERPRISE TECHNICAL REPORT',
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

  logger.info(`Human-Level Document Generation Complete -> Target Total Pages: ${requestedTotalPages}`);
  return finalDocumentJSON;
};
