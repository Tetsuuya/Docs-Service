import fs from 'fs';
import { GoogleGenerativeAI } from '@google/generative-ai';
import { config } from '../../../config/env.js';
import { logger } from '../../../utils/logger.js';

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
 * Exported for use in template service
 */
export async function callGemini(promptInput, isJson = true, maxTokens = 8192) {
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
 * Multi-Pass Document Generation Engine with Pure Content-Driven Visual Decision:
 * 100% Dynamic - Visuals (diagram, image, table, none) are chosen strictly based on what the section content actually describes!
 * Zero pre-defined slots or hardcoded forced elements.
 */
export const generateDocumentContent = async (userPrompt, file = null) => {
  if (!config.geminiApiKey || config.geminiApiKey === 'your_gemini_api_key_here') {
    throw new Error('GEMINI_API_KEY is not configured in .env file');
  }

  logger.info(`Starting Content-Driven AI Document Generation for prompt: "${userPrompt}" ${file ? `with file: ${file.originalname}` : ''}`);

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
  1. INLINE MARKDOWN FORMATTING: Use inline markdown formatting in text:
     - **bold** for emphasis
     - *italics* for subtle emphasis
     - \`code\` for technical terms
     - ~~strikethrough~~ for deprecated/old information
     - __underline__ for legal terms or definitions
     - ==highlight== for critical warnings or key takeaways
     - [link text](url) for external references and documentation links
  
  2. DOMAIN-AWARE COLOR PALETTE: Generate a harmonized 6-character Hex color palette (primaryColor, secondaryColor, accentColor, lightBgColor, textColor).
  
  3. CLASSIFICATION TAG: Invent an uppercase classification tag for "docTypeTag" (e.g., "ASTRONOMICAL TECHNICAL STUDY", "ENTERPRISE REPORT", "CULTURAL HISTORICAL STUDY").

  4. DOCUMENT METADATA: Provide author name, keywords (comma-separated), and subject for document properties.

  Tasks:
  1. Create Document Title, Subtitle, docTypeTag, and Running Header text.
  2. Generate a Table of Contents containing EXACTLY ${targetSectionsCount} MAJOR SECTION HEADINGS.
  3. Generate a relevant summary data table.
  4. Add strategic hyperlinks where official documentation or sources are referenced.

  Return ONLY JSON matching:
  {
    "theme": {
      "primaryColor": "1E3A8A",
      "secondaryColor": "2563EB",
      "accentColor": "0284C7",
      "lightBgColor": "F8FAFC",
      "textColor": "334155"
    },
    "docTypeTag": "ENTERPRISE REPORT",
    "title": "Exact Title or Topic with Perfect Spelling",
    "subtitle": "Comprehensive Strategic & Technical Documentation",
    "author": "AI Documentation Engine",
    "keywords": "enterprise, strategy, technical, analysis",
    "executiveOverview": "Executive overview providing strategic context with **key metrics** and *quality assurance standards*. See [official documentation](https://example.com) for details.",
    "headerText": "DOCS SERVICE | EXECUTIVE PUBLICATION",
    "sectionHeadings": [
      // You MUST provide EXACTLY ${targetSectionsCount} section headings here!
    ],
    "table": {
      "title": "Summary Matrix",
      "headers": ["Category / Metric", "Baseline", "Optimized", "Status"],
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

  // STEP 2: Pure Content-Driven Visual Decision per Section
  logger.info(`Step 2: Generating Calibrated Content with Content-Driven Visual Decision for all ${sectionHeadings.length} sections...`);

  const fullSections = await Promise.all(
    sectionHeadings.map(async (heading, idx) => {
      const sectionPrompt = `
      You are an expert human publishing director writing Section ${idx + 1} of ${sectionHeadings.length} for document "${outlinePlan.title}".
      Section Topic: "${heading}"
      Overall Document Request: "${userPrompt}"
      ${fileTextContext ? `\nReference Context Material:\n"""\n${fileTextContext.substring(0, 4000)}\n"""\n` : ''}

      AI HUMAN-JUDGMENT FORMATTING INTELLIGENCE:
      You have complete editorial control. Make smart, context-aware decisions about:

      1. HEADING HIERARCHY (headingLevel: 1-6):
         - Is this a MAJOR section? → H1 (level 1)
         - Sub-section under major topic? → H2 (level 2)
         - Detailed subsection? → H3 (level 3)
         - Minor point? → H4-H6 (level 4-6)
         Base on: content importance, document structure, topic depth

      2. HEADING ALIGNMENT (headingAlignment):
         - Major titles, banners? → "center"
         - Standard sections? → "left"
         - Special emphasis? → "center"

      3. PARAGRAPH ALIGNMENT (for each paragraph object):
         - Body text, explanations? → "left" or "justify"
         - Standalone statements, quotes? → "center"
         - Metadata, dates? → "right"
         - Dense professional content? → "justify"

      4. LIST TYPE INTELLIGENCE (listType):
         - Sequential steps, procedures, instructions? → "numbered"
         - Features, benefits, unordered items? → "bullet"
         - No list needed (narrative flow)? → omit bulletList entirely
         
      5. MULTI-LEVEL LISTS (for nested hierarchy):
         - Items can have "level": 0 (main), 1 (sub), 2 (sub-sub)
         - Use levels for: sub-steps, nested features, hierarchical info

      6. INLINE MARKDOWN FORMATTING (use intelligently in text):
         - **bold** for emphasis, key terms
         - *italics* for subtle emphasis
         - \`code\` for technical terms, commands, code
         - ~~strikethrough~~ for deprecated APIs, old versions
         - __underline__ for legal terms, important definitions
         - ==highlight== for CRITICAL warnings, key takeaways
         - [link text](url) for documentation, sources, references
         Apply formatting where it ENHANCES meaning, not decoration

      7. HORIZONTAL DIVIDERS (horizontalDivider: true/false):
         - Add after section if major topic transition follows
         - DON'T add if seamless narrative continues
         - Consider for visual breathing room

      8. VISUAL ELEMENTS (Pure Content-Based Decision):
         - Physical objects, scenes, products? → "image"
         - Step-by-step processes, workflows? → "diagram"
         - Numerical data, specs, comparisons? → "table"
         - Pure narrative, history, policy? → "none"

      Return ONLY JSON matching:
      {
        "heading": "${heading}",
        "headingLevel": 1,
        "headingAlignment": "left",
        "paragraphs": [
          {
            "text": "Paragraph with **bold**, *italic*, \`code\`, ~~strike~~, __underline__, ==highlight==, and [links](https://example.com)...",
            "alignment": "justify"
          },
          "Simple string paragraphs also work and default to left alignment"
        ],
        "listType": "numbered",
        "bulletList": [
          {
            "text": "**Main item 1**: Description with formatting",
            "level": 0
          },
          {
            "text": "Sub-item under item 1",
            "level": 1
          },
          {
            "text": "**Main item 2**: Another main point",
            "level": 0
          }
        ],
        "horizontalDivider": false,
        "calloutBox": {
          "type": "tip",
          "title": "Strategic Highlight",
          "text": "Use ==highlight== and **bold** in callouts too!"
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
          headingLevel: 2,
          headingAlignment: 'left',
          paragraphs: [`Comprehensive details and strategic breakdown regarding **${heading}** for request "${userPrompt}".`],
          listType: 'bullet',
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

  // Consecutive Duplicate Prevention (Prevents identical visual cards on adjacent pages)
  let lastType = null;
  fullSections.forEach((sec, idx) => {
    const currentType = sec.visualNeed ? sec.visualNeed.type : 'none';
    if (currentType !== 'none' && currentType === lastType) {
      logger.info(`Section ${idx + 1} ("${sec.heading}") has consecutive duplicate visual "${currentType}" -> setting to "none" for page flow.`);
      sec.visualNeed.type = 'none';
      lastType = 'none';
    } else {
      lastType = currentType;
    }
  });

  logger.info(`Step 2 Complete -> Generated ${fullSections.length} calibrated sections based purely on content.`);

  // STEP 3: Assemble Full Document JSON AST with AI-determined metadata
  const finalDocumentJSON = {
    prompt: userPrompt,
    requestedTotalPages: requestedTotalPages,
    theme: outlinePlan.theme,
    docTypeTag: outlinePlan.docTypeTag || 'ENTERPRISE REPORT',
    title: outlinePlan.title,
    subtitle: outlinePlan.subtitle,
    author: outlinePlan.author || 'AI Documentation Engine',
    keywords: outlinePlan.keywords || 'professional, documentation, analysis',
    executiveOverview: outlinePlan.executiveOverview || `This document outlines comprehensive technical analysis and strategic roadmap details regarding **${outlinePlan.title}**.`,
    sectionHeadings: sectionHeadings,
    headerText: outlinePlan.headerText,
    pages: fullSections,
    sections: fullSections,
    table: outlinePlan.table,
    imageBuffer: uploadedImageBuffer,
    imageMimeType: uploadedImageMimeType
  };

  logger.info(`AI Human-Judgment Document Generation Complete -> Target Total Pages: ${requestedTotalPages}`);
  return finalDocumentJSON;
};
