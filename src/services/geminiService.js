import { GoogleGenerativeAI } from '@google/generative-ai';
import { config } from '../config/env.js';
import { logger } from '../utils/logger.js';

// Initialize Gemini API client
const genAI = new GoogleGenerativeAI(config.geminiApiKey);

/**
 * Service to interact with Gemini 2.0 Flash Lite for structured document content planning.
 * Enforces JSON output matching title, subtitle, sections, callout boxes, and formatted data tables.
 */
export const generateDocumentContent = async (userPrompt) => {
  logger.info(`Calling Gemini API (${config.geminiModel}) for prompt: "${userPrompt}"`);

  if (!config.geminiApiKey || config.geminiApiKey === 'your_gemini_api_key_here') {
    throw new Error('GEMINI_API_KEY is not configured in .env file');
  }

  // Use Gemini 2.0 Flash Lite model
  const model = genAI.getGenerativeModel({
    model: config.geminiModel,
    generationConfig: {
      responseMimeType: 'application/json'
    }
  });

  const prompt = `
  You are an expert document designer and senior technical report writer.
  Based on the user's prompt request: "${userPrompt}", generate a comprehensive, professional document structure.

  You MUST return ONLY a JSON object with this EXACT structure:
  {
    "title": "Document Title",
    "subtitle": "Executive Summary or Subtitle Line",
    "headerText": "DOCS SERVICE | CONFIDENTIAL REPORT",
    "sections": [
      {
        "heading": "Section Heading Title",
        "paragraphs": [
          "First detailed paragraph explaining the topic...",
          "Second paragraph with key background details..."
        ],
        "calloutBox": "Key takeaway summary box or strategic recommendation for this section"
      },
      {
        "heading": "Secondary Analysis & Metrics",
        "paragraphs": [
          "Paragraph discussing implementation metrics, scope, and technical roadmap..."
        ]
      }
    ],
    "table": {
      "title": "Project Scope & Financial Allocation Table",
      "headers": ["Project Module / Task", "Department", "Status", "Budget Allocation"],
      "rows": [
        ["Infrastructure Cloud Migration", "DevOps & Cloud", "Completed", "$25,000"],
        ["AI Document Generation Engine", "Backend Engineering", "In Progress", "$18,500"],
        ["Security & Compliance Audit", "InfoSec Team", "Planned", "$12,000"],
        ["User Acceptance Testing", "QA & Support", "Pending", "$6,500"]
      ]
    }
  }
  `;

  const result = await model.generateContent(prompt);
  const responseText = result.response.text();
  logger.info(`Gemini API Response Received (${responseText.length} chars)`);

  try {
    const documentJSON = JSON.parse(responseText);
    return documentJSON;
  } catch (parseErr) {
    logger.error('Failed to parse Gemini JSON output:', parseErr.message);
    throw new Error('Invalid JSON response returned by Gemini model');
  }
};
