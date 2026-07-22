import { GoogleGenerativeAI } from '@google/generative-ai';
import { config } from '../config/env.js';
import { logger } from '../utils/logger.js';

// Initialize Gemini API client
const genAI = new GoogleGenerativeAI(config.geminiApiKey);

/**
 * Service to interact with Gemini API for structured document content planning.
 * Configured with working model string: gemini-flash-lite-latest
 */
export const generateDocumentContent = async (userPrompt) => {
  if (!config.geminiApiKey || config.geminiApiKey === 'your_gemini_api_key_here') {
    throw new Error('GEMINI_API_KEY is not configured in .env file');
  }

  // Model fallback candidate list matching boss's key configuration
  const candidateModels = [
    config.geminiModel,
    'gemini-flash-lite-latest',
    'gemini-2.0-flash-lite-preview-02-05',
    'gemini-1.5-flash-latest'
  ];

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

  let lastError = null;

  for (const modelName of candidateModels) {
    if (!modelName) continue;
    try {
      logger.info(`Attempting Gemini API call with model: "${modelName}"`);
      const model = genAI.getGenerativeModel({
        model: modelName,
        generationConfig: {
          responseMimeType: 'application/json'
        }
      });

      const result = await model.generateContent(prompt);
      const responseText = result.response.text();
      logger.info(`Gemini API Response Received successfully from model "${modelName}" (${responseText.length} chars)`);

      return JSON.parse(responseText);
    } catch (err) {
      logger.warn(`Model "${modelName}" failed: ${err.message}`);
      lastError = err;
    }
  }

  logger.error('All candidate Gemini models failed');
  throw lastError || new Error('Failed to generate document content with Gemini API');
};
