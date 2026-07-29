import fs from 'fs';
import path from 'path';
import admZip from 'adm-zip';
import { execSync } from 'child_process';
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

async function callGeminiVision(parts, isJson = true, maxTokens = 8192) {
  let lastError = null;

  for (const modelName of candidateModels) {
    if (!modelName) continue;
    try {
      const model = genAI.getGenerativeModel({
        model: modelName,
        generationConfig: {
          temperature: 0.2,
          maxOutputTokens: maxTokens,
          ...(isJson ? { responseMimeType: 'application/json' } : {})
        }
      });

      const result = await model.generateContent(parts);
      const responseText = result.response.text();
      return isJson ? JSON.parse(responseText) : responseText;
    } catch (err) {
      logger.warn(`Template Parser Gemini call with model "${modelName}" failed: ${err.message}`);
      lastError = err;
    }
  }
  throw lastError || new Error('Gemini API calls failed for template parsing');
}

/**
 * Executes python pdfMasterExtractor.py to dump all PDF text, images, diagrams, layers, and shapes to cached storage
 */
export const runMasterExtractor = (pdfFilePath) => {
  try {
    const scriptPath = path.join(process.cwd(), 'src', 'services', 'pptx', 'template', 'pdfMasterExtractor.py');
    const outputDir = path.join(process.cwd(), 'temp', 'template_extracted');
    
    logger.info(`Running Master Template Extractor on: "${pdfFilePath}"...`);
    const cmd = `python "${scriptPath}" "${pdfFilePath}" "${outputDir}"`;
    execSync(cmd, { encoding: 'utf8' });

    const metadataPath = path.join(outputDir, 'master_template_data.json');
    if (fs.existsSync(metadataPath)) {
      const content = fs.readFileSync(metadataPath, 'utf8');
      return JSON.parse(content);
    }
  } catch (err) {
    logger.warn(`Master Extractor warning: ${err.message}`);
  }
  return null;
};

/**
 * Inspects a PDF Presentation Template slide-by-slide
 */
export const parsePdfTemplate = async (pdfFilePath) => {
  logger.info(`Analyzing PDF Master Template: "${pdfFilePath}"`);

  if (!fs.existsSync(pdfFilePath)) {
    throw new Error(`PDF Template File not found: ${pdfFilePath}`);
  }

  const masterExtractedData = runMasterExtractor(pdfFilePath);
  const pdfBuffer = fs.readFileSync(pdfFilePath);
  const base64Data = pdfBuffer.toString('base64');

  const pdfPart = {
    inlineData: {
      data: base64Data,
      mimeType: 'application/pdf'
    }
  };

  const promptText = `
You are a senior executive presentation layout parser.
Examine every page of this PDF presentation deck.

JSON SCHEMA TO RETURN (Strict JSON, no markdown formatting):
{
  "templateName": "${path.basename(pdfFilePath)}",
  "totalSlides": 43,
  "brandTheme": {
    "primaryColor": "000000",
    "secondaryColor": "475569",
    "accentColor": "2563EB",
    "backgroundColor": "FFFFFF",
    "fontFamily": "Montserrat"
  },
  "slides": [
    {
      "slideIndex": 1,
      "layoutCategory": "title_slide",
      "title": "Cover Title",
      "description": "Cover title slide with brand header",
      "containers": [
        { "id": "main_title", "type": "title", "sampleText": "AGENCE DÉCLIC" }
      ]
    }
  ]
}
`;

  const blueprint = await callGeminiVision([pdfPart, promptText], true);
  if (masterExtractedData) {
    blueprint.extractedData = masterExtractedData;
  }
  return blueprint;
};

/**
 * Inspects a PPTX Master Presentation deck slide-by-slide using OpenXML zip parsing
 */
export const parsePptxTemplate = async (pptxFilePath) => {
  logger.info(`Analyzing PPTX Master Template: "${pptxFilePath}"`);

  if (!fs.existsSync(pptxFilePath)) {
    throw new Error(`PPTX Template File not found: ${pptxFilePath}`);
  }

  // Attempt Deep Python Metadata Extraction (Fonts, Hex Colors, Bounding Boxes, Slide Geometry)
  try {
    const extractorScript = path.join(process.cwd(), 'src', 'services', 'pptx', 'template', 'pptxDeepExtractor.py');
    const tempDir = path.join(process.cwd(), 'temp');
    fs.mkdirSync(tempDir, { recursive: true });
    const outputJsonPath = path.join(tempDir, `extracted_blueprint_${Date.now()}.json`);

    const cmd = `python "${extractorScript}" "${pptxFilePath}" "${outputJsonPath}"`;
    execSync(cmd, { encoding: 'utf8' });

    if (fs.existsSync(outputJsonPath)) {
      const blueprintContent = fs.readFileSync(outputJsonPath, 'utf8');
      fs.unlinkSync(outputJsonPath);
      const blueprint = JSON.parse(blueprintContent);
      logger.info(`✅ Deep Python Template Extraction complete: ${blueprint.totalSlides} slides parsed.`);
      return blueprint;
    }
  } catch (err) {
    logger.warn(`Python Deep Extractor fallback to JS Zip: ${err.message}`);
  }

  // Fallback: JSZip XML Inspection
  try {
    const zip = new admZip(pptxFilePath);
    const zipEntries = zip.getEntries();

    const slideEntries = zipEntries.filter(entry => 
      entry.entryName.startsWith('ppt/slides/slide') && entry.entryName.endsWith('.xml')
    );

    slideEntries.sort((a, b) => {
      const numA = parseInt(a.entryName.match(/slide(\d+)\.xml/)?.[1] || '0', 10);
      const numB = parseInt(b.entryName.match(/slide(\d+)\.xml/)?.[1] || '0', 10);
      return numA - numB;
    });

    const slideMetadata = [];

    slideEntries.forEach((entry, idx) => {
      const xmlContent = entry.getData().toString('utf8');
      const shapeMatches = Array.from(xmlContent.matchAll(/<p:cNvPr[^>]*name="([^"]+)"/g)).map(m => m[1]);
      
      const textShapeMap = [];
      const spRegex = /<p:sp>[\s\S]*?<p:cNvPr[^>]*name="([^"]+)"[\s\S]*?<\/p:sp>/g;
      let match;
      while ((match = spRegex.exec(xmlContent)) !== null) {
        const shapeName = match[1];
        const spXml = match[0];
        const textSample = Array.from(spXml.matchAll(/<a:t[^>]*>(.*?)<\/a:t>/g)).map(m => m[1]).join(' ').trim();
        if (textSample && (shapeName.includes('TextBox') || shapeName.includes('AutoShape') || shapeName.includes('Title'))) {
          textShapeMap.push({ shapeName, sampleText: textSample });
        }
      }

      const category = idx === 0 ? 'title_slide' : (idx === 1 ? 'agenda' : (textShapeMap.length >= 3 ? '3_column_cards' : 'content_slide'));

      slideMetadata.push({
        slideIndex: idx + 1,
        entryName: entry.entryName,
        layoutCategory: category,
        shapeCount: shapeMatches.length,
        shapeNames: textShapeMap.map(t => t.shapeName),
        sampleTexts: textShapeMap.map(t => `${t.shapeName}: "${t.sampleText}"`)
      });
    });

    return {
      templateName: path.basename(pptxFilePath),
      totalSlides: slideEntries.length,
      brandTheme: {
        primaryColor: "071E3D",
        secondaryColor: "1E293B",
        accentColor: "38B6FF",
        fontFamily: "Poppins"
      },
      slides: slideMetadata
    };
  } catch (err) {
    logger.warn(`Zip inspection fallback returned default theme: ${err.message}`);
    return {
      templateName: path.basename(pptxFilePath),
      totalSlides: 5,
      brandTheme: {
        primaryColor: "071E3D",
        secondaryColor: "1E293B",
        accentColor: "38B6FF",
        fontFamily: "Poppins"
      },
      slides: []
    };
  }
};

/**
 * Master Template Inspector Entry Point
 */
export const parsePresentationTemplate = async (templateFilePath) => {
  if (!templateFilePath || !fs.existsSync(templateFilePath)) {
    logger.warn(`No valid template file path provided: ${templateFilePath}`);
    return null;
  }

  let ext = path.extname(templateFilePath).toLowerCase();

  // Inspect file header if Multer uploaded file without extension (e.g. uploads/a6b7c8d9)
  if (!ext) {
    try {
      const buffer = fs.readFileSync(templateFilePath);
      if (buffer.length >= 4 && buffer[0] === 0x50 && buffer[1] === 0x4B) { // PK zip header
        ext = '.pptx';
      } else if (buffer.length >= 4 && buffer.toString('utf8', 0, 4) === '%PDF') {
        ext = '.pdf';
      } else {
        ext = '.pptx';
      }
    } catch (e) {
      ext = '.pptx';
    }
  }

  if (ext === '.pdf') {
    return await parsePdfTemplate(templateFilePath);
  } else {
    return await parsePptxTemplate(templateFilePath);
  }
};
