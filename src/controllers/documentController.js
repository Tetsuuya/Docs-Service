import { generateDocumentContent } from '../services/docx/scratch/scratchModeService.js';
import { generateFillModeDocument } from '../services/docx/fill/fillModeService.js';
import { buildDocxFile } from '../services/docx/scratch/scratchDocxBuilder.js';
import { buildFillModeDocx } from '../services/docx/fill/fillDocxBuilder.js';
import { buildPptxFile } from '../services/pptx/pptxService.js';
import { buildXlsxFile } from '../services/xlsx/xlsxService.js';
import { convertToPdf } from '../services/pdf/pdfService.js';
import { saveDocumentJsonHistory, generateDocumentId } from '../utils/historyStorage.js';
import { parseTemplate } from '../services/docx/shared/templateService.js';
import { logger } from '../utils/logger.js';
import fs from 'fs';

/**
 * Main unified controller endpoint handling format routing (DOCX, PPTX, XLSX, PDF)
 * Supports modes: 'scratch' (generate new), 'fill' (analyze structure & regenerate), 'reference' (inspire from doc)
 */
export const handleGenerateDocument = async (req, res) => {
  try {
    let { prompt, format = 'docx', mode = 'scratch' } = req.body;
    
    // Auto-detect format based on uploaded file extension (if not explicitly set)
    // Only override format if the file extension clearly indicates a different format
    if (req.file) {
      const origName = (req.file.originalname || '').toLowerCase();
      if (origName.endsWith('.pptx')) {
        format = 'pptx';
      } else if (origName.endsWith('.docx')) {
        format = 'docx';
      } else if (origName.endsWith('.xlsx')) {
        format = 'xlsx';
      }
      // Note: Removed ZIP header detection as both .docx and .pptx are ZIP files
    } else if (prompt.toLowerCase().includes('create a ppt') || prompt.toLowerCase().includes('powerpoint')) {
      format = 'pptx';
    }

    logger.info(`Unified Request Received -> Format: ${format}, Mode: ${mode}, Prompt: "${prompt}"`);

    if (!prompt) {
      logger.warn('Unified Request Failed: Prompt is missing');
      return res.status(400).json({ error: 'Prompt is required' });
    }

    // MODE: REFERENCE (use uploaded doc as style/structure reference, generate new content)
    if (mode === 'reference' && format === 'docx') {
      if (!req.file) {
        return res.status(400).json({ error: 'Reference document is required for reference mode' });
      }

      const referencePath = req.file.path;
      const docId = generateDocumentId();

      try {
        const referenceAnalysis = await parseTemplate(referencePath);
        logger.info(`Reference Document Analysis -> Text length: ${referenceAnalysis.fullText.length} chars`);

        const enhancedPrompt = `
        REFERENCE DOCUMENT CONTEXT:
        ${referenceAnalysis.fullText.substring(0, 5000)}
        
        USER REQUEST: ${prompt}
        
        INSTRUCTIONS:
        - Analyze the reference document's structure, tone, and formatting style
        - Generate NEW content based on user request
        - Match the professional tone and structure of the reference
        - Do NOT copy content, but INSPIRE from the style and organization
        `;

        const contentData = await generateDocumentContent(enhancedPrompt, null);
        contentData.id = docId;
        contentData.mode = 'reference';

        if (fs.existsSync(referencePath)) fs.unlinkSync(referencePath);

        await saveDocumentJsonHistory(contentData, { 
          id: docId, 
          prompt, 
          format, 
          mode: 'reference'
        });

        const buffer = await buildDocxFile(contentData);
        const safeFilename = (contentData.title || 'generated_from_reference')
          .toLowerCase()
          .replace(/[^a-z0-9]+/g, '_')
          .replace(/^_+|_+$/g, '') || 'generated_document';

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${safeFilename}.docx"`);
        res.setHeader('X-Document-Id', docId);
        res.setHeader('X-Mode', 'reference');
        return res.send(buffer);
      } catch (error) {
        if (fs.existsSync(referencePath)) fs.unlinkSync(referencePath);
        throw error;
      }
    }

    // MODE: FILL - Keep template structure, insert AI-generated content
    // Uses separate fillModeService to avoid conflicts with scratch mode
    if (mode === 'fill' && format === 'docx') {
      if (!req.file) {
        return res.status(400).json({ error: 'Template file is required for fill mode' });
      }

      const templatePath = req.file.path;
      const docId = generateDocumentId();

      try {
        // Step 1: Analyze template
        const templateAnalysis = await parseTemplate(templatePath);
        
        // Detect if user requested specific page count
        const pageMatch = prompt.toLowerCase().match(/\b(\d+)\s*(?:-| )?page/);
        const requestedPages = pageMatch ? parseInt(pageMatch[1], 10) : null;
        
        logger.info(`📄 TEMPLATE UPLOADED:`);
        logger.info(`   - Size: ${templateAnalysis.textLength} characters`);
        logger.info(`   - Pages: ${templateAnalysis.pages || 'Unknown'}`);
        logger.info(`   - Structure: ${templateAnalysis.structure.headingCount} headings`);
        logger.info(`   - Visual elements: ${templateAnalysis.structure.imageCount} images, ${templateAnalysis.structure.tableCount} tables`);
        logger.info(`   - Formatting: ${templateAnalysis.structure.hasBoldText ? 'Bold' : ''} ${templateAnalysis.structure.hasItalicText ? 'Italic' : ''}`);
        logger.info(`   - Colors: ${templateAnalysis.structure.colors.length > 0 ? templateAnalysis.structure.colors.join(', ') : 'None'}`);
        if (requestedPages) {
          logger.info(`   - Requested: ${requestedPages} pages (will expand structure)`);
        } else {
          logger.info(`   - Mode: Fill template as-is`);
        }
        
        // Step 2: Use dedicated fill mode service (separate from scratch)
        logger.info(`🤖 Fill Mode: Analyzing structure and generating matching content...`);
        const contentData = await generateFillModeDocument(
          templateAnalysis.fullText, 
          templateAnalysis,
          prompt,
          requestedPages
        );
        contentData.id = docId;
        
        if (fs.existsSync(templatePath)) fs.unlinkSync(templatePath);
        
        await saveDocumentJsonHistory(contentData, { 
          id: docId, 
          prompt, 
          format, 
          mode: 'fill',
          headingsInTemplate: templateAnalysis.structure.headingCount,
          sectionsGenerated: contentData.sections?.length || 0
        });
        
        const buffer = await buildFillModeDocx(contentData);
        const safeFilename = (contentData.title || 'filled_document')
          .toLowerCase()
          .replace(/[^a-z0-9]+/g, '_')
          .replace(/^_+|_+$/g, '') || 'filled_document';
        
        logger.info(`✅ Fill Mode Complete: ${contentData.sections?.length || 0} sections generated (AI determined optimal structure)`);
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${safeFilename}.docx"`);
        res.setHeader('X-Document-Id', docId);
        res.setHeader('X-Mode', 'fill');
        res.setHeader('X-Sections-Generated', (contentData.sections?.length || 0).toString());
        res.setHeader('X-Template-Headings', templateAnalysis.structure.headingCount.toString());
        return res.send(buffer);
      } catch (error) {
        if (fs.existsSync(templatePath)) fs.unlinkSync(templatePath);
        throw error;
      }
    }

    // MODE: SCRATCH (generate from scratch)
    if (format === 'docx') {
      const docId = generateDocumentId();
      const contentData = await generateDocumentContent(prompt, req.file);
      contentData.id = docId;

      await saveDocumentJsonHistory(contentData, { id: docId, prompt, format, mode });

      const buffer = await buildDocxFile(contentData);
      const safeFilename = (contentData.title || 'generated_document')
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '_')
        .replace(/^_+|_+$/g, '') || 'generated_document';

      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition', `attachment; filename="${safeFilename}.docx"`);
      res.setHeader('X-Document-Id', docId);
      return res.send(buffer);
    }

    if (format === 'pptx') {
      const docId = generateDocumentId();
      const templatePath = req.file ? req.file.path : null;

      try {
        const buffer = await buildPptxFile(prompt, mode, templatePath);

        if (templatePath && fs.existsSync(templatePath)) {
          fs.unlinkSync(templatePath);
        }

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
        res.setHeader('Content-Disposition', `attachment; filename="presentation_${docId}.pptx"`);
        res.setHeader('X-Document-Id', docId);
        return res.send(buffer);
      } catch (err) {
        if (templatePath && fs.existsSync(templatePath)) {
          fs.unlinkSync(templatePath);
        }
        throw err;
      }
    }

    return res.status(501).json({ 
      message: `Unified generation endpoint skeleton for format: ${format}`,
      format,
      mode
    });
  } catch (error) {
    logger.error('Error in handleGenerateDocument:', error.message);
    return res.status(500).json({ error: error.message });
  }
};

/**
 * Dedicated DOCX endpoint handler
 */
export const handleGenerateDocx = async (req, res) => {
  try {
    const { prompt } = req.body;
    logger.info(`DOCX Request Received -> Prompt: "${prompt}"`);

    if (!prompt) {
      logger.warn('DOCX Request Failed: Prompt is missing');
      return res.status(400).json({ error: 'Prompt is required' });
    }

    const docId = generateDocumentId();
    const contentData = await generateDocumentContent(prompt, req.file);
    contentData.id = docId;

    await saveDocumentJsonHistory(contentData, { id: docId, prompt, format: 'docx', mode: 'scratch' });

    const buffer = await buildDocxFile(contentData);
    const safeFilename = (contentData.title || 'generated_document')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '') || 'generated_document';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${safeFilename}.docx"`);
    res.setHeader('X-Document-Id', docId);
    return res.send(buffer);
  } catch (error) {
    logger.error('Error in handleGenerateDocx:', error.message);
    return res.status(500).json({ error: error.message });
  }
};

/**
 * Dedicated PPTX endpoint handler
 */
export const handleGeneratePptx = async (req, res) => {
  try {
    const { prompt } = req.body;
    if (!prompt) return res.status(400).json({ error: 'Prompt is required' });
    const docId = generateDocumentId();
    const buffer = await buildPptxFile(prompt, 'scratch');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="presentation_${docId}.pptx"`);
    res.setHeader('X-Document-Id', docId);
    return res.send(buffer);
  } catch (error) {
    logger.error('Error in handleGeneratePptx:', error.message);
    return res.status(500).json({ error: error.message });
  }
};

/**
 * Dedicated XLSX endpoint handler
 */
export const handleGenerateXlsx = async (req, res) => {
  try {
    const { prompt } = req.body;
    if (!prompt) return res.status(400).json({ error: 'Prompt is required' });
    const buffer = await buildXlsxFile({ prompt });
    return res.send(buffer);
  } catch (error) {
    logger.error('Error in handleGenerateXlsx:', error.message);
    return res.status(500).json({ error: error.message });
  }
};

/**
 * Dedicated PDF endpoint handler
 */
export const handleGeneratePdf = async (req, res) => {
  try {
    const { prompt } = req.body;
    if (!prompt) return res.status(400).json({ error: 'Prompt is required' });
    const buffer = await convertToPdf({ prompt });
    return res.send(buffer);
  } catch (error) {
    logger.error('Error in handleGeneratePdf:', error.message);
    return res.status(500).json({ error: error.message });
  }
};
