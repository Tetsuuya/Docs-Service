import { generateDocumentContent } from '../services/geminiService.js';
import { buildDocxFile } from '../services/docxService.js';
import { buildPptxFile } from '../services/pptxService.js';
import { buildXlsxFile } from '../services/xlsxService.js';
import { convertToPdf } from '../services/pdfService.js';
import { logger } from '../utils/logger.js';

/**
 * Main unified controller endpoint handling format routing (DOCX, PPTX, XLSX, PDF)
 */
export const handleGenerateDocument = async (req, res) => {
  try {
    const { prompt, format = 'docx', mode = 'scratch' } = req.body;
    logger.info(`Unified Request Received -> Format: ${format}, Mode: ${mode}, Prompt: "${prompt}"`);

    if (!prompt) {
      logger.warn('Unified Request Failed: Prompt is missing');
      return res.status(400).json({ error: 'Prompt is required' });
    }
    
    return res.status(501).json({ 
      message: 'Unified document generation endpoint skeleton',
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

    const buffer = await buildDocxFile({ prompt });
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
    logger.info(`PPTX Request Received -> Prompt: "${prompt}"`);

    if (!prompt) {
      logger.warn('PPTX Request Failed: Prompt is missing');
      return res.status(400).json({ error: 'Prompt is required' });
    }

    const buffer = await buildPptxFile({ prompt });
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
    logger.info(`XLSX Request Received -> Prompt: "${prompt}"`);

    if (!prompt) {
      logger.warn('XLSX Request Failed: Prompt is missing');
      return res.status(400).json({ error: 'Prompt is required' });
    }

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
    logger.info(`PDF Request Received -> Prompt: "${prompt}"`);

    if (!prompt) {
      logger.warn('PDF Request Failed: Prompt is missing');
      return res.status(400).json({ error: 'Prompt is required' });
    }

    const buffer = await convertToPdf({ prompt });
    return res.send(buffer);
  } catch (error) {
    logger.error('Error in handleGeneratePdf:', error.message);
    return res.status(500).json({ error: error.message });
  }
};
