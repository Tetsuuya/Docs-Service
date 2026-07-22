import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import crypto from 'crypto';
import { logger } from './logger.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const HISTORY_DIR = path.join(__dirname, '../../temp/history');

// Ensure history directory exists
if (!fs.existsSync(HISTORY_DIR)) {
  fs.mkdirSync(HISTORY_DIR, { recursive: true });
}

/**
 * Generates a clean, unique Document ID (e.g. doc_1784733250000_a1b2c3)
 */
export const generateDocumentId = () => {
  const timestamp = Date.now();
  const randomHash = crypto.randomBytes(3).toString('hex');
  return `doc_${timestamp}_${randomHash}`;
};

/**
 * Saves generated Document JSON AST to local history storage (temp/history/).
 * Includes unique Document ID, prompt tracking, metadata, and documentAST.
 */
export const saveDocumentJsonHistory = async (documentData, metadata = {}) => {
  try {
    const docId = documentData.id || metadata.id || generateDocumentId();
    const safeTitle = (documentData.title || 'doc')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '_')
      .substring(0, 30);
    const fileName = `${docId}_${safeTitle}.json`;
    const filePath = path.join(HISTORY_DIR, fileName);

    const userPrompt = metadata.prompt || documentData.prompt || '';

    // Attach document ID & prompt inside documentAST
    documentData.id = docId;
    if (userPrompt) documentData.prompt = userPrompt;

    const payload = {
      id: docId,
      savedAt: new Date().toISOString(),
      prompt: userPrompt,
      metadata: {
        id: docId,
        prompt: userPrompt,
        format: metadata.format || 'docx',
        mode: metadata.mode || 'scratch'
      },
      documentAST: documentData
    };

    await fs.promises.writeFile(filePath, JSON.stringify(payload, null, 2), 'utf-8');
    logger.info(`Saved Document JSON AST to history: temp/history/${fileName} [ID: ${docId}]`);
    return { docId, filePath };
  } catch (err) {
    logger.warn(`Failed to save document JSON history: ${err.message}`);
    return null;
  }
};
