import { Router } from 'express';
import { 
  handleGenerateDocument, 
  handleGenerateDocx, 
  handleGeneratePptx, 
  handleGenerateXlsx, 
  handleGeneratePdf 
} from '../controllers/documentController.js';

const router = Router();

// Health check endpoint
router.get('/health', (req, res) => {
  res.json({ status: 'ok', service: 'Docs-Service', timestamp: new Date() });
});

// Main unified generation endpoint
router.post('/generate', handleGenerateDocument);

// Dedicated format sub-routes (Sub-server microservice endpoints)
router.post('/generate/docx', handleGenerateDocx);
router.post('/generate/pptx', handleGeneratePptx);
router.post('/generate/xlsx', handleGenerateXlsx);
router.post('/generate/pdf', handleGeneratePdf);

export default router;
