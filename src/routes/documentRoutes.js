import { Router } from 'express';
import multer from 'multer';
import { 
  handleGenerateDocument, 
  handleGenerateDocx, 
  handleGeneratePptx, 
  handleGenerateXlsx, 
  handleGeneratePdf 
} from '../controllers/documentController.js';

const router = Router();
const upload = multer({ dest: 'temp/uploads/' });

// Health check endpoint
router.get('/health', (req, res) => {
  res.json({ status: 'ok', service: 'Docs-Service', timestamp: new Date() });
});

// Main unified generation endpoint with optional context/design file upload
router.post('/generate', upload.single('contextFile'), handleGenerateDocument);

// Dedicated format sub-routes (Sub-server microservice endpoints)
router.post('/generate/docx', upload.single('contextFile'), handleGenerateDocx);
router.post('/generate/pptx', upload.single('contextFile'), handleGeneratePptx);
router.post('/generate/xlsx', upload.single('contextFile'), handleGenerateXlsx);
router.post('/generate/pdf', upload.single('contextFile'), handleGeneratePdf);

export default router;
