import express from 'express';
import cors from 'cors';
import path from 'path';
import { fileURLToPath } from 'url';
import documentRoutes from './routes/documentRoutes.js';
import { requestLogger } from './utils/logger.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
    
const app = express();

// Middleware
app.use(cors({ exposedHeaders: ['Content-Disposition', 'X-Document-Id'] }));
app.use(express.json());

// Request Logging Middleware
app.use(requestLogger);

// Serve static frontend files from public folder (disable serving index.html at root '/')
app.use(express.static(path.join(__dirname, '../public'), { index: false }));

// Serve frontend at /index and /index.html
app.get(['/index', '/index.html'], (req, res) => {
  res.sendFile(path.join(__dirname, '../public/index.html'));
});

// Health and Version status endpoint
app.get(['/health', '/api/version'], (req, res) => {
  const version = process.env.BUILD_VERSION || process.env.COMMIT_SHA || 'v1.0.1-release';
  res.json({
    status: 'ok',
    service: 'Docs-Service',
    version,
    timestamp: new Date().toISOString()
  });
});

// API Routes
app.use('/api', documentRoutes);

export default app;
