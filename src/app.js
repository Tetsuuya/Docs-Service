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
app.use(cors());
app.use(express.json());

// Request Logging Middleware
app.use(requestLogger);

// Serve static frontend files from public folder
app.use(express.static(path.join(__dirname, '../public')));

// API Routes
app.use('/api', documentRoutes);

export default app;
