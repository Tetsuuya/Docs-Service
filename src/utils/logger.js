/**
 * Enterprise Logger Utility for Docs-Service
 */
const formatTimestamp = () => new Date().toISOString();

export const logger = {
  info: (message, meta = '') => {
    console.log(`\x1b[36m[INFO]\x1b[0m [${formatTimestamp()}] ${message}`, meta ? JSON.stringify(meta) : '');
  },
  warn: (message, meta = '') => {
    console.warn(`\x1b[33m[WARN]\x1b[0m [${formatTimestamp()}] ${message}`, meta ? JSON.stringify(meta) : '');
  },
  error: (message, meta = '') => {
    console.error(`\x1b[31m[ERROR]\x1b[0m [${formatTimestamp()}] ${message}`, meta ? JSON.stringify(meta) : '');
  }
};

/**
 * Express Middleware for logging incoming HTTP requests
 */
export const requestLogger = (req, res, next) => {
  const start = Date.now();
  res.on('finish', () => {
    const duration = Date.now() - start;
    const statusColor = res.statusCode >= 500 ? '\x1b[31m' : res.statusCode >= 400 ? '\x1b[33m' : '\x1b[32m';
    console.log(
      `${statusColor}[${req.method}]\x1b[0m [${formatTimestamp()}] ${req.originalUrl} - Status: ${res.statusCode} (${duration}ms)`
    );
  });
  next();
};
