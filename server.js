import app from './src/app.js';
import { config } from './src/config/env.js';
import { logger } from './src/utils/logger.js';

export const BUILD_VERSION = process.env.BUILD_VERSION || process.env.COMMIT_SHA || 'v1.0.1-release';

app.listen(config.port, () => {
  const maskedKey = config.geminiApiKey 
    ? `${config.geminiApiKey.substring(0, 6)}...${config.geminiApiKey.substring(config.geminiApiKey.length - 4)}` 
    : 'MISSING!';

  console.log(`====================================================`);
  console.log(`🚀 Docs Service Server running on port ${config.port}`);
  console.log(`📌 Build Version: \x1b[36m${BUILD_VERSION}\x1b[0m`);
  console.log(`🔗 Local URL: http://localhost:${config.port}`);
  if (config.geminiApiKey && config.geminiApiKey !== 'your_gemini_api_key_here') {
    logger.info(`Gemini API Key Loaded: \x1b[32mYES\x1b[0m (${maskedKey})`);
  } else {
    logger.warn(`Gemini API Key: \x1b[31mNOT CONFIGURED IN .env\x1b[0m`);
  }
  console.log(`====================================================`);
});
