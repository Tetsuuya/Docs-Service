import fs from 'fs';
import path from 'path';
import https from 'https';
import { logger } from '../../utils/logger.js';
import { config } from '../../config/env.js';

/**
 * Hugging Face model pipeline — ordered by quality, falls back on error/loading
 * Note: Only models supported by HF's free hf-inference provider are listed.
 */
const HF_MODELS = [
  'stabilityai/stable-diffusion-3-medium-diffusers', // Confirmed working on hf-inference free tier
];

/**
 * Generate an image via Hugging Face Inference API (free tier).
 * Returns the image as a Buffer on success, or null on failure.
 *
 * @param {string} prompt  - The detailed image prompt from Gemini
 * @param {string} outputPath - File path to save the JPEG image
 * @param {number} modelIndex - Index into HF_MODELS (increments on retry)
 */
async function generateImageWithHuggingFace(prompt, outputPath, modelIndex = 0) {
  if (modelIndex >= HF_MODELS.length) {
    logger.warn('HF Image Generation: All models exhausted.');
    return null;
  }

  const modelId = HF_MODELS[modelIndex];
  const hfToken = config.hfToken;

  if (!hfToken) {
    logger.error('HF_TOKEN is not set in .env — cannot call Hugging Face Inference API.');
    return null;
  }

  const body = JSON.stringify({ inputs: prompt });
  const options = {
    hostname: 'router.huggingface.co',
    path: `/hf-inference/models/${modelId}`,
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${hfToken}`,
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(body),
      'x-wait-for-model': 'true',  // Auto-wait if model is loading (cold start)
    },
  };

  return new Promise((resolve) => {
    logger.info(`  [HF] Model: ${modelId}`);

    const req = https.request(options, (res) => {
      const chunks = [];

      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        const rawBody = Buffer.concat(chunks);

        if (res.statusCode === 200) {
          // Successful image response (binary)
          try {
            fs.writeFileSync(outputPath, rawBody);
            const size = fs.statSync(outputPath).size;
            if (size > 2000) {
              logger.info(`  [HF] ✅ Saved ${Math.round(size / 1024)}KB → ${outputPath}`);
              resolve(outputPath);
            } else {
              logger.warn(`  [HF] File too small (${size}B), trying next model...`);
              resolve(generateImageWithHuggingFace(prompt, outputPath, modelIndex + 1));
            }
          } catch (writeErr) {
            logger.warn(`  [HF] Write error: ${writeErr.message}`);
            resolve(generateImageWithHuggingFace(prompt, outputPath, modelIndex + 1));
          }
        } else if (res.statusCode === 503) {
          // Model still loading — x-wait-for-model should handle this, but retry once more
          logger.warn(`  [HF] 503 Model loading (${modelId}), switching model...`);
          resolve(generateImageWithHuggingFace(prompt, outputPath, modelIndex + 1));
        } else if (res.statusCode === 429) {
          logger.warn(`  [HF] 429 Rate limited on ${modelId}, switching model...`);
          resolve(generateImageWithHuggingFace(prompt, outputPath, modelIndex + 1));
        } else {
          const errText = rawBody.toString().substring(0, 200);
          logger.warn(`  [HF] HTTP ${res.statusCode} on ${modelId}: ${errText}`);
          resolve(generateImageWithHuggingFace(prompt, outputPath, modelIndex + 1));
        }
      });
    });

    req.on('error', (err) => {
      logger.warn(`  [HF] Request error on ${modelId}: ${err.message}`);
      resolve(generateImageWithHuggingFace(prompt, outputPath, modelIndex + 1));
    });

    req.on('timeout', () => {
      req.destroy();
      logger.warn(`  [HF] Timeout on ${modelId}, switching model...`);
      resolve(generateImageWithHuggingFace(prompt, outputPath, modelIndex + 1));
    });

    req.setTimeout(60000); // 60s — HF cold starts can be slow
    req.write(body);
    req.end();
  });
}

/**
 * Automatically generates AI images using FREE Pollinations AI
 * @param {Object|string} presentationDataOrTopic - Presentation JSON AST from Gemini
 * @returns {Promise<Object>} - Map of generated image file paths keyed by slide index
 */
export const generateTopicImages = async (presentationDataOrTopic) => {
  const isObject = typeof presentationDataOrTopic === 'object';
  const promptTopic = isObject ? (presentationDataOrTopic.title || presentationDataOrTopic.topic || 'presentation') : presentationDataOrTopic;

  logger.info(`🤖 Auto-Generating AI Images for: "${promptTopic}"...`);

  const tempImgDir = path.join(process.cwd(), 'temp', 'images');
  fs.mkdirSync(tempImgDir, { recursive: true });

  const resultImages = {};
  const slides = isObject && Array.isArray(presentationDataOrTopic.slides) ? presentationDataOrTopic.slides : [];

  // Only generate images for slides that request them
  const imageTasks = [];
  slides.forEach((slide, idx) => {
    if (slide.hasImage && slide.imagePrompt) {
      imageTasks.push({
        index: idx,
        prompt: slide.imagePrompt,
        file: `slide_${idx}.jpg`
      });
    }
  });

  logger.info(`  → ${imageTasks.length} slides requested images`);

  for (const task of imageTasks) {
    try {
      logger.info(`  → Slide ${task.index + 1}: "${task.prompt.substring(0, 60)}..."`);
      
      const imgPath = path.join(tempImgDir, task.file);
      
      // Use Hugging Face Inference API (free tier)
      const result = await generateImageWithHuggingFace(task.prompt, imgPath);
      
      if (result && fs.existsSync(imgPath) && fs.statSync(imgPath).size > 2000) {
        resultImages[task.index] = imgPath;
        logger.info(`     ✅ Generated (${Math.round(fs.statSync(imgPath).size / 1024)}KB)`);
      } else {
        logger.warn(`     ❌ Failed or too small`);
      }
    } catch (err) {
      logger.warn(`     ❌ Error: ${err.message}`);
    }
  }

  logger.info(`📸 Image Generation Complete: ${Object.keys(resultImages).length}/${imageTasks.length} successful`);
  return resultImages;
};
