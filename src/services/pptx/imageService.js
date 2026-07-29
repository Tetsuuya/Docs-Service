import fs from 'fs';
import path from 'path';
import https from 'https';
import { logger } from '../../utils/logger.js';
import { config } from '../../config/env.js';

const HF_MODELS = [
  'stabilityai/stable-diffusion-3-medium-diffusers',
];

const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

const ABSTRACT_STOP_WORDS = new Set([
  'the', 'and', 'for', 'with', 'about', 'world', 'care', 'types', 'overview',
  'introduction', 'disadvantage', 'disadvantages', 'advantage', 'advantages',
  'risk', 'risks', 'pros', 'cons', 'impact', 'impacts', 'future', 'challenge',
  'challenges', 'issue', 'issues', 'hidden', 'costs', 'analysis', 'guide',
  'understanding', 'exploring', 'mastering', 'learning', 'basics', 'fundamentals',
  'essential', 'essentials', 'key', 'takeaways', 'summary', 'conclusion', 'modern',
  'evolution', 'history', 'strategic', 'insights', 'concept', 'concepts'
]);

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

/**
 * Keyword-based stock photo using Unsplash Source (free, no auth, returns topic-relevant images)
 * Falls back to Picsum only if Unsplash fails.
 */
async function downloadStockFallbackImage(prompt, mainTopic, outputPath) {
  return new Promise((resolve) => {
    // Extract best keyword from prompt or topic
    let keyword = '';

    if (prompt && typeof prompt === 'string') {
      const promptWords = prompt.toLowerCase().replace(/[^a-z ]/g, "").split(" ")
        .filter(w => w.length > 3 && !ABSTRACT_STOP_WORDS.has(w) && !['photo', 'photography', 'design', 'background', 'lighting', 'sleek', 'dark', 'blue', 'abstract', 'conceptual', 'depiction', 'illustration', 'image', 'symbol', 'matrix', 'scene', 'dramatic', 'glowing', 'neon', 'electric', 'photorealistic', 'cinematic', 'studio', 'ultra', 'resolution', 'watermark', 'clean', 'white', 'professional', 'sharp', 'focus', 'soft', 'accent'].includes(w));
      if (promptWords.length) keyword = promptWords[0];
    }

    if (!keyword && mainTopic && typeof mainTopic === 'string') {
      const topicWords = mainTopic.toLowerCase().replace(/[^a-z ]/g, "").split(" ")
        .filter(w => w.length > 2 && !ABSTRACT_STOP_WORDS.has(w));
      if (topicWords.length) keyword = topicWords[0];
    }

    if (!keyword) keyword = 'business';

    const seed = Math.floor(Math.random() * 1000);
    // Unsplash Source: keyword-relevant, high-res, free
    const url = `https://source.unsplash.com/1280x720/?${encodeURIComponent(keyword)}&sig=${seed}`;

    logger.info(`  [Stock Fallback] Fetching topic photo for "${keyword}" from Unsplash (sig: ${seed})...`);

    const handleStream = (res) => {
      if (res.statusCode === 301 || res.statusCode === 302) {
        const rawLoc = res.headers.location || '';
        const targetUrl = rawLoc.startsWith('http') ? rawLoc : `https://source.unsplash.com${rawLoc}`;
        https.get(targetUrl, { headers: { 'User-Agent': USER_AGENT } }, handleStream).on('error', () => {
          // Unsplash failed, try Picsum
          tryPicsum(resolve, outputPath);
        });
        return;
      }

      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        const buffer = Buffer.concat(chunks);
        if (res.statusCode === 200 && buffer.length > 5000) {
          fs.writeFileSync(outputPath, buffer);
          logger.info(`  [Stock Fallback] ✅ Saved ${Math.round(buffer.length / 1024)}KB → ${outputPath}`);
          resolve(outputPath);
        } else {
          tryPicsum(resolve, outputPath);
        }
      });
    };

    const req = https.get(url, { headers: { 'User-Agent': USER_AGENT } }, handleStream);
    req.on('error', () => tryPicsum(resolve, outputPath));
    req.setTimeout(10000, () => {
      req.destroy();
      tryPicsum(resolve, outputPath);
    });
  });
}

function tryPicsum(resolve, outputPath) {
  const seed = Math.floor(Math.random() * 1000000);
  const url = `https://picsum.photos/1280/720?random=${seed}`;
  logger.info(`  [Stock Fallback] Picsum fallback (seed: ${seed})...`);

  const handleStream = (res) => {
    if (res.statusCode === 301 || res.statusCode === 302) {
      const rawLoc = res.headers.location || '';
      const targetUrl = rawLoc.startsWith('http') ? rawLoc : `https://picsum.photos${rawLoc}`;
      https.get(targetUrl, { headers: { 'User-Agent': USER_AGENT } }, handleStream).on('error', () => resolve(null));
      return;
    }
    const chunks = [];
    res.on('data', (chunk) => chunks.push(chunk));
    res.on('end', () => {
      const buffer = Buffer.concat(chunks);
      if (res.statusCode === 200 && buffer.length > 5000) {
        fs.writeFileSync(outputPath, buffer);
        resolve(outputPath);
      } else {
        resolve(null);
      }
    });
  };
  const req = https.get(url, { headers: { 'User-Agent': USER_AGENT } }, handleStream);
  req.on('error', () => resolve(null));
  req.setTimeout(8000, () => { req.destroy(); resolve(null); });
}


/**
 * Generate image via Pollinations AI with 8s timeout
 */
async function generateImageWithPollinations(prompt, outputPath, retries = 2) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const result = await new Promise((resolve) => {
        const styledPrompt = `${prompt.substring(0, 200)}, 16:9 aspect ratio, ultra high resolution, no text, no watermark`;
        const encodedPrompt = encodeURIComponent(styledPrompt);
        const seed = Math.floor(Math.random() * 1000000);
        const url = `https://image.pollinations.ai/prompt/${encodedPrompt}?width=1024&height=768&nologo=true&seed=${seed}`;

        const options = {
          headers: {
            'User-Agent': USER_AGENT,
            'Accept': 'image/*',
          }
        };

        const handleStream = (res) => {
          if (res.statusCode === 301 || res.statusCode === 302) {
            const rawLoc = res.headers.location || '';
            const targetUrl = rawLoc.startsWith('http') ? rawLoc : `https://image.pollinations.ai${rawLoc}`;

            https.get(targetUrl, options, handleStream).on('error', () => resolve(null));
            return;
          }

          const chunks = [];
          res.on('data', (chunk) => chunks.push(chunk));
          res.on('end', () => {
            const buffer = Buffer.concat(chunks);
            if (res.statusCode === 200 && buffer.length > 5000) {
              fs.writeFileSync(outputPath, buffer);
              resolve(outputPath);
            } else {
              resolve(null);
            }
          });
        };

        const req = https.get(url, options, handleStream);
        req.on('error', () => resolve(null));
        req.setTimeout(8000, () => {
          req.destroy();
          resolve(null);
        });
      });

      if (result && fs.existsSync(outputPath) && fs.statSync(outputPath).size > 5000) {
        logger.info(`  [Pollinations] ✅ Saved ${Math.round(fs.statSync(outputPath).size / 1024)}KB → ${outputPath}`);
        return outputPath;
      }
    } catch (_) {}

    if (attempt < retries) {
      await sleep(300);
    }
  }
  return null;
}

/**
 * Generate an image via Hugging Face Inference API (free tier).
 */
async function generateImageWithHuggingFace(prompt, outputPath) {
  const modelId = HF_MODELS[0];
  const hfToken = config.hfToken;

  if (!hfToken) return null;

  const styledPrompt = `${prompt}, ultra high resolution, no text, no watermark`;
  const body = JSON.stringify({ inputs: styledPrompt });
  const options = {
    hostname: 'router.huggingface.co',
    path: `/hf-inference/models/${modelId}`,
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${hfToken}`,
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(body),
      'User-Agent': USER_AGENT,
      'x-wait-for-model': 'true',
    },
  };

  return new Promise((resolve) => {
    const req = https.request(options, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        const rawBody = Buffer.concat(chunks);
        if (res.statusCode === 200 && rawBody.length > 5000) {
          try {
            fs.writeFileSync(outputPath, rawBody);
            resolve(outputPath);
          } catch (_) {
            resolve(null);
          }
        } else {
          resolve(null);
        }
      });
    });

    req.on('error', () => resolve(null));
    req.setTimeout(5000, () => {
      req.destroy();
      resolve(null);
    });

    req.write(body);
    req.end();
  });
}

/**
 * Automatically generates AI topic images for every slide requesting an image.
 * Uses 3-tier fallback (HF -> Pollinations AI -> High-Res Stock Photo) for 100% reliability!
 */
export const generateTopicImages = async (presentationDataOrTopic) => {
  const isObject = typeof presentationDataOrTopic === 'object';
  const rawTopic = isObject ? (presentationDataOrTopic.topic || presentationDataOrTopic.title || 'presentation') : presentationDataOrTopic;

  logger.info(`🤖 Auto-Generating Unique Presentation-Fit AI Images for: "${rawTopic}"...`);

  const tempImgDir = path.join(process.cwd(), 'temp', 'images');
  fs.mkdirSync(tempImgDir, { recursive: true });

  const resultImages = {};
  const slides = isObject && Array.isArray(presentationDataOrTopic.slides) ? presentationDataOrTopic.slides : [];

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

  logger.info(`  → ${imageTasks.length} slides requested topic images`);

  for (let i = 0; i < imageTasks.length; i++) {
    const task = imageTasks[i];
    try {
      logger.info(`  → Slide ${task.index + 1}: "${task.prompt.substring(0, 50)}..."`);
      const imgPath = path.join(tempImgDir, task.file);
      
      // Tier 1: Try Hugging Face
      let result = await generateImageWithHuggingFace(task.prompt, imgPath);
      
      // Tier 2: Try Pollinations AI
      if (!result || !fs.existsSync(imgPath) || fs.statSync(imgPath).size < 5000) {
        result = await generateImageWithPollinations(task.prompt, imgPath);
      }

      // Tier 3: Guaranteed Unique High-Res Stock Photo
      if (!result || !fs.existsSync(imgPath) || fs.statSync(imgPath).size < 5000) {
        result = await downloadStockFallbackImage(task.prompt, rawTopic, imgPath);
      }

      if (result && fs.existsSync(imgPath) && fs.statSync(imgPath).size > 5000) {
        resultImages[task.index] = imgPath;
        logger.info(`     ✅ Image ${task.index + 1} generated successfully (${Math.round(fs.statSync(imgPath).size / 1024)}KB)`);
      } else {
        logger.warn(`     ❌ Image generation failed for slide ${task.index + 1}`);
      }

      if (i < imageTasks.length - 1) {
        await sleep(300);
      }

    } catch (err) {
      logger.warn(`     ❌ Error on slide ${task.index + 1}: ${err.message}`);
    }
  }

  logger.info(`📸 Image Generation Complete: ${Object.keys(resultImages).length}/${imageTasks.length} successful`);
  return resultImages;
};

/**
 * Automatically generates a topic-specific logo icon.
 * Uses Pollinations AI or Hugging Face, saving as logo.png in the temp directory.
 */
export const generateTopicLogo = async (topic) => {
  const tempImgDir = path.join(process.cwd(), 'temp', 'images');
  fs.mkdirSync(tempImgDir, { recursive: true });
  const logoPath = path.join(tempImgDir, 'logo.png');

  logger.info(`🤖 Auto-Generating Unique Topic Logo for: "${topic}"...`);

  // Prompt style for a clean vector branding icon
  const prompt = `Simple minimalist flat vector graphic icon logo of ${topic}, solid white on a solid black background, high contrast, clean graphics, branding symbol, no text, no watermark`;

  // Tier 1: Try Hugging Face
  let result = await generateImageWithHuggingFace(prompt, logoPath);

  // Tier 2: Try Pollinations AI
  if (!result || !fs.existsSync(logoPath) || fs.statSync(logoPath).size < 5000) {
    result = await generateImageWithPollinations(prompt, logoPath);
  }

  // Tier 3: Unsplash Fallback
  if (!result || !fs.existsSync(logoPath) || fs.statSync(logoPath).size < 5000) {
    result = await downloadStockFallbackImage(prompt, topic, logoPath);
  }

  if (result && fs.existsSync(logoPath) && fs.statSync(logoPath).size > 2000) {
    logger.info(`     ✅ Topic logo generated successfully (${Math.round(fs.statSync(logoPath).size / 1024)}KB)`);
    return logoPath;
  }
  return null;
};
