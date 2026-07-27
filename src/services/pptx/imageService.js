import fs from 'fs';
import path from 'path';
import { execSync } from 'child_process';
import { logger } from '../../utils/logger.js';

/**
 * Helper to construct an exact, highly relevant image prompt matching a slide's specific title & content
 */
function buildSlideImagePrompt(slide, topic, defaultRole) {
  if (slide?.imagePrompt && slide.imagePrompt.length > 10) {
    return slide.imagePrompt;
  }
  const title = slide?.title || topic || 'presentation';
  const snippet = Array.isArray(slide?.bullets) && slide.bullets.length > 0 ? (typeof slide.bullets[0] === 'string' ? slide.bullets[0] : slide.bullets[0].text || '') : '';
  const cleanSnippet = snippet.replace(/[^\w\s]/gi, '').slice(0, 80);
  
  return `3D cinematic presentation graphic of ${title}${cleanSnippet ? ', ' + cleanSnippet : ''}, dark executive theme`;
}

/**
 * Automatically generates AI images on-the-fly for slides that request them
 * @param {Object|string} presentationDataOrTopic - Presentation JSON AST from Gemini
 * @returns {Promise<Object>} - Map of generated image file paths keyed by slide index
 */
export const generateTopicImages = async (presentationDataOrTopic) => {
  const isObject = typeof presentationDataOrTopic === 'object';
  const promptTopic = isObject ? (presentationDataOrTopic.title || presentationDataOrTopic.topic || 'presentation') : presentationDataOrTopic;

  logger.info(`🤖 Auto-Generating AI Images for: "${promptTopic}"...`);

  const tempImgDir = path.join(process.cwd(), 'temp', 'images');
  fs.mkdirSync(tempImgDir, { recursive: true });

  const helperScript = path.join(process.cwd(), 'temp', 'test_image_generation.py');
  const resultImages = {};

  const slides = isObject && Array.isArray(presentationDataOrTopic.slides) ? presentationDataOrTopic.slides : [];

  // DYNAMIC: Only generate images for slides that request them
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
      const cmd = `python "${helperScript}" "${task.prompt}" "slide_${task.index}"`;
      execSync(cmd, { encoding: 'utf8', timeout: 45000 }); // 45 seconds (industry standard for API calls)

      const imgPath = path.join(tempImgDir, task.file);
      if (fs.existsSync(imgPath) && fs.statSync(imgPath).size > 2000) {
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
