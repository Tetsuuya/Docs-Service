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
 * Automatically generates AI images on-the-fly using Gemini-planned imagePrompts per slide
 * @param {Object|string} presentationDataOrTopic - Presentation JSON AST from Gemini or raw prompt topic
 * @returns {Promise<Object>} - Map of generated image file paths keyed by slide index or layout role
 */
export const generateTopicImages = async (presentationDataOrTopic) => {
  const isObject = typeof presentationDataOrTopic === 'object';
  const promptTopic = isObject ? (presentationDataOrTopic.title || presentationDataOrTopic.topic || 'presentation') : presentationDataOrTopic;

  logger.info(`🤖 Auto-Generating Exact Slide-Matched AI Images for presentation: "${promptTopic}"...`);

  const tempImgDir = path.join(process.cwd(), 'temp', 'images');
  fs.mkdirSync(tempImgDir, { recursive: true });

  const helperScript = path.join(process.cwd(), 'temp', 'test_image_generation.py');
  const resultImages = {};

  const slides = isObject && Array.isArray(presentationDataOrTopic.slides) ? presentationDataOrTopic.slides : [];

  // Construct exact slide-matched prompts
  const slide1Prompt = buildSlideImagePrompt(slides[0], promptTopic, 'hero cover');
  const slide2Prompt = buildSlideImagePrompt(slides[1], promptTopic, 'section 1 feature');
  const slide4Prompt = buildSlideImagePrompt(slides[3], promptTopic, 'section 2 deep dive');
  const slide5Prompt = buildSlideImagePrompt(slides[4], promptTopic, 'key performance metric');
  const slide6Prompt = buildSlideImagePrompt(slides[5], promptTopic, 'future outlook conclusion');

  const imageTasks = [
    { key: 'banner', prompt: slide1Prompt, file: 'banner.jpg' },
    { key: 'terrestrial', prompt: slide2Prompt, file: 'terrestrial.jpg' },
    { key: 'gasGiants', prompt: slide4Prompt, file: 'gasGiants.jpg' },
    { key: 'statVisual', prompt: slide5Prompt, file: 'statVisual.jpg' },
    { key: 'exploration', prompt: slide6Prompt, file: 'exploration.jpg' }
  ];

  for (const task of imageTasks) {
    try {
      logger.info(`  -> Generating Image [${task.key}] matching slide topic: "${task.prompt}"`);
      const cmd = `python "${helperScript}" "${task.prompt}" "${task.key}"`;
      execSync(cmd, { encoding: 'utf8' });

      const imgPath = path.join(tempImgDir, task.file);
      if (fs.existsSync(imgPath) && fs.statSync(imgPath).size > 2000) {
        resultImages[task.key] = imgPath;
      }
    } catch (err) {
      logger.warn(`Notice generating image for ${task.key}: ${err.message}`);
    }
  }

  logger.info(`✅ Slide-Matched AI Images successfully generated! (${Object.keys(resultImages).length} images)`);
  return resultImages;
};
