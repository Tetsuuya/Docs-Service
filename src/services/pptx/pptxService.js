import { generatePptxContent } from './scratch/scratchPptxService.js';
import { buildScratchPptx } from './scratch/scratchPptxBuilder.js';
import { parsePresentationTemplate } from './template/templateParser.js';
import { generatePptxFillPlan } from './fill/pptxPlanGenerator.js';
import { buildFillPptx } from './fill/fillPptxBuilder.js';
import { buildReferenceInspiredPptx } from './reference/referencePptxBuilder.js';
import { generateTopicImages } from './imageService.js';
import { logger } from '../../utils/logger.js';
import fs from 'fs';
import path from 'path';

/**
 * Master PPTX Service Endpoint
 * Handles PPTX generation across modes: 'scratch', 'fill', 'reference'
 */
export const buildPptxFile = async (presentationDataOrPrompt, mode = 'scratch', templateFilePath = null, imagePaths = {}) => {
  logger.info(`PPTX Service Called -> Mode: ${mode}${templateFilePath ? `, Template: ${templateFilePath}` : ''}`);

  if (mode === 'scratch') {
    let presentationData = presentationDataOrPrompt;
    if (typeof presentationDataOrPrompt === 'string') {
      presentationData = await generatePptxContent(presentationDataOrPrompt);
    }
    
    // Generate images for scratch mode (like reference does)
    let autoImages = imagePaths;
    if (!autoImages || Object.keys(autoImages).length === 0) {
      autoImages = await generateTopicImages(presentationData);
    }
    
    return await buildScratchPptx(presentationData, autoImages);
  }

  if (mode === 'fill') {
    if (!templateFilePath) {
      throw new Error(`Template file path is required for mode '${mode}'`);
    }
    const userPrompt = typeof presentationDataOrPrompt === 'string' ? presentationDataOrPrompt : (presentationDataOrPrompt.prompt || 'Presentation');
    const templateBlueprint = await parsePresentationTemplate(templateFilePath);
    const fillPlan = await generatePptxFillPlan(userPrompt, templateBlueprint);
    return await buildFillPptx(fillPlan, templateBlueprint, templateFilePath);
  }

  if (mode === 'reference') {
    const userPrompt = typeof presentationDataOrPrompt === 'string' ? presentationDataOrPrompt : (presentationDataOrPrompt.prompt || 'Presentation');
    let templateBlueprint = null;
    if (templateFilePath && fs.existsSync(templateFilePath)) {
      templateBlueprint = await parsePresentationTemplate(templateFilePath);
    }

    // Step 1: Call AI Planner (Gemini) to structure content and plan per-slide imagePrompts
    const presentationData = await generatePptxContent(userPrompt);

    // Step 2: Auto-generate custom AI images using Gemini's planned imagePrompts for each slide
    let autoImages = imagePaths;
    if (!autoImages || Object.keys(autoImages).length === 0) {
      autoImages = await generateTopicImages(presentationData);
    }

    // Inject extracted brand tokens from template reference if available
    if (templateBlueprint && templateBlueprint.brandTheme) {
      presentationData.theme = {
        primaryColor: templateBlueprint.brandTheme.primaryColor || '071E3D',
        secondaryColor: templateBlueprint.brandTheme.secondaryColor || '1E293B',
        accentColor: templateBlueprint.brandTheme.accentColor || '38B6FF',
        backgroundColor: '0B0E17',
        cardBgColor: '111827',
        textColor: 'FFFFFF',
        fontFamily: templateBlueprint.brandTheme.fontFamily || 'Poppins'
      };
    }

    return await buildReferenceInspiredPptx(presentationData, autoImages);
  }

  throw new Error(`Unsupported PPTX mode '${mode}'`);
};
