import { generatePptxContent } from './scratch/scratchPptxService.js';
import { buildScratchPptx } from './scratch/scratchPptxBuilder.js';
import { parsePresentationTemplate } from './template/templateParser.js';
import { generatePptxFillPlan } from './fill/pptxPlanGenerator.js';
import { buildFillPptx } from './fill/fillPptxBuilder.js';
import { buildReferenceInspiredPptx } from './reference/referencePptxBuilder.js';
import { generateTopicImages } from './imageService.js';
import { logger } from '../../utils/logger.js';
import fs from 'fs';

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
    
    // Step 1: Scan master template & build blueprint catalog
    const templateBlueprint = await parsePresentationTemplate(templateFilePath);
    
    // Step 2: AI Planner matches prompt topic to template placeholders
    const fillPlan = await generatePptxFillPlan(userPrompt, templateBlueprint);
    
    // Step 3: Auto-generate AI topic images for slides with image prompts
    let autoImages = imagePaths;
    const slidesWithImagePrompts = (fillPlan.selectedSlides || []).filter(s => s.imagePrompt);
    if (slidesWithImagePrompts.length > 0 && (!autoImages || Object.keys(autoImages).length === 0)) {
      const imagePayload = {
        title: fillPlan.presentationTitle || userPrompt,
        slides: fillPlan.selectedSlides.map(s => ({
          hasImage: !!s.imagePrompt,
          imagePrompt: s.imagePrompt
        }))
      };
      autoImages = await generateTopicImages(imagePayload);
    }
    
    // Step 4: Inject AI text + AI images into original template slides
    return await buildFillPptx(fillPlan, templateBlueprint, templateFilePath, autoImages);
  }

  if (mode === 'reference') {
    const userPrompt = typeof presentationDataOrPrompt === 'string' ? presentationDataOrPrompt : (presentationDataOrPrompt.prompt || 'Presentation');
    let templateBlueprint = null;
    if (templateFilePath && fs.existsSync(templateFilePath)) {
      templateBlueprint = await parsePresentationTemplate(templateFilePath);
    }

    const presentationData = await generatePptxContent(userPrompt);

    let autoImages = imagePaths;
    if (!autoImages || Object.keys(autoImages).length === 0) {
      autoImages = await generateTopicImages(presentationData);
    }

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
