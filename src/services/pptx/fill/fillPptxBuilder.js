import fs from 'fs';
import path from 'path';
import { execSync } from 'child_process';
import pptxgen from 'pptxgenjs';
import { logger } from '../../../utils/logger.js';

/**
 * Phase 4: Template Fill Engine & Slide Assembly
 * Uses Direct Run Text Overwriting for native PPTX master decks or High-Fidelity Vector Card Builder for PDF templates
 */
export const buildFillPptx = async (fillPlan, templateBlueprint, masterFilePath = null) => {
  logger.info(`Phase 4: Assembling High-Fidelity Presentation for: "${fillPlan.presentationTitle || 'Untitled Presentation'}"`);

  const isPdfMaster = masterFilePath && path.extname(masterFilePath).toLowerCase() === '.pdf' && fs.existsSync(masterFilePath);
  const isPptxMaster = masterFilePath && path.extname(masterFilePath).toLowerCase() === '.pptx' && fs.existsSync(masterFilePath);

  // OPTION A: Native PPTX Master Direct Run Overwriting (Preserves 100% of Canva fonts, colors, and layout positions)
  if (isPptxMaster) {
    try {
      logger.info(`Executing Direct PPTX Run Text Overwrite on: "${masterFilePath}"`);
      const tempDir = path.join(process.cwd(), 'temp');
      fs.mkdirSync(tempDir, { recursive: true });

      const planJsonPath = path.join(tempDir, `fill_plan_${Date.now()}.json`);
      fs.writeFileSync(planJsonPath, JSON.stringify(fillPlan, null, 2));

      const outputPptxPath = path.join(tempDir, `filled_direct_${Date.now()}.pptx`);
      const helperScript = path.join(process.cwd(), 'src', 'services', 'pptx', 'fill', 'directPptxReplacer.py');

      const cmd = `python "${helperScript}" "${masterFilePath}" "${planJsonPath}" "${outputPptxPath}"`;
      execSync(cmd, { encoding: 'utf8' });

      if (fs.existsSync(outputPptxPath)) {
        const buffer = fs.readFileSync(outputPptxPath);
        fs.unlinkSync(outputPptxPath);
        fs.unlinkSync(planJsonPath);

        logger.info(`✅ Direct PPTX Master overwrite complete (${buffer.length} bytes)`);
        return buffer;
      }
    } catch (err) {
      logger.warn(`Direct PPTX Processing warning: ${err.message}`);
    }
  }

  // OPTION B: High-Fidelity Vector Card Builder using Gemini Vision Brand Theme
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = fillPlan.presentationTitle || 'Filled Presentation';
  pptx.author = 'Docs-Service AI Template Engine';

  const brandTheme = templateBlueprint.brandTheme || {};
  const primaryColor = (brandTheme.primaryColor || '0F172A').replace('#', '');
  const secondaryColor = (brandTheme.secondaryColor || '475569').replace('#', '');
  const accentColor = (brandTheme.accentColor || '2563EB').replace('#', '');
  const cardBgColor = 'F8FAFC';
  const fontFamily = brandTheme.fontFamily || 'Montserrat';

  (fillPlan.selectedSlides || []).forEach((slideData, idx) => {
    const slide = pptx.addSlide();
    
    // Top Accent Brand Header Line
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 0, y: 0, w: 13.333, h: 0.15,
      fill: { color: accentColor }
    });

    const fill = slideData.fillContent || {};
    const isTitleSlide = slideData.layoutCategory === 'title_slide' || idx === 0;

    if (isTitleSlide) {
      // Main Cover Card Container
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: 0.8, y: 1.2, w: 11.733, h: 5.2,
        fill: { color: cardBgColor },
        line: { color: 'E2E8F0', width: 1 }
      });

      const titleText = fill.main_title || fill.title || fillPlan.presentationTitle || 'Presentation Title';
      slide.addText(titleText, {
        x: 1.2, y: 2.0, w: 10.9, h: 1.8,
        fontSize: 38, bold: true, color: primaryColor, fontFace: fontFamily,
        align: 'left', valign: 'middle'
      });

      const subtitleText = fill.subtitle || fill.description || fillPlan.topic || '';
      if (subtitleText) {
        slide.addText(subtitleText, {
          x: 1.2, y: 3.9, w: 10.9, h: 1.2,
          fontSize: 20, color: accentColor, fontFace: fontFamily,
          align: 'left', valign: 'top'
        });
      }
    } else {
      // Section Header Title
      const slideTitle = fill.title || fill.heading || slideData.title || `Section ${idx + 1}`;
      slide.addText(slideTitle, {
        x: 0.8, y: 0.4, w: 11.733, h: 0.8,
        fontSize: 26, bold: true, color: primaryColor, fontFace: fontFamily
      });

      const category = slideData.layoutCategory || '3_column_cards';

      if (category === '3_column_cards' || fill.col1_title || fill.card_1_title) {
        const cards = [
          { title: fill.col1_title || fill.card_1_title || 'Pillar 1', body: fill.col1_body || fill.card_1_body || fill.col1 || '' },
          { title: fill.col2_title || fill.card_2_title || 'Pillar 2', body: fill.col2_body || fill.card_2_body || fill.col2 || '' },
          { title: fill.col3_title || fill.card_3_title || 'Pillar 3', body: fill.col3_body || fill.card_3_body || fill.col3 || '' }
        ];

        const cardWidth = 3.6;
        const gap = 0.45;
        const topY = 1.4;

        cards.forEach((card, cIdx) => {
          const cardX = 0.8 + cIdx * (cardWidth + gap);

          slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: cardX, y: topY, w: cardWidth, h: 5.2,
            fill: { color: cardBgColor },
            line: { color: 'CBD5E1', width: 1 }
          });

          slide.addShape(pptx.shapes.RECTANGLE, {
            x: cardX, y: topY, w: cardWidth, h: 0.08,
            fill: { color: accentColor }
          });

          slide.addText(card.title, {
            x: cardX + 0.2, y: topY + 0.3, w: cardWidth - 0.4, h: 0.6,
            fontSize: 16, bold: true, color: primaryColor, fontFace: fontFamily
          });

          slide.addText(card.body, {
            x: cardX + 0.2, y: topY + 1.0, w: cardWidth - 0.4, h: 3.8,
            fontSize: 13, color: secondaryColor, fontFace: fontFamily, valign: 'top', wrap: true
          });
        });
      } else {
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 0.8, y: 1.4, w: 11.733, h: 5.2,
          fill: { color: cardBgColor },
          line: { color: 'E2E8F0', width: 1 }
        });

        const textEntries = Object.entries(fill)
          .filter(([k]) => k !== 'title' && k !== 'heading' && k !== 'main_title')
          .map(([, v]) => (typeof v === 'string' ? v : JSON.stringify(v)));

        const bulletItems = (textEntries.length > 0 ? textEntries : ['Key presentation takeaway point']).map(txt => ({
          text: txt,
          options: { bullet: true, fontSize: 15, color: primaryColor, spaceAfter: 14, fontFace: fontFamily }
        }));

        slide.addText(bulletItems, {
          x: 1.1, y: 1.7, w: 11.1, h: 4.6,
          valign: 'top', wrap: true
        });
      }
    }

    if (slideData.speakerNotes) {
      slide.addNotes(slideData.speakerNotes);
    }
  });

  const buffer = await pptx.write({ outputType: 'nodebuffer' });
  return buffer;
};
