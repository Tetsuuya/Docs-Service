import pptxgen from 'pptxgenjs';
import fs from 'fs';
import { logger } from '../../../utils/logger.js';

/**
 * INTELLIGENT SCRATCH MODE BUILDER
 * Builds presentations from AI-designed structure with dynamic layouts
 * @param {Object} presentationData - AI-generated presentation structure
 * @param {Object} imagePaths - Dynamic image paths keyed by slide index
 * @returns {Promise<Buffer>} - PPTX file buffer
 */
export const buildScratchPptx = async (presentationData, imagePaths = {}) => {
  logger.info(`🎨 Building Professional Presentation: "${presentationData.title || 'Untitled'}"`);

  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = presentationData.title || 'Presentation';
  pptx.author = 'Docs-Service AI Designer';

  // Extract AI-chosen theme
  const theme = presentationData.theme || {};
  const primaryColor = (theme.primaryColor || 'FFFFFF').replace('#', '');
  const secondaryColor = (theme.secondaryColor || 'CBD5E1').replace('#', '');
  const accentColor = (theme.accentColor || '38B6FF').replace('#', '');
  const backgroundColor = (theme.backgroundColor || '0F172A').replace('#', '');
  const cardBgColor = (theme.cardBgColor || '1E293B').replace('#', '');
  const fontFamily = theme.fontFamily || 'Poppins';

  const slides = presentationData.slides || [];
  
  // Track what gets built
  let imagesPlaced = 0;
  let imagesRequested = 0;
  let layoutCounts = {};

  logger.info(`📊 Build Plan:`);
  logger.info(`   → Total Slides: ${slides.length}`);
  logger.info(`   → Available Images: ${Object.keys(imagePaths).length}`);

  slides.forEach((slideData, index) => {
    const slide = pptx.addSlide();
    slide.background = { color: backgroundColor };

    const isHero = slideData.type === 'hero' || slideData.layout === 'hero' || index === 0;
    const layout = slideData.layout || 'list';
    
    // Track layouts
    layoutCounts[layout] = (layoutCounts[layout] || 0) + 1;
    
    // Get image if this slide requested one
    const slideImg = imagePaths[index] || null;
    const imageRequested = slideData.hasImage === true;
    const imageAvailable = slideImg && fs.existsSync(slideImg);
    
    if (imageRequested) imagesRequested++;
    if (imageAvailable) imagesPlaced++;
    
    logger.info(`   → Slide ${index + 1}: ${layout} | Image: ${imageRequested ? (imageAvailable ? '✅ placed' : '❌ missing') : '⚪ not requested'}`);

    // Top accent bar
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 0, y: 0, w: '100%', h: 0.08,
      fill: { color: accentColor }
    });

    if (isHero) {
      // ═══════════════════════════════════════════════════════════
      // HERO SLIDE (Title or Conclusion)
      // ═══════════════════════════════════════════════════════════
      
      // Topic badge
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: 0.5, y: 0.4, w: 3.5, h: 0.35,
        fill: { color: '334155' },
        line: { color: accentColor, width: 1.5 },
        rectRadius: 0.1
      });

      const tag = (presentationData.title || 'PRESENTATION').toUpperCase();
      slide.addText(`🏷️ ${tag}`, {
        x: 0.55, y: 0.42, w: 3.4, h: 0.3,
        fontSize: 10, bold: true, color: accentColor, fontFace: fontFamily, align: 'center'
      });

      // Main title
      slide.addText(slideData.title || presentationData.title || 'Presentation', {
        x: 0.5, y: 0.9, w: 5.5, h: 2.3,
        fontSize: 28, bold: true, color: primaryColor, fontFace: fontFamily, valign: 'middle', wrap: true
      });

      // Subtitle
      if (slideData.subtitle || presentationData.subtitle) {
        slide.addText(slideData.subtitle || presentationData.subtitle, {
          x: 0.5, y: 3.3, w: 5.5, h: 1.2,
          fontSize: 13, color: secondaryColor, fontFace: fontFamily, valign: 'top', wrap: true
        });
      }

      // Image frame
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: 5.8, y: 0.8, w: 3.7, h: 3.8,
        fill: { color: cardBgColor },
        line: { color: accentColor, width: 1.5 },
        rectRadius: 0.15
      });

      if (slideImg && fs.existsSync(slideImg)) {
        slide.addImage({
          path: slideImg,
          x: 5.88, y: 0.88, w: 3.54, h: 3.64
        });
      }

      // Footer
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 0.5, y: 5.0, w: 9.0, h: 0.03,
        fill: { color: '334155' }
      });

      slide.addText('DOCS-SERVICE AI DESIGN ENGINE 2026', {
        x: 0.5, y: 5.1, w: 9.0, h: 0.25,
        fontSize: 9, color: '94A3B8', fontFace: fontFamily
      });

    } else {
      // ═══════════════════════════════════════════════════════════
      // CONTENT SLIDES (Dynamic layouts based on AI choice)
      // ═══════════════════════════════════════════════════════════
      
      // Slide title
      slide.addText(slideData.title || `Section ${index + 1}`, {
        x: 0.5, y: 0.3, w: 9.0, h: 0.5,
        fontSize: 20, bold: true, color: primaryColor, fontFace: fontFamily
      });

      // Subtitle (optional)
      if (slideData.subtitle) {
        slide.addText(slideData.subtitle, {
          x: 0.5, y: 0.8, w: 9.0, h: 0.3,
          fontSize: 12, color: secondaryColor, fontFace: fontFamily
        });
      }

      // Footer
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 0.5, y: 5.1, w: 9.0, h: 0.02,
        fill: { color: '334155' }
      });

      slide.addText(`PRESENTATION | PAGE ${index + 1}`, {
        x: 0.5, y: 5.15, w: 9.0, h: 0.25,
        fontSize: 9, color: '94A3B8', fontFace: fontFamily
      });

      const contentTop = slideData.subtitle ? 1.25 : 1.0;
      const hasImage = slideImg && fs.existsSync(slideImg);

      // ─────────────────────────────────────────────────────────
      // LAYOUT: SPLIT (Image + Content)
      // ─────────────────────────────────────────────────────────
      if (hasImage && layout === 'split') {
        // Left: Image
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 0.5, y: contentTop, w: 4.1, h: 3.7,
          fill: { color: cardBgColor },
          line: { color: '334155', width: 1 },
          rectRadius: 0.15
        });
        
        slide.addImage({
          path: slideImg,
          x: 0.58, y: contentTop + 0.08, w: 3.94, h: 3.54
        });
        
        // Right: Bullets
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 4.8, y: contentTop, w: 4.7, h: 3.7,
          fill: { color: cardBgColor },
          line: { color: '334155', width: 1 },
          rectRadius: 0.15
        });
        
        slide.addShape(pptx.shapes.RECTANGLE, {
          x: 4.8, y: contentTop, w: 4.7, h: 0.08,
          fill: { color: accentColor }
        });
        
        const bullets = Array.isArray(slideData.bullets) ? slideData.bullets : (slideData.content ? [slideData.content] : ['Key insight']);
        const bulletItems = bullets.map(b => ({
          text: typeof b === 'string' ? b : (b.text || ''),
          options: { bullet: true, fontSize: 11, color: primaryColor, spaceAfter: 12, fontFace: fontFamily }
        }));
        
        slide.addText(bulletItems, {
          x: 5.0, y: contentTop + 0.25, w: 4.3, h: 3.3,
          valign: 'top', wrap: true
        });
      }
      
      // ─────────────────────────────────────────────────────────
      // LAYOUT: CARDS (Feature cards or pillars)
      // ─────────────────────────────────────────────────────────
      else if (layout === 'cards' && Array.isArray(slideData.cards) && slideData.cards.length > 0) {
        const cardCount = Math.min(slideData.cards.length, 3);
        const totalWidth = 9.0;
        const gap = 0.3;
        const cardWidth = (totalWidth - (gap * (cardCount - 1))) / cardCount;

        slideData.cards.slice(0, 3).forEach((card, cIdx) => {
          const cardX = 0.5 + cIdx * (cardWidth + gap);

          slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: cardX, y: contentTop, w: cardWidth, h: 3.7,
            fill: { color: cardBgColor },
            line: { color: '334155', width: 1 },
            rectRadius: 0.15
          });

          slide.addShape(pptx.shapes.RECTANGLE, {
            x: cardX, y: contentTop, w: cardWidth, h: 0.08,
            fill: { color: accentColor }
          });

          slide.addText(card.title || `Pillar ${cIdx + 1}`, {
            x: cardX + 0.15, y: contentTop + 0.25, w: cardWidth - 0.3, h: 0.5,
            fontSize: 13, bold: true, color: primaryColor, fontFace: fontFamily
          });

          slide.addText(card.description || '', {
            x: cardX + 0.15, y: contentTop + 0.8, w: cardWidth - 0.3, h: 2.7,
            fontSize: 10, color: secondaryColor, fontFace: fontFamily, valign: 'top', wrap: true
          });
        });
      }
      
      // ─────────────────────────────────────────────────────────
      // LAYOUT: STAT (Big number with context)
      // ─────────────────────────────────────────────────────────
      else if (layout === 'stat' && (slideData.statNumber || slideData.bullets)) {
        // Stat box
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 0.5, y: contentTop, w: 2.7, h: 3.7,
          fill: { color: cardBgColor },
          line: { color: accentColor, width: 1.5 },
          rectRadius: 0.15
        });

        if (slideData.statNumber) {
          slide.addText(slideData.statNumber, {
            x: 0.55, y: contentTop + 0.6, w: 2.6, h: 1.2,
            fontSize: 40, bold: true, color: accentColor, align: 'center', fontFace: fontFamily
          });
        }

        if (slideData.statLabel) {
          slide.addText(slideData.statLabel, {
            x: 0.55, y: contentTop + 1.9, w: 2.6, h: 1.4,
            fontSize: 11, bold: true, color: primaryColor, align: 'center', fontFace: fontFamily, wrap: true
          });
        }

        // Supporting content
        if (Array.isArray(slideData.bullets) && slideData.bullets.length > 0) {
          slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 3.4, y: contentTop, w: 6.1, h: 3.7,
            fill: { color: cardBgColor },
            line: { color: '334155', width: 1 },
            rectRadius: 0.15
          });

          const bulletItems = slideData.bullets.map(b => ({
            text: b,
            options: { bullet: true, fontSize: 11, color: primaryColor, spaceAfter: 12, fontFace: fontFamily }
          }));

          slide.addText(bulletItems, {
            x: 3.55, y: contentTop + 0.3, w: 5.8, h: 3.2,
            valign: 'top', wrap: true
          });
        }
      }
      
      // ─────────────────────────────────────────────────────────
      // LAYOUT: COMPARISON (Side-by-side A vs B)
      // ─────────────────────────────────────────────────────────
      else if (layout === 'comparison' && slideData.leftSide && slideData.rightSide) {
        const colWidth = 4.2;
        const gap = 0.6;
        
        // Left side
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 0.5, y: contentTop, w: colWidth, h: 3.7,
          fill: { color: cardBgColor },
          line: { color: '334155', width: 1 },
          rectRadius: 0.15
        });
        
        slide.addShape(pptx.shapes.RECTANGLE, {
          x: 0.5, y: contentTop, w: colWidth, h: 0.08,
          fill: { color: accentColor }
        });
        
        slide.addText(slideData.leftSide.title || 'Before', {
          x: 0.65, y: contentTop + 0.25, w: colWidth - 0.3, h: 0.5,
          fontSize: 14, bold: true, color: primaryColor, fontFace: fontFamily
        });
        
        const leftBullets = Array.isArray(slideData.leftSide.bullets) ? slideData.leftSide.bullets : [slideData.leftSide.description || ''];
        const leftItems = leftBullets.map(b => ({
          text: b,
          options: { bullet: true, fontSize: 10, color: secondaryColor, spaceAfter: 10, fontFace: fontFamily }
        }));
        
        slide.addText(leftItems, {
          x: 0.65, y: contentTop + 0.8, w: colWidth - 0.3, h: 2.7,
          valign: 'top', wrap: true
        });
        
        // Right side
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 0.5 + colWidth + gap, y: contentTop, w: colWidth, h: 3.7,
          fill: { color: cardBgColor },
          line: { color: accentColor, width: 2 },
          rectRadius: 0.15
        });
        
        slide.addShape(pptx.shapes.RECTANGLE, {
          x: 0.5 + colWidth + gap, y: contentTop, w: colWidth, h: 0.08,
          fill: { color: accentColor }
        });
        
        slide.addText(slideData.rightSide.title || 'After', {
          x: 0.65 + colWidth + gap, y: contentTop + 0.25, w: colWidth - 0.3, h: 0.5,
          fontSize: 14, bold: true, color: primaryColor, fontFace: fontFamily
        });
        
        const rightBullets = Array.isArray(slideData.rightSide.bullets) ? slideData.rightSide.bullets : [slideData.rightSide.description || ''];
        const rightItems = rightBullets.map(b => ({
          text: b,
          options: { bullet: true, fontSize: 10, color: secondaryColor, spaceAfter: 10, fontFace: fontFamily }
        }));
        
        slide.addText(rightItems, {
          x: 0.65 + colWidth + gap, y: contentTop + 0.8, w: colWidth - 0.3, h: 2.7,
          valign: 'top', wrap: true
        });
      }
      
      // ─────────────────────────────────────────────────────────
      // LAYOUT: TIMELINE / PROCESS (Sequential steps with badges)
      // ─────────────────────────────────────────────────────────
      else if ((layout === 'timeline' || layout === 'process') && Array.isArray(slideData.steps) && slideData.steps.length > 0) {
        const stepCount = Math.min(slideData.steps.length, 5);
        const stepHeight = 3.7 / stepCount;
        
        slideData.steps.slice(0, 5).forEach((step, idx) => {
          const stepY = contentTop + (idx * stepHeight);
          
          // Step container
          slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 0.5, y: stepY, w: 9.0, h: stepHeight - 0.1,
            fill: { color: cardBgColor },
            line: { color: '334155', width: 1 },
            rectRadius: 0.1
          });
          
          // Step number badge (circular)
          slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 0.7, y: stepY + (stepHeight/2) - 0.25, w: 0.5, h: 0.5,
            fill: { color: accentColor },
            rectRadius: 0.25  // Makes it circular
          });
          
          slide.addText(`${idx + 1}`, {
            x: 0.7, y: stepY + (stepHeight/2) - 0.25, w: 0.5, h: 0.5,
            fontSize: 14, bold: true, color: backgroundColor, fontFace: fontFamily,
            align: 'center', valign: 'middle'
          });
          
          // Step content
          const stepText = typeof step === 'string' ? step : (step.title || step.description || '');
          slide.addText(stepText, {
            x: 1.4, y: stepY + 0.1, w: 7.9, h: stepHeight - 0.2,
            fontSize: 11, color: primaryColor, fontFace: fontFamily,
            valign: 'middle', wrap: true
          });
        });
      }
      
      // ─────────────────────────────────────────────────────────
      // LAYOUT: QUOTE (Customer testimonial or impactful quote)
      // ─────────────────────────────────────────────────────────
      else if (layout === 'quote' && slideData.quote) {
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 1.5, y: contentTop + 0.5, w: 7.0, h: 2.5,
          fill: { color: cardBgColor },
          line: { color: accentColor, width: 2 },
          rectRadius: 0.2
        });

        // Large quote mark
        slide.addText('"', {
          x: 1.8, y: contentTop + 0.3, w: 0.5, h: 0.5,
          fontSize: 60, bold: true, color: accentColor, fontFace: fontFamily
        });

        // Quote text
        slide.addText(slideData.quote, {
          x: 2.0, y: contentTop + 0.9, w: 6.0, h: 1.5,
          fontSize: 16, italic: true, color: primaryColor, fontFace: fontFamily, valign: 'middle', wrap: true
        });

        // Attribution
        if (slideData.author) {
          slide.addText(`— ${slideData.author}`, {
            x: 2.0, y: contentTop + 2.6, w: 6.0, h: 0.3,
            fontSize: 12, color: secondaryColor, fontFace: fontFamily, align: 'right'
          });
        }
      }
      
      // ─────────────────────────────────────────────────────────
      // LAYOUT: MINIMAL (Single powerful sentence)
      // ─────────────────────────────────────────────────────────
      else if (layout === 'minimal' && slideData.statement) {
        // No container - just centered impactful text
        slide.addText(slideData.statement, {
          x: 1.0, y: contentTop + 1.0, w: 8.0, h: 2.0,
          fontSize: 32, bold: true, color: primaryColor, fontFace: fontFamily, 
          align: 'center', valign: 'middle', wrap: true
        });

        // Optional supporting text
        if (slideData.support) {
          slide.addText(slideData.support, {
            x: 1.5, y: contentTop + 3.2, w: 7.0, h: 0.5,
            fontSize: 14, color: secondaryColor, fontFace: fontFamily, 
            align: 'center', wrap: true
          });
        }
      }
      
      // ─────────────────────────────────────────────────────────
      // LAYOUT: TABLE (Data presentation)
      // ─────────────────────────────────────────────────────────
      else if (layout === 'table' && slideData.tableData) {
        const rows = slideData.tableData.rows || [];
        const headers = slideData.tableData.headers || [];
        
        if (headers.length > 0 && rows.length > 0) {
          const tableData = [
            headers.map(h => ({ text: h, options: { bold: true, color: primaryColor, fill: cardBgColor } })),
            ...rows.map(row => row.map(cell => ({ text: cell, options: { color: secondaryColor } })))
          ];

          slide.addTable(tableData, {
            x: 0.5, y: contentTop, w: 9.0, h: 3.5,
            fontSize: 11, fontFace: fontFamily,
            border: { pt: 1, color: '334155' },
            fill: { color: backgroundColor },
            align: 'left',
            valign: 'middle'
          });
        }
      }
      
      // ─────────────────────────────────────────────────────────
      // LAYOUT: LIST (Default bullet points)
      // ─────────────────────────────────────────────────────────
      else {
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 0.5, y: contentTop, w: 9.0, h: 3.7,
          fill: { color: cardBgColor },
          line: { color: '334155', width: 1 },
          rectRadius: 0.15
        });

        const bullets = Array.isArray(slideData.bullets) ? slideData.bullets : [slideData.content || ''];
        const bulletItems = bullets.map(b => ({
          text: typeof b === 'string' ? b : (b.text || ''),
          options: { bullet: true, fontSize: 11, color: primaryColor, spaceAfter: 12, fontFace: fontFamily }
        }));

        slide.addText(bulletItems, {
          x: 0.7, y: contentTop + 0.25, w: 8.6, h: 3.3,
          valign: 'top', wrap: true
        });
      }
    }

    // Speaker notes
    if (slideData.speakerNotes) {
      slide.addNotes(slideData.speakerNotes);
    }
  });

  const buffer = await pptx.write({ outputType: 'nodebuffer' });
  
  logger.info(`✅ Build Complete!`);
  logger.info(`   → Slides Generated: ${slides.length}`);
  logger.info(`   → Images Placed: ${imagesPlaced}/${imagesRequested} requested`);
  logger.info(`   → Layout Distribution: ${Object.entries(layoutCounts).map(([k, v]) => `${k}(${v})`).join(', ')}`);
  
  return buffer;
};
