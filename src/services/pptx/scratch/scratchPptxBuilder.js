import pptxgen from 'pptxgenjs';
import { logger } from '../../../utils/logger.js';

/**
 * Builds a PowerPoint (.pptx) file buffer from structured presentation JSON data
 * @param {Object} presentationData - Structured JSON containing title, theme, and slides array
 * @returns {Promise<Buffer>} - Node.js Buffer containing the generated PPTX binary data
 */
export const buildScratchPptx = async (presentationData) => {
  logger.info(`Building PPTX file for: "${presentationData.title || 'Untitled Presentation'}"`);

  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';

  // Metadata
  pptx.title = presentationData.title || 'Presentation';
  pptx.author = 'Docs-Service AI';

  // Extract Theme (AGENCE DÉCLIC Style: Clean white BG with dark text)
  const theme = presentationData.theme || {};
  const primaryColor = (theme.primaryColor || '071E3D').replace('#', '');      // Dark navy primary text
  const secondaryColor = (theme.secondaryColor || '0070FF').replace('#', '');  // Bright blue accent
  const accentColor = (theme.accentColor || '489EF9').replace('#', '');        // Light blue secondary
  const backgroundColor = (theme.backgroundColor || 'FFFFFF').replace('#', ''); // Clean white background
  const cardBgColor = (theme.cardBgColor || 'F8F9FA').replace('#', '');        // Subtle gray card BG
  const textColor = (theme.textColor || '032853').replace('#', '');            // Dark body text
  const fontFamily = theme.fontFamily || 'Inter';

  const slides = presentationData.slides || [];

  slides.forEach((slideData, index) => {
    const slide = pptx.addSlide();
    slide.background = { color: backgroundColor };

    const isTitleSlide = slideData.type === 'title' || index === 0;

    // Slide Header / Accent Top Bar (AGENCE style - thinner accent line)
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: '100%',
      h: 0.08,
      fill: { color: secondaryColor }
    });

    if (isTitleSlide) {
      // ----------------------------------------------------
      // TITLE SLIDE LAYOUT (AGENCE DÉCLIC Style)
      // ----------------------------------------------------
      // Main Title - Large, bold, dark navy
      slide.addText(slideData.title || presentationData.title || 'Presentation', {
        x: 1.0,
        y: 2.5,
        w: 8.0,
        h: 2.0,
        fontSize: 48,
        bold: true,
        color: primaryColor,
        fontFace: fontFamily,
        valign: 'middle',
        wrap: true
      });

      // Subtitle - Accent color highlight
      if (slideData.subtitle || presentationData.subtitle) {
        slide.addText(slideData.subtitle || presentationData.subtitle, {
          x: 1.0,
          y: 4.7,
          w: 8.0,
          h: 1.0,
          fontSize: 22,
          color: secondaryColor,
          fontFace: fontFamily,
          valign: 'top',
          wrap: true
        });
      }
    } else {
      // ----------------------------------------------------
      // CONTENT SLIDES LAYOUT (AGENCE DÉCLIC Style)
      // ----------------------------------------------------
      // Footer text (simple, subtle)
      slide.addText(` ${presentationData.title || ''}`, {
        x: 0.5,
        y: 7.0,
        w: 8.0,
        h: 0.3,
        fontSize: 9,
        color: '94A3B8',
        fontFace: fontFamily
      });

      // Slide Number Footer
      slide.addText(`Page ${index + 1}`, {
        x: 11.0,
        y: 7.0,
        w: 1.5,
        h: 0.3,
        fontSize: 9,
        color: '94A3B8',
        align: 'right',
        fontFace: fontFamily
      });

      // Slide Title (Dark, bold, larger)
      slide.addText(slideData.title || `Section ${index + 1}`, {
        x: 0.5,
        y: 0.3,
        w: 12.0,
        h: 0.8,
        fontSize: 28,
        bold: true,
        color: primaryColor,
        fontFace: fontFamily
      });

      // Slide Subtitle if present (Accent color)
      if (slideData.subtitle) {
        slide.addText(slideData.subtitle, {
          x: 0.5,
          y: 1.1,
          w: 12.0,
          h: 0.5,
          fontSize: 16,
          color: secondaryColor,
          fontFace: fontFamily
        });
      }

      const layoutType = slideData.layout || 'cards';
      const contentTop = slideData.subtitle ? 1.8 : 1.5;

      if (layoutType === 'cards' && Array.isArray(slideData.cards) && slideData.cards.length > 0) {
        // LAYOUT: CARDS GRID (Clean, minimal borders)
        const cardCount = Math.min(slideData.cards.length, 3);
        const totalWidth = 12.0;
        const gap = 0.4;
        const cardWidth = (totalWidth - (gap * (cardCount - 1))) / cardCount;
        const cardHeight = 4.5;

        slideData.cards.slice(0, 3).forEach((card, cIdx) => {
          const cardX = 0.5 + cIdx * (cardWidth + gap);

          // Card background container (subtle)
          slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: cardX,
            y: contentTop,
            w: cardWidth,
            h: cardHeight,
            fill: { color: cardBgColor },
            line: { color: 'E0E0E0', width: 1 },
            rectRadius: 0.1
          });

          // Card accent top line
          slide.addShape(pptx.shapes.RECTANGLE, {
            x: cardX,
            y: contentTop,
            w: cardWidth,
            h: 0.06,
            fill: { color: secondaryColor }
          });

          // Card Title (Dark, bold)
          slide.addText(card.title || `Point ${cIdx + 1}`, {
            x: cardX + 0.2,
            y: contentTop + 0.3,
            w: cardWidth - 0.4,
            h: 0.6,
            fontSize: 16,
            bold: true,
            color: primaryColor,
            fontFace: fontFamily
          });

          // Card Description (Dark text on light BG)
          slide.addText(card.description || '', {
            x: cardX + 0.2,
            y: contentTop + 1.0,
            w: cardWidth - 0.4,
            h: cardHeight - 1.2,
            fontSize: 13,
            color: textColor,
            fontFace: fontFamily,
            valign: 'top',
            wrap: true
          });
        });
      } else if (layoutType === 'stat' && (slideData.statNumber || slideData.bullets)) {
        // LAYOUT: STAT & CALLOUT (AGENCE style - clean stat box)
        // Big Stat Box (Left) - Minimal, clean
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 0.5,
          y: contentTop,
          w: 4.0,
          h: 4.5,
          fill: { color: cardBgColor },
          line: { color: secondaryColor, width: 2 },
          rectRadius: 0.15
        });

        if (slideData.statNumber) {
          slide.addText(slideData.statNumber, {
            x: 0.7,
            y: contentTop + 1.0,
            w: 3.6,
            h: 1.5,
            fontSize: 60,
            bold: true,
            color: secondaryColor,
            align: 'center',
            fontFace: fontFamily
          });
        }

        if (slideData.statLabel) {
          slide.addText(slideData.statLabel, {
            x: 0.7,
            y: contentTop + 2.7,
            w: 3.6,
            h: 1.2,
            fontSize: 14,
            color: primaryColor,
            align: 'center',
            fontFace: fontFamily,
            wrap: true
          });
        }

        // Supporting Bullets Box (Right)
        if (Array.isArray(slideData.bullets) && slideData.bullets.length > 0) {
          slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 4.8,
            y: contentTop,
            w: 7.7,
            h: 4.5,
            fill: { color: cardBgColor },
            line: { color: 'E0E0E0', width: 1 },
            rectRadius: 0.15
          });

          const bulletItems = slideData.bullets.map(b => ({
            text: b,
            options: { bullet: true, fontSize: 14, color: textColor, spaceAfter: 14, fontFace: fontFamily }
          }));

          slide.addText(bulletItems, {
            x: 5.2,
            y: contentTop + 0.5,
            w: 7.0,
            h: 3.5,
            valign: 'top',
            wrap: true
          });
        }
      } else {
        // LAYOUT: BULLETS (Default - clean list)
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 0.5,
          y: contentTop,
          w: 12.0,
          h: 4.5,
          fill: { color: cardBgColor },
          line: { color: 'E0E0E0', width: 1 },
          rectRadius: 0.15
        });

        const bullets = Array.isArray(slideData.bullets) ? slideData.bullets : [slideData.content || ''];
        const bulletItems = bullets.map(b => ({
          text: typeof b === 'string' ? b : (b.text || ''),
          options: { bullet: true, fontSize: 15, color: textColor, spaceAfter: 16, fontFace: fontFamily }
        }));

        slide.addText(bulletItems, {
          x: 0.9,
          y: contentTop + 0.6,
          w: 11.2,
          h: 3.5,
          valign: 'top',
          wrap: true
        });
      }
    }

    // Speaker Notes
    if (slideData.speakerNotes) {
      slide.addNotes(slideData.speakerNotes);
    }
  });

  // Export buffer
  const buffer = await pptx.write({ outputType: 'nodebuffer' });
  return buffer;
};
