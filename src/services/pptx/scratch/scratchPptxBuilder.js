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

  // Extract Theme
  const theme = presentationData.theme || {};
  const primaryColor = (theme.primaryColor || '0F172A').replace('#', '');
  const secondaryColor = (theme.secondaryColor || '2563EB').replace('#', '');
  const accentColor = (theme.accentColor || '38BDF8').replace('#', '');
  const backgroundColor = (theme.backgroundColor || 'F8FAFC').replace('#', '');
  const cardBgColor = (theme.cardBgColor || 'FFFFFF').replace('#', '');
  const textColor = (theme.textColor || '1E293B').replace('#', '');
  const fontFamily = theme.fontFamily || 'Calibri';

  const slides = presentationData.slides || [];

  slides.forEach((slideData, index) => {
    const slide = pptx.addSlide();
    slide.background = { color: backgroundColor };

    const isTitleSlide = slideData.type === 'title' || index === 0;

    // Slide Header / Accent Top Bar
    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 0,
      y: 0,
      w: '100%',
      h: 0.15,
      fill: { color: secondaryColor }
    });

    if (isTitleSlide) {
      // ----------------------------------------------------
      // TITLE SLIDE LAYOUT
      // ----------------------------------------------------
      // Decorative main title card background
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: 0.8,
        y: 1.5,
        w: 11.7,
        h: 4.5,
        fill: { color: cardBgColor },
        line: { color: accentColor, width: 1.5 },
        rectRadius: 0.2
      });

      // Accent left vertical bar inside card
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 0.8,
        y: 1.5,
        w: 0.25,
        h: 4.5,
        fill: { color: primaryColor }
      });

      // Title Text
      slide.addText(slideData.title || presentationData.title || 'Presentation', {
        x: 1.3,
        y: 2.2,
        w: 10.8,
        h: 1.5,
        fontSize: 36,
        bold: true,
        color: primaryColor,
        fontFace: fontFamily,
        valign: 'middle'
      });

      // Subtitle Text
      if (slideData.subtitle || presentationData.subtitle) {
        slide.addText(slideData.subtitle || presentationData.subtitle, {
          x: 1.3,
          y: 3.8,
          w: 10.8,
          h: 1.0,
          fontSize: 20,
          color: secondaryColor,
          fontFace: fontFamily,
          valign: 'top'
        });
      }
    } else {
      // ----------------------------------------------------
      // CONTENT SLIDES LAYOUT
      // ----------------------------------------------------
      // Footer text
      slide.addText(`${presentationData.title || ''}`, {
        x: 0.8,
        y: 7.1,
        w: 8.0,
        h: 0.3,
        fontSize: 10,
        color: '94A3B8',
        fontFace: fontFamily
      });

      // Slide Number Footer
      slide.addText(`Slide ${index + 1}`, {
        x: 11.0,
        y: 7.1,
        w: 1.5,
        h: 0.3,
        fontSize: 10,
        color: '94A3B8',
        align: 'right',
        fontFace: fontFamily
      });

      // Slide Title
      slide.addText(slideData.title || `Section ${index + 1}`, {
        x: 0.8,
        y: 0.4,
        w: 11.7,
        h: 0.8,
        fontSize: 26,
        bold: true,
        color: primaryColor,
        fontFace: fontFamily
      });

      // Slide Subtitle if present
      if (slideData.subtitle) {
        slide.addText(slideData.subtitle, {
          x: 0.8,
          y: 1.1,
          w: 11.7,
          h: 0.4,
          fontSize: 14,
          color: secondaryColor,
          fontFace: fontFamily
        });
      }

      const layoutType = slideData.layout || 'cards';
      const contentTop = slideData.subtitle ? 1.6 : 1.3;

      if (layoutType === 'cards' && Array.isArray(slideData.cards) && slideData.cards.length > 0) {
        // LAYOUT: CARDS GRID (2, 3, or 4 cards across)
        const cardCount = Math.min(slideData.cards.length, 4);
        const totalWidth = 11.7;
        const gap = 0.3;
        const cardWidth = (totalWidth - (gap * (cardCount - 1))) / cardCount;
        const cardHeight = 4.8;

        slideData.cards.slice(0, 4).forEach((card, cIdx) => {
          const cardX = 0.8 + cIdx * (cardWidth + gap);

          // Card background container
          slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: cardX,
            y: contentTop,
            w: cardWidth,
            h: cardHeight,
            fill: { color: cardBgColor },
            line: { color: 'E2E8F0', width: 1 },
            rectRadius: 0.15
          });

          // Card header bar
          slide.addShape(pptx.shapes.RECTANGLE, {
            x: cardX,
            y: contentTop,
            w: cardWidth,
            h: 0.1,
            fill: { color: secondaryColor }
          });

          // Card Title
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

          // Card Description
          slide.addText(card.description || '', {
            x: cardX + 0.2,
            y: contentTop + 1.0,
            w: cardWidth - 0.4,
            h: cardHeight - 1.2,
            fontSize: 12,
            color: textColor,
            fontFace: fontFamily,
            valign: 'top',
            wrap: true
          });
        });
      } else if (layoutType === 'stat' && (slideData.statNumber || slideData.bullets)) {
        // LAYOUT: STAT & CALLOUT
        // Big Stat Box (Left)
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 0.8,
          y: contentTop,
          w: 4.5,
          h: 4.8,
          fill: { color: primaryColor },
          rectRadius: 0.2
        });

        if (slideData.statNumber) {
          slide.addText(slideData.statNumber, {
            x: 1.0,
            y: contentTop + 1.0,
            w: 4.1,
            h: 1.4,
            fontSize: 54,
            bold: true,
            color: accentColor,
            align: 'center',
            fontFace: fontFamily
          });
        }

        if (slideData.statLabel) {
          slide.addText(slideData.statLabel, {
            x: 1.0,
            y: contentTop + 2.6,
            w: 4.1,
            h: 1.2,
            fontSize: 16,
            bold: true,
            color: 'FFFFFF',
            align: 'center',
            fontFace: fontFamily,
            wrap: true
          });
        }

        // Supporting Bullets Box (Right)
        if (Array.isArray(slideData.bullets) && slideData.bullets.length > 0) {
          slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
            x: 5.6,
            y: contentTop,
            w: 6.9,
            h: 4.8,
            fill: { color: cardBgColor },
            line: { color: 'E2E8F0', width: 1 },
            rectRadius: 0.2
          });

          const bulletItems = slideData.bullets.map(b => ({
            text: b,
            options: { bullet: true, fontSize: 14, color: textColor, spaceAfter: 12, fontFace: fontFamily }
          }));

          slide.addText(bulletItems, {
            x: 5.9,
            y: contentTop + 0.4,
            w: 6.3,
            h: 4.0,
            valign: 'top',
            wrap: true
          });
        }
      } else {
        // LAYOUT: BULLETS (Default)
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
          x: 0.8,
          y: contentTop,
          w: 11.7,
          h: 4.8,
          fill: { color: cardBgColor },
          line: { color: 'E2E8F0', width: 1 },
          rectRadius: 0.2
        });

        const bullets = Array.isArray(slideData.bullets) ? slideData.bullets : [slideData.content || ''];
        const bulletItems = bullets.map(b => ({
          text: typeof b === 'string' ? b : (b.text || ''),
          options: { bullet: true, fontSize: 15, color: textColor, spaceAfter: 14, fontFace: fontFamily }
        }));

        slide.addText(bulletItems, {
          x: 1.2,
          y: contentTop + 0.5,
          w: 10.9,
          h: 3.8,
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
