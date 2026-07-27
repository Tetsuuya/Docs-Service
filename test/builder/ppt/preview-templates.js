import { VISUAL_TEMPLATES } from '../../../src/services/pptx/scratch/visualTemplates.js';
import pptxgen from 'pptxgenjs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Generate visual preview of all 10 templates
 * Creates one slide per template showing the design
 */

async function generateTemplatePreviews() {
  console.log('═══════════════════════════════════════════════════════════');
  console.log('🎨 TEMPLATE VISUAL PREVIEW GENERATOR');
  console.log('═══════════════════════════════════════════════════════════\n');

  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = 'Visual Template Showcase';
  pptx.author = 'Docs-Service';

  const templateNames = Object.keys(VISUAL_TEMPLATES);

  console.log(`Generating ${templateNames.length} template previews...\n`);

  templateNames.forEach((templateName, idx) => {
    const template = VISUAL_TEMPLATES[templateName];
    const slide = pptx.addSlide();

    console.log(`[${idx + 1}/10] ${template.name}`);
    console.log(`  Features:`);

    // Apply background
    const bgColor = template.colors?.overrideBackground || '0F172A';
    slide.background = { color: bgColor };

    // Colors
    const primaryColor = template.colors?.overridePrimary || 'FFFFFF';
    const secondaryColor = template.colors?.overrideSecondary || 'CBD5E1';
    const accentColor = template.colors?.overrideAccent || '38B6FF';
    const cardBgColor = template.colors?.overrideCardBg || '1E293B';

    console.log(`    BG: #${bgColor}, Accent: #${accentColor}`);

    // Top bar (if template uses it)
    if (template.content?.showTopBar) {
      const barHeight = template.content.topBarHeight || 0.08;
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 0, y: 0, w: '100%', h: barHeight,
        fill: { color: accentColor }
      });
      console.log(`    Top bar: YES (${barHeight}in)`);
    } else {
      console.log(`    Top bar: NO`);
    }

    // Badge (if template uses it)
    if (template.content?.showBadge !== false) {
      const badgeX = template.content?.badgePosition?.x || 0.5;
      const badgeY = template.content?.badgePosition?.y || 0.3;
      
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: badgeX, y: badgeY, w: 2.5, h: 0.3,
        fill: { color: '334155' },
        line: { color: accentColor, width: 1 },
        rectRadius: 0.1
      });
      
      slide.addText(`🏷️ ${template.name.toUpperCase()}`, {
        x: badgeX, y: badgeY, w: 2.5, h: 0.3,
        fontSize: 10, bold: true, color: accentColor, align: 'center', valign: 'middle'
      });
      console.log(`    Badge: YES at (${badgeX}, ${badgeY})`);
    }

    // Title
    const titleX = template.content?.titlePosition?.x || 0.5;
    const titleY = template.content?.titlePosition?.y || 0.3;
    const titleAlign = template.content?.centered ? 'center' : 'left';
    
    slide.addText('Sample Slide Title', {
      x: titleX, y: titleY, w: 8.0, h: 0.5,
      fontSize: 22, bold: true, color: primaryColor, align: titleAlign
    });
    console.log(`    Title: (${titleX}, ${titleY}) ${titleAlign}-aligned`);

    // Subtitle
    const subtitleX = template.content?.subtitlePosition?.x || 0.5;
    const subtitleY = template.content?.subtitlePosition?.y || 0.8;
    
    slide.addText('This is how the subtitle appears in this template', {
      x: subtitleX, y: subtitleY, w: 8.0, h: 0.3,
      fontSize: 12, color: secondaryColor, align: titleAlign
    });

    // Content area demo
    const contentTop = template.content?.contentTop || 1.25;
    const cardRadius = template.content?.cardRadius || 0.15;
    const borderWidth = template.content?.cardBorderWidth || 1;
    
    console.log(`    Content starts: ${contentTop}in, Radius: ${cardRadius}, Border: ${borderWidth}px`);

    // Show 3 sample cards to demonstrate layout
    const cardWidth = 2.8;
    const gap = 0.3;
    
    for (let i = 0; i < 3; i++) {
      const cardX = 0.5 + i * (cardWidth + gap);
      
      slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
        x: cardX, y: contentTop, w: cardWidth, h: 3.0,
        fill: { color: cardBgColor },
        line: borderWidth > 0 ? { color: '334155', width: borderWidth } : { type: 'none' },
        rectRadius: cardRadius
      });
      
      if (template.content?.useAccentHeaders) {
        slide.addShape(pptx.shapes.RECTANGLE, {
          x: cardX, y: contentTop, w: cardWidth, h: 0.08,
          fill: { color: accentColor }
        });
      }
      
      slide.addText(`Card ${i + 1}`, {
        x: cardX + 0.15, y: contentTop + 0.25, w: cardWidth - 0.3, h: 0.4,
        fontSize: 13, bold: true, color: primaryColor
      });
      
      slide.addText('Sample content text showing how text appears in this template design.', {
        x: cardX + 0.15, y: contentTop + 0.7, w: cardWidth - 0.3, h: 2.0,
        fontSize: 10, color: secondaryColor, valign: 'top', wrap: true
      });
    }

    // Footer (if template uses it)
    if (template.content?.showFooter) {
      const footerY = template.content.footerY || 5.15;
      
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 0.5, y: footerY - 0.05, w: 9.0, h: 0.02,
        fill: { color: '334155' }
      });
      
      slide.addText(`${template.name} Template | Slide ${idx + 1}`, {
        x: 0.5, y: footerY, w: 9.0, h: 0.25,
        fontSize: 9, color: '94A3B8'
      });
      console.log(`    Footer: YES at ${footerY}in`);
    } else {
      console.log(`    Footer: NO`);
    }

    // Special features note
    const features = [];
    if (template.colors?.useGradients) features.push('Gradients');
    if (template.colors?.luxuryMode) features.push('Luxury');
    if (template.colors?.monochromeMode) features.push('Monochrome');
    if (template.colors?.highContrast) features.push('High Contrast');
    if (template.content?.asymmetricLayout) features.push('Asymmetric');
    if (template.content?.magazineStyle) features.push('Magazine');
    if (template.content?.gridBased) features.push('Grid');
    if (template.content?.brutalist) features.push('Brutalist');
    
    if (features.length > 0) {
      console.log(`    Special: ${features.join(', ')}`);
    }
    
    console.log();
  });

  // Save to file
  const outputPath = path.join(__dirname, 'output', 'TEMPLATE_PREVIEW.pptx');
  const buffer = await pptx.write({ outputType: 'nodebuffer' });
  
  const fs = await import('fs');
  fs.default.writeFileSync(outputPath, buffer);

  console.log('═══════════════════════════════════════════════════════════');
  console.log('✅ Template preview generated!');
  console.log(`📁 File: ${outputPath}`);
  console.log('═══════════════════════════════════════════════════════════\n');
  console.log('Open the file to see all 10 templates side-by-side!');
}

generateTemplatePreviews().catch(err => {
  console.error('❌ Error:', err);
  process.exit(1);
});
