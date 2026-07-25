import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  PageBreak,
  AlignmentType,
  WidthType,
  ShadingType,
  HeadingLevel
} from 'docx';
import { logger } from '../../../utils/logger.js';

/**
 * Simple markdown parser for fill mode
 * Supports: **bold**, *italic*, `code`
 */
function parseMarkdownToTextRuns(text, baseOptions = {}) {
  if (!text) return [new TextRun({ text: '', ...baseOptions })];

  const {
    size = 22,
    font = 'Calibri',
    color = '000000',
    bold = false,
    italics = false
  } = baseOptions;

  const runs = [];
  const regex = /(\*\*(?:[^*]|\*[^*])+\*\*)|(\*(?:[^*])+\*)|(`[^`]+`)/g;

  let lastIndex = 0;
  let match;

  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      const plainText = text.substring(lastIndex, match.index);
      runs.push(new TextRun({ text: plainText, font, size, color, bold, italics }));
    }

    const token = match[0];
    
    // **bold**
    if (token.startsWith('**') && token.endsWith('**')) {
      const content = token.slice(2, -2);
      runs.push(new TextRun({ text: content, font, size, color, bold: true, italics }));
    }
    // *italic*
    else if (token.startsWith('*') && token.endsWith('*')) {
      const content = token.slice(1, -1);
      runs.push(new TextRun({ text: content, font, size, color, bold, italics: true }));
    }
    // `code`
    else if (token.startsWith('`') && token.endsWith('`')) {
      const content = token.slice(1, -1);
      runs.push(new TextRun({ text: content, font: 'Consolas', size, color: '0F172A', bold }));
    }

    lastIndex = regex.lastIndex;
  }

  if (lastIndex < text.length) {
    runs.push(new TextRun({ text: text.substring(lastIndex), font, size, color, bold, italics }));
  }

  return runs.length > 0 ? runs : [new TextRun({ text: '', ...baseOptions })];
}

/**
 * FILL MODE DOCX BUILDER
 * Creates clean, simple documents matching template style
 * No fancy cover pages, headers, or custom themes
 * Just content, tables, and basic formatting
 */
export const buildFillModeDocx = async (data) => {
  logger.info(`Building Fill Mode DOCX -> Title: "${data.title}"`);
  
  // Extract template styling
  const templateStyle = data.templateStyle || {};
  const primaryColor = templateStyle.colors?.[0] || '000000';
  const headerColor = templateStyle.colors?.[1] || '4472C4';
  const shadingColor = templateStyle.shadingColors?.[0] || 'F2F2F2';
  const baseFontSize = templateStyle.fontSizes?.[0] || 11;
  const headingFontSize = Math.max(...(templateStyle.fontSizes || [11]));
  
  logger.info(`Using template style: Primary color=${primaryColor}, Shading=${shadingColor}, Font=${baseFontSize}pt`);
  
  const children = [];
  
  // Simple title page
  children.push(
    new Paragraph({
      text: data.title || 'Document',
      heading: HeadingLevel.TITLE,
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      run: {
        size: headingFontSize * 2, // Convert to half-points
        color: primaryColor
      }
    })
  );
  
  if (data.subtitle) {
    children.push(
      new Paragraph({
        text: data.subtitle,
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 }
      })
    );
  }
  
  // Add content sections
  if (data.pages && Array.isArray(data.pages)) {
    data.pages.forEach((page, idx) => {
      // Section heading
      if (page.heading) {
        const headingLevel = page.headingLevel === 1 ? HeadingLevel.HEADING_1 : 
                             page.headingLevel === 3 ? HeadingLevel.HEADING_3 : 
                             HeadingLevel.HEADING_2;
        
        children.push(
          new Paragraph({
            text: page.heading,
            heading: headingLevel,
            spacing: { before: 300, after: 200 }
          })
        );
      }
      
      // Paragraphs
      if (page.paragraphs && Array.isArray(page.paragraphs)) {
        page.paragraphs.forEach(p => {
          const paraText = typeof p === 'string' ? p : p.text || '';
          const alignment = (typeof p === 'object' && p.alignment) ? 
            (p.alignment === 'center' ? AlignmentType.CENTER : 
             p.alignment === 'right' ? AlignmentType.RIGHT : 
             AlignmentType.LEFT) : AlignmentType.LEFT;
          
          children.push(
            new Paragraph({
              children: parseMarkdownToTextRuns(paraText),
              alignment: alignment,
              spacing: { after: 200 }
            })
          );
        });
      }
      
      // Lists
      if (page.bulletList && Array.isArray(page.bulletList)) {
        page.bulletList.forEach(item => {
          const itemText = typeof item === 'string' ? item : item.text || '';
          const level = (typeof item === 'object' && item.level) || 0;
          
          children.push(
            new Paragraph({
              children: parseMarkdownToTextRuns(itemText),
              bullet: { level: level },
              spacing: { after: 100 }
            })
          );
        });
      }
      
      // Tables
      if (page.table && page.table.headers && page.table.rows) {
        // Table title
        if (page.table.title) {
          children.push(
            new Paragraph({
              text: page.table.title,
              bold: true,
              spacing: { before: 200, after: 100 }
            })
          );
        }
        
        const tableRows = [];
        
        // Header row
        const headerCells = page.table.headers.map(h => 
          new TableCell({
            children: [
              new Paragraph({ 
                children: [new TextRun({ text: h, bold: true, color: 'FFFFFF', size: baseFontSize * 2 })] 
              })
            ],
            shading: { fill: headerColor }
          })
        );
        tableRows.push(new TableRow({ children: headerCells }));
        
        // Data rows
        page.table.rows.forEach((row, rowIdx) => {
          const cells = row.map(cell => 
            new TableCell({
              children: [new Paragraph({ text: String(cell), run: { size: baseFontSize * 2 } })],
              shading: { fill: rowIdx % 2 === 0 ? 'FFFFFF' : shadingColor }
            })
          );
          tableRows.push(new TableRow({ children: cells }));
        });
        
        children.push(
          new Table({
            rows: tableRows,
            width: { size: 100, type: WidthType.PERCENTAGE }
          })
        );
        children.push(new Paragraph({ text: '', spacing: { after: 200 } }));
      }
      
      // Diagrams (simple text representation)
      if (page.visualNeed && page.visualNeed.type === 'diagram' && page.visualNeed.diagram) {
        const diagram = page.visualNeed.diagram;
        
        if (diagram.title) {
          children.push(
            new Paragraph({
              text: diagram.title,
              bold: true,
              alignment: AlignmentType.CENTER,
              spacing: { before: 200, after: 100 }
            })
          );
        }
        
        if (diagram.steps && Array.isArray(diagram.steps)) {
          diagram.steps.forEach((step, stepIdx) => {
            children.push(
              new Paragraph({
                text: `${stepIdx + 1}. ${step}`,
                spacing: { after: 100, left: 400 }
              })
            );
          });
        }
        
        children.push(new Paragraph({ text: '', spacing: { after: 200 } }));
      }
      
      // Page break between major sections (not after every page)
      if (idx < data.pages.length - 1 && page.headingLevel === 1) {
        children.push(new Paragraph({ children: [new PageBreak()] }));
      }
    });
  }
  
  // Create document - simple, no headers/footers
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: 1440,
            right: 1440,
            bottom: 1440,
            left: 1440
          }
        }
      },
      children: children
    }]
  });
  
  const buffer = await Packer.toBuffer(doc);
  logger.info(`Fill Mode DOCX Complete -> Buffer size: ${buffer.length} bytes`);
  return buffer;
};

export default {
  buildFillModeDocx
};
