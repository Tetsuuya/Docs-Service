import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  Header,
  Footer,
  PageNumber,
  PageBreak,
  AlignmentType,
  BorderStyle,
  WidthType,
  ShadingType,
  HeadingLevel,
  ImageRun
} from 'docx';
import { logger } from '../utils/logger.js';

/**
 * Parses inline markdown formatting (**bold**, *italics*, `code`) into an array of styled TextRun objects.
 */
function parseMarkdownToTextRuns(text, baseOptions = {}) {
  if (!text) return [new TextRun({ text: '', ...baseOptions })];

  const {
    size = 22,
    font = 'Calibri',
    color = '334155',
    bold = false,
    italics = false
  } = baseOptions;

  const runs = [];
  // Regex matches **bold**, *italics*, `code`
  const regex = /(\*\*(?:[^*]|\*[^*])+\*\*|\*(?:[^*])+\*|`[^`]+`)/g;

  let lastIndex = 0;
  let match;

  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      const plainText = text.substring(lastIndex, match.index);
      runs.push(new TextRun({ text: plainText, font, size, color, bold, italics }));
    }

    const token = match[0];
    if (token.startsWith('**') && token.endsWith('**')) {
      const content = token.slice(2, -2);
      runs.push(new TextRun({ text: content, bold: true, font, size, color }));
    } else if (token.startsWith('*') && token.endsWith('*')) {
      const content = token.slice(1, -1);
      runs.push(new TextRun({ text: content, italics: true, font, size, color }));
    } else if (token.startsWith('`') && token.endsWith('`')) {
      const content = token.slice(1, -1);
      runs.push(
        new TextRun({
          text: content,
          font: 'Consolas',
          size: Math.max(16, size - 2),
          color: '0F172A',
          shading: { fill: 'F1F5F9', type: ShadingType.CLEAR }
        })
      );
    }

    lastIndex = regex.lastIndex;
  }

  if (lastIndex < text.length) {
    const plainText = text.substring(lastIndex);
    runs.push(new TextRun({ text: plainText, font, size, color, bold, italics }));
  }

  return runs.length > 0 ? runs : [new TextRun({ text, font, size, color, bold, italics })];
}

/**
 * Builds a styled Microsoft Word (.docx) file using Gemini's dynamic JSON color theme and AST layouts.
 * Features:
 * - Executive Cover Page (Title Banner, Embedded Badge Image, Metadata, PageBreak)
 * - Header/Footer titlePage suppression (No header/footer on page 1)
 * - Executive Image Placeholder Cards (Structured visual descriptions & captions)
 * - Human-Level AI Visual Detector Rendering (Diagrams, Image Cards, Specification Tables)
 * - Inline Markdown Parser (**bold**, *italics*, `code`)
 * - Multi-type Callout Cards (Info, Warning, Success, Tip, Note)
 * - Stat / KPI Grid Cards (Big numbers + labels)
 * - Code Blocks & Blockquotes
 * - Styled Data Tables (Header fill, bold white text, alternating row shading)
 */
export const buildDocxFile = async (data) => {
  logger.info(`Building Claude-Level DOCX Document -> Title: "${data.title || 'Untitled Document'}"`);

  // Default values fallback
  const titleText = data.title || 'Document Report';
  const subtitleText = data.subtitle || 'Generated via AI Docs Service';
  const headerLabel = data.headerText || 'DOCS SERVICE | CONFIDENTIAL';
  const imageBuffer = data.imageBuffer || null;

  // Dynamic Theme Colors from Gemini (Hex strings without #)
  const theme = data.theme || {};
  const COLOR_NAVY = theme.primaryColor || '1E3A8A';
  const COLOR_BLUE = theme.secondaryColor || '2563EB';
  const COLOR_ACCENT = theme.accentColor || '0284C7';
  const COLOR_DARK = '0F172A';
  const COLOR_BODY = theme.textColor || '334155';
  const COLOR_MUTED = '64748B';
  const COLOR_LIGHT_BG = theme.lightBgColor || 'F8FAFC';
  const COLOR_BORDER = 'CBD5E1';
  const COLOR_WHITE = 'FFFFFF';

  logger.info(`Applying Dynamic Theme -> Primary: #${COLOR_NAVY}, Secondary: #${COLOR_BLUE}, LightBg: #${COLOR_LIGHT_BG}`);

  const children = [];

  // --- EXECUTIVE COVER PAGE ---
  const docTypeTag = (data.docTypeTag || 'DOCUMENT OVERVIEW').toUpperCase();

  // Title Banner Box with Primary Color Shading Fill
  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: COLOR_NAVY, type: ShadingType.CLEAR },
              margins: { top: 360, bottom: 360, left: 300, right: 300 },
              borders: {
                top: { style: BorderStyle.NONE },
                bottom: { style: BorderStyle.NONE },
                left: { style: BorderStyle.NONE },
                right: { style: BorderStyle.NONE }
              },
              children: [
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  spacing: { before: 0, after: 120 },
                  children: [
                    new TextRun({
                      text: `🏷️ ${docTypeTag}`,
                      bold: true,
                      size: 18,
                      color: COLOR_ACCENT,
                      font: 'Calibri'
                    })
                  ]
                }),
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  spacing: { before: 0, after: 120 },
                  children: [
                    new TextRun({
                      text: titleText,
                      bold: true,
                      size: 40,
                      color: COLOR_WHITE,
                      font: 'Calibri'
                    })
                  ]
                }),
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  spacing: { before: 0, after: 0 },
                  children: [
                    new TextRun({
                      text: subtitleText,
                      italics: true,
                      size: 24,
                      color: COLOR_LIGHT_BG,
                      font: 'Calibri'
                    })
                  ]
                })
              ]
            })
          ]
        })
      ]
    })
  );

  children.push(new Paragraph({ spacing: { after: 360 } }));

  // Embedded Cover / Badge Image (if uploaded)
  if (imageBuffer) {
    try {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 120, after: 360 },
          children: [
            new ImageRun({
              data: imageBuffer,
              transformation: {
                width: 220,
                height: 220
              }
            })
          ]
        })
      );
      logger.info('Successfully embedded uploaded image onto Cover Page!');
    } catch (err) {
      logger.warn(`Failed embedding image onto Cover Page: ${err.message}`);
    }
  }

  // Executive Overview Card Block
  if (data.executiveOverview) {
    children.push(
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                shading: { fill: COLOR_LIGHT_BG, type: ShadingType.CLEAR },
                margins: { top: 160, bottom: 160, left: 200, right: 200 },
                borders: {
                  top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
                  bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
                  left: { style: BorderStyle.SINGLE, size: 24, color: COLOR_NAVY },
                  right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER }
                },
                children: [
                  new Paragraph({
                    spacing: { before: 0, after: 60 },
                    children: [
                      new TextRun({ text: '📋 EXECUTIVE OVERVIEW', bold: true, size: 18, color: COLOR_NAVY, font: 'Calibri' })
                    ]
                  }),
                  new Paragraph({
                    spacing: { before: 0, after: 0, line: 240 },
                    children: parseMarkdownToTextRuns(data.executiveOverview, { size: 20, color: COLOR_BODY, font: 'Calibri' })
                  })
                ]
              })
            ]
          })
        ]
      })
    );
    children.push(new Paragraph({ spacing: { after: 200 } }));
  }

  // Table of Contents Box
  const tocHeadingsList = data.sectionHeadings || (data.pages ? data.pages.map(p => p.heading).filter(Boolean) : []);
  if (tocHeadingsList.length > 0) {
    const tocRows = [
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: COLOR_NAVY, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 160, right: 160 },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_NAVY },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_NAVY },
              left: { style: BorderStyle.SINGLE, size: 4, color: COLOR_NAVY },
              right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_NAVY }
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: '📚 TABLE OF CONTENTS / DOCUMENT OUTLINE', bold: true, size: 18, color: COLOR_WHITE, font: 'Calibri' })
                ]
              })
            ]
          })
        ]
      })
    ];

    tocHeadingsList.forEach((hText, hIdx) => {
      tocRows.push(
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: hIdx % 2 === 0 ? COLOR_LIGHT_BG : COLOR_WHITE, type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 160, right: 160 },
              borders: {
                top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
                bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
                left: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
                right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER }
              },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ text: `Page ${hIdx + 2}: `, bold: true, size: 18, color: COLOR_NAVY, font: 'Calibri' }),
                    new TextRun({ text: hText, size: 18, color: COLOR_BODY, font: 'Calibri' })
                  ]
                })
              ]
            })
          ]
        })
      );
    });

    children.push(
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: tocRows
      })
    );
    children.push(new Paragraph({ spacing: { after: 200 } }));
  }

  // Metadata Card Block
  children.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: COLOR_LIGHT_BG, type: ShadingType.CLEAR },
              margins: { top: 160, bottom: 160, left: 200, right: 200 },
              borders: {
                top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
                bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
                left: { style: BorderStyle.SINGLE, size: 24, color: COLOR_BLUE },
                right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER }
              },
              children: [
                new Paragraph({
                  spacing: { before: 0, after: 40 },
                  children: [
                    new TextRun({ text: 'PUBLICATION METADATA', bold: true, size: 18, color: COLOR_NAVY, font: 'Calibri' })
                  ]
                }),
                new Paragraph({
                  spacing: { before: 0, after: 20 },
                  children: [
                    new TextRun({ text: 'Author / Engine: ', bold: true, size: 18, color: COLOR_MUTED, font: 'Calibri' }),
                    new TextRun({ text: 'Docs Service Enterprise Publishing Engine', size: 18, color: COLOR_BODY, font: 'Calibri' })
                  ]
                }),
                new Paragraph({
                  spacing: { before: 0, after: 0 },
                  children: [
                    new TextRun({ text: 'Generated Date: ', bold: true, size: 18, color: COLOR_MUTED, font: 'Calibri' }),
                    new TextRun({ text: new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }), size: 18, color: COLOR_BODY, font: 'Calibri' })
                  ]
                })
              ]
            })
          ]
        })
      ]
    })
  );

  // Cover Page Break -> Section 1 starts cleanly on Page 2
  children.push(
    new Paragraph({
      children: [new PageBreak()]
    })
  );

  // --- RENDER PAGE-STRUCTURED CONTENT (AST BLOCKS) ---
  const pagesList = (data.pages && Array.isArray(data.pages) && data.pages.length > 0) ? data.pages : (data.sections || []);

  pagesList.forEach((sec, idx) => {
    // 1. Heading
    if (sec.heading) {
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 200, after: 140 },
          children: [
            new TextRun({
              text: sec.heading,
              bold: true,
              size: 32,
              color: COLOR_NAVY,
              font: 'Calibri'
            })
          ]
        })
      );
    }

    // 2. Paragraphs (with markdown parsing)
    if (Array.isArray(sec.paragraphs)) {
      sec.paragraphs.forEach((pText) => {
        children.push(
          new Paragraph({
            spacing: { before: 0, after: 140, line: 260 },
            children: parseMarkdownToTextRuns(pText, { size: 22, color: COLOR_BODY, font: 'Calibri' })
          })
        );
      });
    }

    // 3. HUMAN-LEVEL AI VISUAL DETECTOR RENDERER
    const visual = sec.visualNeed || {};

    // A. Diagram Visual (Process Flow Workflow Steps)
    if (visual.type === 'diagram' && visual.diagram && Array.isArray(visual.diagram.steps)) {
      const steps = visual.diagram.steps;
      const diagTitle = visual.diagram.title || 'TECHNICAL PROCESS WORKFLOW';

      const stepCells = steps.map((stepTitle, sIdx) => {
        return new TableCell({
          shading: { fill: COLOR_LIGHT_BG, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 120, right: 120 },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BLUE },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
            left: { style: BorderStyle.SINGLE, size: 12, color: COLOR_BLUE },
            right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER }
          },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 20 },
              children: [
                new TextRun({ text: `STEP ${sIdx + 1}`, bold: true, size: 14, color: COLOR_BLUE, font: 'Calibri' })
              ]
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 0 },
              children: [
                new TextRun({ text: stepTitle, bold: true, size: 18, color: COLOR_NAVY, font: 'Calibri' })
              ]
            }),
            ...(sIdx < steps.length - 1 ? [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 20, after: 0 },
                children: [
                  new TextRun({ text: '➔', bold: true, size: 18, color: COLOR_ACCENT, font: 'Calibri' })
                ]
              })
            ] : [])
          ]
        });
      });

      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  shading: { fill: COLOR_NAVY, type: ShadingType.CLEAR },
                  margins: { top: 80, bottom: 80, left: 160, right: 160 },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({ text: `📌 DIAGRAM: ${diagTitle.toUpperCase()}`, bold: true, size: 16, color: COLOR_WHITE, font: 'Calibri' })
                      ]
                    })
                  ]
                })
              ]
            }),
            new TableRow({ children: stepCells })
          ]
        })
      );
      children.push(new Paragraph({ spacing: { after: 160 } }));
    }

    // B. Executive Image Placeholder Card (Structured Placeholder with Description & Caption)
    if (visual.type === 'image' && visual.image) {
      const imgInfo = visual.image;
      const captionText = imgInfo.caption || `Figure ${idx + 1}.1: Recommended Image Visual`;
      const imageDesc = imgInfo.imagePrompt || imgInfo.description || 'Illustrative topic photo or diagram';
      const actualImageBuf = visual.imageBuffer || sec.imageBuffer || null;

      const cellChildren = [
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { before: 0, after: 60 },
          children: [
            new TextRun({ text: '🖼️ RECOMMENDED IMAGE PLACEHOLDER', bold: true, size: 18, color: COLOR_NAVY, font: 'Calibri' })
          ]
        })
      ];

      if (actualImageBuf) {
        try {
          cellChildren.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 60, after: 60 },
              children: [
                new ImageRun({
                  data: actualImageBuf,
                  transformation: { width: 420, height: 240 }
                })
              ]
            })
          );
        } catch (e) {
          // ignore
        }
      } else {
        cellChildren.push(
          new Paragraph({
            spacing: { before: 0, after: 60, line: 240 },
            children: [
              new TextRun({ text: 'Suggested Visual: ', bold: true, size: 18, color: COLOR_MUTED, font: 'Calibri' }),
              new TextRun({ text: imageDesc, italics: true, size: 18, color: COLOR_BODY, font: 'Calibri' })
            ]
          })
        );
      }

      cellChildren.push(
        new Paragraph({
          spacing: { before: 0, after: 0 },
          children: [
            new TextRun({ text: captionText, bold: true, size: 18, color: COLOR_NAVY, font: 'Calibri' })
          ]
        })
      );

      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  shading: { fill: COLOR_LIGHT_BG, type: ShadingType.CLEAR },
                  margins: { top: 140, bottom: 140, left: 180, right: 180 },
                  borders: {
                    top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
                    bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
                    left: { style: BorderStyle.SINGLE, size: 24, color: COLOR_ACCENT },
                    right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER }
                  },
                  children: cellChildren
                })
              ]
            })
          ]
        })
      );
      children.push(new Paragraph({ spacing: { after: 160 } }));
    }

    // C. Section Data Table Visual
    if (visual.type === 'table' && visual.table && Array.isArray(visual.table.headers) && Array.isArray(visual.table.rows)) {
      const secTable = visual.table;
      const secTableRows = [];

      const secHeaderCells = secTable.headers.map((hText) => {
        return new TableCell({
          shading: { fill: COLOR_NAVY, type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 140, right: 140 },
          children: [
            new Paragraph({
              alignment: AlignmentType.LEFT,
              children: [
                new TextRun({ text: String(hText), bold: true, size: 18, color: COLOR_WHITE, font: 'Calibri' })
              ]
            })
          ]
        });
      });

      secTableRows.push(new TableRow({ children: secHeaderCells, tableHeader: true, cantSplit: true }));

      secTable.rows.forEach((rArray, rIdx) => {
        const rowFill = rIdx % 2 !== 0 ? COLOR_LIGHT_BG : COLOR_WHITE;
        const rCells = rArray.map((cText) => {
          return new TableCell({
            shading: { fill: rowFill, type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 140, right: 140 },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
              left: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
              right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER }
            },
            children: [
              new Paragraph({
                children: parseMarkdownToTextRuns(String(cText), { size: 18, color: COLOR_BODY, font: 'Calibri' })
              })
            ]
          });
        });
        secTableRows.push(new TableRow({ children: rCells, cantSplit: true }));
      });

      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: secTableRows
        })
      );
      children.push(new Paragraph({ spacing: { after: 160 } }));
    }

    // 4. Stat Cards / KPI Grid
    if (Array.isArray(sec.statCards) && sec.statCards.length > 0) {
      const cardCells = sec.statCards.map((card) => {
        return new TableCell({
          shading: { fill: COLOR_LIGHT_BG, type: ShadingType.CLEAR },
          margins: { top: 140, bottom: 140, left: 160, right: 160 },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
            left: { style: BorderStyle.SINGLE, size: 16, color: COLOR_BLUE },
            right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER }
          },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 40 },
              children: [
                new TextRun({ text: String(card.value || '0'), bold: true, size: 36, color: COLOR_NAVY, font: 'Calibri' })
              ]
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 20 },
              children: [
                new TextRun({ text: String(card.label || '').toUpperCase(), bold: true, size: 16, color: COLOR_MUTED, font: 'Calibri' })
              ]
            }),
            ...(card.subtext ? [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({ text: String(card.subtext), italics: true, size: 16, color: COLOR_BODY, font: 'Calibri' })
                ]
              })
            ] : [])
          ]
        });
      });

      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [new TableRow({ children: cardCells })]
        })
      );
      children.push(new Paragraph({ spacing: { after: 160 } }));
    }

    // 5. Bullet List (with markdown parsing)
    if (Array.isArray(sec.bulletList) && sec.bulletList.length > 0) {
      sec.bulletList.forEach((bText) => {
        children.push(
          new Paragraph({
            bullet: { level: 0 },
            spacing: { before: 0, after: 100, line: 240 },
            children: parseMarkdownToTextRuns(bText, { size: 22, color: COLOR_BODY, font: 'Calibri' })
          })
        );
      });
      children.push(new Paragraph({ spacing: { after: 100 } }));
    }

    // 6. Code Block Snippet
    if (sec.codeBlock) {
      const codeObj = typeof sec.codeBlock === 'string' ? { code: sec.codeBlock, language: 'text' } : sec.codeBlock;
      const codeLines = (codeObj.code || '').split('\n');

      const codeParagraphs = [
        new Paragraph({
          spacing: { before: 0, after: 60 },
          children: [
            new TextRun({ text: `💻 ${ (codeObj.language || 'CODE').toUpperCase() } SNIPPET`, bold: true, size: 16, color: COLOR_ACCENT, font: 'Consolas' })
          ]
        }),
        ...codeLines.map(line => 
          new Paragraph({
            spacing: { before: 0, after: 20, line: 220 },
            children: [
              new TextRun({ text: line, size: 18, color: '0F172A', font: 'Consolas' })
            ]
          })
        )
      ];

      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  shading: { fill: 'F1F5F9', type: ShadingType.CLEAR },
                  margins: { top: 120, bottom: 120, left: 160, right: 160 },
                  borders: {
                    top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
                    bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
                    left: { style: BorderStyle.SINGLE, size: 16, color: COLOR_MUTED },
                    right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER }
                  },
                  children: codeParagraphs
                })
              ]
            })
          ]
        })
      );
      children.push(new Paragraph({ spacing: { after: 160 } }));
    }

    // 7. Blockquote Card
    if (sec.blockquote) {
      const quoteObj = typeof sec.blockquote === 'string' ? { quote: sec.blockquote } : sec.blockquote;

      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  shading: { fill: COLOR_LIGHT_BG, type: ShadingType.CLEAR },
                  margins: { top: 140, bottom: 140, left: 180, right: 180 },
                  borders: {
                    top: { style: BorderStyle.NONE },
                    bottom: { style: BorderStyle.NONE },
                    right: { style: BorderStyle.NONE },
                    left: { style: BorderStyle.SINGLE, size: 24, color: COLOR_BLUE }
                  },
                  children: [
                    new Paragraph({
                      spacing: { before: 0, after: quoteObj.author ? 40 : 0, line: 240 },
                      children: [
                        new TextRun({ text: `“${quoteObj.quote}”`, italics: true, size: 22, color: COLOR_DARK, font: 'Calibri' })
                      ]
                    }),
                    ...(quoteObj.author ? [
                      new Paragraph({
                        spacing: { before: 0, after: 0 },
                        children: [
                          new TextRun({ text: `— ${quoteObj.author}`, bold: true, size: 18, color: COLOR_MUTED, font: 'Calibri' })
                        ]
                      })
                    ] : [])
                  ]
                })
              ]
            })
          ]
        })
      );
      children.push(new Paragraph({ spacing: { after: 160 } }));
    }

    // 8. Multi-Type Callout Box
    if (sec.calloutBox) {
      let type = 'info';
      let textStr = '';
      let titleStr = '';

      if (typeof sec.calloutBox === 'string') {
        textStr = sec.calloutBox;
      } else {
        type = sec.calloutBox.type || 'info';
        textStr = sec.calloutBox.text || '';
        titleStr = sec.calloutBox.title || '';
      }

      let calloutBorderColor = COLOR_NAVY;
      let calloutFill = COLOR_LIGHT_BG;
      let iconLabel = '💡 NOTE:';

      if (type === 'warning') {
        calloutBorderColor = 'D97706';
        calloutFill = 'FEF3C7';
        iconLabel = '⚠️ WARNING:';
      } else if (type === 'success') {
        calloutBorderColor = '059669';
        calloutFill = 'ECFDF5';
        iconLabel = '✅ STRATEGIC HIGHLIGHT:';
      } else if (type === 'tip') {
        calloutBorderColor = '7C3AED';
        calloutFill = 'F5F3FF';
        iconLabel = '🚀 TIP / KEY TAKEAWAY:';
      }

      const displayHeader = titleStr ? `${iconLabel} ${titleStr}` : iconLabel;

      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  shading: { fill: calloutFill, type: ShadingType.CLEAR },
                  borders: {
                    top: { style: BorderStyle.NONE },
                    bottom: { style: BorderStyle.NONE },
                    right: { style: BorderStyle.NONE },
                    left: { style: BorderStyle.SINGLE, size: 24, color: calloutBorderColor }
                  },
                  margins: { top: 120, bottom: 120, left: 180, right: 180 },
                  children: [
                    new Paragraph({
                      spacing: { before: 0, after: 40 },
                      children: [
                        new TextRun({ text: displayHeader, bold: true, size: 20, color: calloutBorderColor, font: 'Calibri' })
                      ]
                    }),
                    new Paragraph({
                      spacing: { before: 0, after: 0, line: 240 },
                      children: parseMarkdownToTextRuns(textStr, { size: 20, color: COLOR_BODY, font: 'Calibri' })
                    })
                  ]
                })
              ]
            })
          ]
        })
      );
      children.push(new Paragraph({ spacing: { after: 160 } }));
    }

    // Insert Page Break after each section
    if (idx < pagesList.length - 1) {
      children.push(
        new Paragraph({
          children: [new PageBreak()]
        })
      );
    }
  });

  // --- RENDER STYLED DATA TABLE ---
  const tableData = data.table || null;
  if (tableData && Array.isArray(tableData.headers) && Array.isArray(tableData.rows)) {
    if (tableData.title) {
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 320, after: 160 },
          children: [
            new TextRun({
              text: tableData.title,
              bold: true,
              size: 28,
              color: COLOR_BLUE,
              font: 'Calibri'
            })
          ]
        })
      );
    }

    const tableRows = [];

    // Header Row with Dynamic Primary Color Fill
    const headerCells = tableData.headers.map((hText) => {
      return new TableCell({
        shading: { fill: COLOR_NAVY, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 160, right: 160 },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_NAVY },
          bottom: { style: BorderStyle.SINGLE, size: 12, color: COLOR_NAVY },
          left: { style: BorderStyle.SINGLE, size: 4, color: COLOR_NAVY },
          right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_NAVY }
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.LEFT,
            children: [
              new TextRun({
                text: String(hText),
                bold: true,
                size: 20,
                color: COLOR_WHITE,
                font: 'Calibri'
              })
            ]
          })
        ]
      });
    });

    tableRows.push(new TableRow({ children: headerCells, tableHeader: true, cantSplit: true }));

    // Data Rows with Dynamic Alternating Row Shading
    tableData.rows.forEach((rowArray, rIdx) => {
      const isOdd = rIdx % 2 !== 0;
      const rowFill = isOdd ? COLOR_LIGHT_BG : COLOR_WHITE;

      const cells = rowArray.map((cellText) => {
        return new TableCell({
          shading: { fill: rowFill, type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 160, right: 160 },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
            left: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER },
            right: { style: BorderStyle.SINGLE, size: 4, color: COLOR_BORDER }
          },
          children: [
            new Paragraph({
              alignment: AlignmentType.LEFT,
              children: parseMarkdownToTextRuns(String(cellText), { size: 20, color: COLOR_BODY, font: 'Calibri' })
            })
          ]
        });
      });

      tableRows.push(new TableRow({ children: cells, cantSplit: true }));
    });

    children.push(
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: tableRows
      })
    );
  }

  // --- CREATE DOCUMENT INSTANCE WITH DIFFERENT FIRST PAGE HEADER/FOOTER ---
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1440,
              bottom: 1440,
              left: 1440,
              right: 1440
            }
          },
          titlePage: true // Suppresses running headers & footers on Page 1 (Cover Page)
        },
        headers: {
          first: new Header({ children: [] }), // Cover page header is empty
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                border: {
                  bottom: { color: COLOR_BORDER, space: 4, style: BorderStyle.SINGLE, size: 4 }
                },
                spacing: { after: 200 },
                children: [
                  new TextRun({
                    text: headerLabel,
                    size: 18,
                    color: COLOR_MUTED,
                    bold: true,
                    font: 'Calibri'
                  })
                ]
              })
            ]
          })
        },
        footers: {
          first: new Footer({ children: [] }), // Cover page footer is empty
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                border: {
                  top: { color: COLOR_BORDER, space: 4, style: BorderStyle.SINGLE, size: 4 }
                },
                spacing: { before: 200 },
                children: [
                  new TextRun({ text: 'Page ', size: 18, color: COLOR_MUTED, font: 'Calibri' }),
                  new TextRun({ children: [PageNumber.CURRENT], size: 18, color: COLOR_MUTED, font: 'Calibri' }),
                  new TextRun({ text: ' of ', size: 18, color: COLOR_MUTED, font: 'Calibri' }),
                  new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: COLOR_MUTED, font: 'Calibri' })
                ]
              })
            ]
          })
        },
        children: children
      }
    ]
  });

  const buffer = await Packer.toBuffer(doc);
  logger.info(`Claude-Level DOCX File Generation Completed -> Buffer size: ${buffer.length} bytes`);
  return buffer;
};
