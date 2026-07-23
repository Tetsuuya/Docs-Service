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
 * Builds a styled Microsoft Word (.docx) file using Gemini's dynamic JSON color theme and layouts.
 * Features:
 * - Executive Cover Page (Title Banner, Embedded Badge Image, Metadata, PageBreak)
 * - Dynamic Theme Colors per document (Primary, Secondary, Accent, Light Background)
 * - Running Header & Footer (Page X of Y)
 * - Bulleted Lists & Styled Typography
 * - Callout Highlight Boxes (Dynamic border accent + light background)
 * - Styled Data Tables (Dynamic Header fill, bold white text, alternating row shading)
 */
export const buildDocxFile = async (data) => {
  logger.info(`Building DOCX Document -> Title: "${data.title || 'Untitled Document'}"`);

  // Default values fallback
  const titleText = data.title || 'Document Report';
  const subtitleText = data.subtitle || 'Generated via AI Docs Service';
  const headerLabel = data.headerText || 'DOCS SERVICE | CONFIDENTIAL';
  const sections = data.sections || [];
  const tableData = data.table || null;
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

  // 1. Build Document Children Array
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
                      size: 18, // 9pt
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
                      size: 40, // 20pt
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
                      size: 24, // 12pt
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
                    children: [
                      new TextRun({ text: data.executiveOverview, size: 20, color: COLOR_BODY, font: 'Calibri' })
                    ]
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

  // Metadata Card Block (Date, Author, Status)
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
                    new TextRun({ text: 'Docs Service AI Publishing Engine', size: 18, color: COLOR_BODY, font: 'Calibri' })
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

  // --- Render Page-Structured Content (Calibrated Page-by-Page Layout) ---
  const pagesList = (data.pages && Array.isArray(data.pages) && data.pages.length > 0) ? data.pages : (data.sections || []);

  pagesList.forEach((sec, idx) => {
    if (sec.heading) {
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 200, after: 140 },
          children: [
            new TextRun({
              text: sec.heading,
              bold: true,
              size: 32, // 16pt
              color: COLOR_NAVY,
              font: 'Calibri'
            })
          ]
        })
      );
    }

    if (Array.isArray(sec.paragraphs)) {
      sec.paragraphs.forEach((pText) => {
        children.push(
          new Paragraph({
            spacing: { before: 0, after: 140, line: 260 },
            children: [
              new TextRun({
                text: pText,
                size: 22, // 11pt
                color: COLOR_BODY,
                font: 'Calibri'
              })
            ]
          })
        );
      });
    }

    if (Array.isArray(sec.bulletList) && sec.bulletList.length > 0) {
      sec.bulletList.forEach((bText) => {
        children.push(
          new Paragraph({
            bullet: { level: 0 },
            spacing: { before: 0, after: 100, line: 240 },
            children: [
              new TextRun({
                text: bText,
                size: 22, // 11pt
                color: COLOR_BODY,
                font: 'Calibri'
              })
            ]
          })
        );
      });
      children.push(new Paragraph({ spacing: { after: 100 } }));
    }

    // Callout / Highlight Box with Dynamic Accent Color & Background
    if (sec.calloutBox) {
      children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  shading: { fill: COLOR_LIGHT_BG, type: ShadingType.CLEAR },
                  borders: {
                    top: { style: BorderStyle.NONE },
                    bottom: { style: BorderStyle.NONE },
                    right: { style: BorderStyle.NONE },
                    left: { style: BorderStyle.SINGLE, size: 24, color: COLOR_NAVY }
                  },
                  margins: { top: 120, bottom: 120, left: 180, right: 180 },
                  children: [
                    new Paragraph({
                      spacing: { before: 0, after: 0, line: 240 },
                      children: [
                        new TextRun({
                          text: `💡 NOTE: ${sec.calloutBox}`,
                          bold: true,
                          size: 20, // 10pt
                          color: COLOR_NAVY,
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
    }

    // Insert Page Break after each page item (so next page starts on a new page cleanly)
    if (idx < pagesList.length - 1) {
      children.push(
        new Paragraph({
          children: [new PageBreak()]
        })
      );
    }
  });

  // --- Render Styled Table ---
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
              size: 28, // 14pt
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
                text: hText,
                bold: true,
                size: 20, // 10pt
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
              children: [
                new TextRun({
                  text: String(cellText),
                  size: 20, // 10pt
                  color: COLOR_BODY,
                  font: 'Calibri'
                })
              ]
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

  // 2. Create Document Instance with Header & Footer
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1440, // 1 inch
              bottom: 1440,
              left: 1440,
              right: 1440
            }
          }
        },
        // Running Header
        headers: {
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
                    size: 18, // 9pt
                    color: COLOR_MUTED,
                    bold: true,
                    font: 'Calibri'
                  })
                ]
              })
            ]
          })
        },
        // Running Footer (Page Numbers)
        footers: {
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

  // Pack into binary buffer
  const buffer = await Packer.toBuffer(doc);
  logger.info(`DOCX File Generation Completed -> Buffer size: ${buffer.length} bytes`);
  return buffer;
};
