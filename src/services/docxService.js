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
  AlignmentType,
  BorderStyle,
  WidthType,
  ShadingType,
  HeadingLevel
} from 'docx';
import { logger } from '../utils/logger.js';

/**
 * Builds a styled, professional Microsoft Word (.docx) file from structured document JSON.
 * Features:
 * - Running Header & Footer (Page X of Y)
 * - Styled Typography (Title 28pt bold, Heading 1 #1E3A8A, Heading 2 #2563EB)
 * - Callout Highlight Boxes (Left border accent + light background)
 * - Styled Data Tables (Navy Header #1E3A8A, bold white text, alternating row shading #F8FAFC, light borders #E2E8F0)
 */
export const buildDocxFile = async (data) => {
  logger.info(`Building DOCX Document -> Title: "${data.title || 'Untitled Document'}"`);

  // Default values fallback
  const titleText = data.title || 'Document Report';
  const subtitleText = data.subtitle || 'Generated via AI Docs Service';
  const headerLabel = data.headerText || 'DOCS SERVICE | CONFIDENTIAL';
  const sections = data.sections || [];
  const tableData = data.table || null;

  // Primary Theme Colors (Hex strings without # for docx)
  const COLOR_NAVY = '1E3A8A';
  const COLOR_BLUE = '2563EB';
  const COLOR_DARK = '0F172A';
  const COLOR_BODY = '334155';
  const COLOR_MUTED = '64748B';
  const COLOR_LIGHT_BG = 'F8FAFC';
  const COLOR_BORDER = 'CBD5E1';
  const COLOR_WHITE = 'FFFFFF';

  // 1. Build Document Children Array
  const children = [];

  // --- Document Title ---
  children.push(
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 120 },
      children: [
        new TextRun({
          text: titleText,
          bold: true,
          size: 44, // 22pt
          color: COLOR_DARK,
          font: 'Calibri'
        })
      ]
    })
  );

  // --- Document Subtitle ---
  children.push(
    new Paragraph({
      alignment: AlignmentType.LEFT,
      spacing: { before: 0, after: 360 },
      children: [
        new TextRun({
          text: subtitleText,
          italics: true,
          size: 24, // 12pt
          color: COLOR_MUTED,
          font: 'Calibri'
        })
      ]
    })
  );

  // Horizontal Divider Line
  children.push(
    new Paragraph({
      border: {
        bottom: { color: COLOR_BLUE, space: 1, style: BorderStyle.SINGLE, size: 12 }
      },
      spacing: { after: 360 }
    })
  );

  // --- Render Sections ---
  sections.forEach((sec, idx) => {
    if (sec.heading) {
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 280, after: 140 },
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
            spacing: { before: 0, after: 160, line: 276 },
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

    // Callout / Highlight Box
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
                  margins: { top: 140, bottom: 140, left: 200, right: 200 },
                  children: [
                    new Paragraph({
                      spacing: { before: 0, after: 0, line: 260 },
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

      // Spacing after callout box
      children.push(new Paragraph({ spacing: { after: 240 } }));
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

    // Header Row
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

    // Data Rows
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
