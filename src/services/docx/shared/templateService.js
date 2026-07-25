import fs from 'fs';
import mammoth from 'mammoth';
import AdmZip from 'adm-zip';
import { logger } from '../../../utils/logger.js';

/**
 * DEEP TEMPLATE ANALYSIS
 * Extracts visual structure from DOCX: pages, images, tables, colors, formatting
 * This gives AI complete understanding of template's visual layout
 */
export const parseTemplate = async (templatePath) => {
  try {
    logger.info(`Parsing template file: ${templatePath}`);
    
    // Step 1: Extract plain text (for content reference)
    const result = await mammoth.extractRawText({ path: templatePath });
    const fullText = result.value;
    
    // Step 2: Deep analysis - open DOCX as ZIP and parse XML
    const zip = new AdmZip(templatePath);
    const zipEntries = zip.getEntries();
    
    // Extract document.xml (main content structure)
    const documentXmlEntry = zipEntries.find(entry => entry.entryName === 'word/document.xml');
    const documentXml = documentXmlEntry ? documentXmlEntry.getData().toString('utf8') : '';
    
    // Analyze page structure
    const pageBreaks = (documentXml.match(/<w:br w:type="page"\/>/g) || []).length;
    const estimatedPages = Math.max(1, pageBreaks + 1);
    
    // Analyze images
    const imageRels = zipEntries.filter(entry => entry.entryName.startsWith('word/media/')).length;
    
    // Analyze tables with structure
    const tables = (documentXml.match(/<w:tbl>/g) || []).length;
    const tableStructures = extractTableStructures(documentXml);
    
    // Analyze text formatting
    const boldCount = (documentXml.match(/<w:b\/>/g) || []).length;
    const italicCount = (documentXml.match(/<w:i\/>/g) || []).length;
    
    // Extract colors used
    const colorMatches = documentXml.match(/<w:color w:val="([A-F0-9]{6})"\/>/g) || [];
    const colors = [...new Set(colorMatches.map(m => m.match(/([A-F0-9]{6})/)[1]))];
    
    // Extract font sizes
    const fontSizeMatches = documentXml.match(/<w:sz w:val="(\d+)"\/>/g) || [];
    const fontSizes = [...new Set(fontSizeMatches.map(m => {
      const match = m.match(/<w:sz w:val="(\d+)"\/>/);
      return match ? parseInt(match[1]) / 2 : 11; // Convert half-points to points
    }))].sort((a, b) => a - b);
    
    // Extract shading colors
    const shadingMatches = documentXml.match(/<w:shd[^>]*w:fill="([A-F0-9]{6})"[^>]*\/>/g) || [];
    const shadingColors = [...new Set(shadingMatches.map(m => {
      const match = m.match(/w:fill="([A-F0-9]{6})"/);
      return match ? match[1] : null;
    }).filter(c => c && c !== 'auto'))];
    
    // Analyze headings with formatting
    const headings = extractHeadingsWithStyle(documentXml, fullText);
    
    // Parse per-page structure
    const pageStructures = analyzePageStructure(documentXml, estimatedPages);
    
    const textLength = fullText.length;
    const lines = fullText.split('\n').filter(l => l.trim()).length;
    const sections = fullText.split('\n\n').filter(s => s.trim()).length;
    
    const hasTables = tables > 0;
    const hasLists = detectLists(fullText);
    const hasNumberedSections = detectNumberedSections(fullText);
    
    logger.info(`Template Analysis Complete:`);
    logger.info(`   - ${textLength} chars, ${lines} lines`);
    logger.info(`   - Pages detected: ${estimatedPages}`);
    logger.info(`   - Images: ${imageRels}`);
    logger.info(`   - Tables: ${tables}`);
    logger.info(`   - Bold text: ${boldCount} instances`);
    logger.info(`   - Italic text: ${italicCount} instances`);
    logger.info(`   - Colors used: ${colors.length > 0 ? colors.join(', ') : 'None detected'}`);
    logger.info(`   - Headings found: ${headings.length}`);
    
    return {
      fullText,
      textLength,
      lines,
      sections,
      
      // Visual structure analysis
      pages: estimatedPages,
      pageStructures: pageStructures,
      
      // Structural analysis
      structure: {
        hasTables,
        tableCount: tables,
        tableStructures: tableStructures,
        hasLists,
        hasNumberedSections,
        headings,
        headingCount: headings.length,
        
        // Visual elements
        imageCount: imageRels,
        boldCount,
        italicCount,
        colors,
        fontSizes,
        shadingColors,
        
        // Formatting patterns
        hasImages: imageRels > 0,
        hasBoldText: boldCount > 0,
        hasItalicText: italicCount > 0,
        hasColors: colors.length > 0
      },
      
      templatePath
    };
  } catch (error) {
    logger.error(`Template parsing failed: ${error.message}`);
    throw new Error(`Failed to parse template: ${error.message}`);
  }
};

/**
 * Analyze structure of each page in template
 */
function analyzePageStructure(documentXml, pageCount) {
  const pageStructures = [];
  
  // Split by page breaks
  const pageBreakPattern = /<w:br w:type="page"\/>/;
  const xmlPages = documentXml.split(pageBreakPattern);
  
  xmlPages.forEach((pageXml, idx) => {
    if (idx >= pageCount) return;
    
    const structure = {
      pageNumber: idx + 1,
      hasTables: /<w:tbl>/.test(pageXml),
      tableCount: (pageXml.match(/<w:tbl>/g) || []).length,
      hasImages: /<a:blip/.test(pageXml),
      imageCount: (pageXml.match(/<a:blip/g) || []).length,
      boldCount: (pageXml.match(/<w:b\/>/g) || []).length,
      italicCount: (pageXml.match(/<w:i\/>/g) || []).length,
      hasList: /<w:numPr>|<w:ilvl/.test(pageXml),
      
      // Extract headings on this page
      headings: extractHeadingsFromXml(pageXml),
      
      // Paragraph count estimate
      paragraphCount: (pageXml.match(/<w:p>/g) || []).length
    };
    
    pageStructures.push(structure);
  });
  
  return pageStructures;
}

/**
 * Extract headings with style information from XML
 */
function extractHeadingsWithStyle(documentXml, plainText) {
  const headings = [];
  const lines = plainText.split('\n');
  
  lines.forEach(line => {
    const trimmed = line.trim();
    
    if (trimmed.length > 3 && 
        trimmed.length < 100 && 
        /^[A-Z0-9]/.test(trimmed) && 
        !/[.!?]$/.test(trimmed)) {
      headings.push(trimmed);
    }
  });
  
  return headings.slice(0, 30);
}

/**
 * Extract headings from XML chunk
 */
function extractHeadingsFromXml(xmlChunk) {
  // Look for heading styles (pStyle)
  const headingPattern = /<w:pStyle w:val="Heading(\d+)"\/>/g;
  const matches = xmlChunk.match(headingPattern) || [];
  return matches.length;
}

/**
 * Extract table structures from DOCX XML
 */
function extractTableStructures(documentXml) {
  const tables = [];
  const tableRegex = /<w:tbl>[\s\S]*?<\/w:tbl>/g;
  const tableMatches = documentXml.match(tableRegex) || [];
  
  tableMatches.forEach((tableXml, idx) => {
    const rowMatches = tableXml.match(/<w:tr[\s\S]*?<\/w:tr>/g) || [];
    
    const tableInfo = {
      index: idx + 1,
      rows: rowMatches.length,
      cols: 0,
      headers: [],
      hasShading: /<w:shd/.test(tableXml)
    };
    
    // Extract first row as headers
    if (rowMatches.length > 0) {
      const firstRow = rowMatches[0];
      const cellMatches = firstRow.match(/<w:tc[\s\S]*?<\/w:tc>/g) || [];
      tableInfo.cols = cellMatches.length;
      
      // Extract text from header cells
      cellMatches.forEach(cellXml => {
        const textMatches = cellXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
        const cellText = textMatches.map(t => {
          const match = t.match(/<w:t[^>]*>([^<]*)<\/w:t>/);
          return match ? match[1] : '';
        }).join('').trim();
        tableInfo.headers.push(cellText || '[Empty]');
      });
    }
    
    tables.push(tableInfo);
  });
  
  return tables;
}

/**
 * Detect lists in template
 */
function detectLists(text) {
  // Look for bullet points, dashes, or numbered lists
  const listPatterns = [
    /^\s*[-•*]\s+/m,           // Bullet points
    /^\s*\d+\.\s+/m,           // Numbered lists (1. 2. 3.)
    /^\s*[a-z]\)\s+/m,         // Lettered lists (a) b) c)
    /^\s*\([a-z0-9]\)\s+/m,    // Parenthetical lists
  ];
  
  return listPatterns.some(pattern => pattern.test(text));
}

/**
 * Detect numbered sections (e.g., 1.1, 1.2, 2.1)
 */
function detectNumberedSections(text) {
  return /\b\d+\.\d+\b/.test(text);
}

/**
 * Extract potential headings from text
 */
function extractHeadings(text) {
  const headings = [];
  const lines = text.split('\n');
  
  lines.forEach(line => {
    const trimmed = line.trim();
    
    // Heading indicators:
    // - Short line (< 80 chars)
    // - Starts with capital or number
    // - Ends without period (not a sentence)
    // - Not too short (> 3 chars)
    
    if (trimmed.length > 3 && 
        trimmed.length < 80 && 
        /^[A-Z0-9]/.test(trimmed) && 
        !/[.!?]$/.test(trimmed)) {
      headings.push(trimmed);
    }
  });
  
  return headings.slice(0, 20); // Return first 20 potential headings
}

export default {
  parseTemplate
};
