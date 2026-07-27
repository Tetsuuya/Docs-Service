import { callGemini } from '../scratch/scratchModeService.js';
import { logger } from '../../../utils/logger.js';

/**
 * Fill Mode Document Generation
 * Analyzes uploaded template and generates content matching its structure
 * Can extend template structure to match requested page count
 * Separate from scratch mode to avoid conflicts
 */
export const generateFillModeDocument = async (templateText, templateStructure, userPrompt, requestedPages = null) => {
  logger.info('Starting Fill Mode Generation...');
  
  // Safety check
  if (!templateStructure || !templateStructure.structure) {
    logger.error('Template structure is missing or malformed!');
    logger.error(`templateStructure keys: ${Object.keys(templateStructure || {}).join(', ')}`);
    throw new Error('Template structure analysis failed - structure object is missing');
  }
  
  const templateHeadingCount = templateStructure.structure.headingCount;
  const templatePageCount = templateStructure.pages || 2;
  
  logger.info(`Template Structure: ${templatePageCount} pages, ${templateHeadingCount} headings, ${templateStructure.structure.tableCount} tables, ${templateStructure.structure.imageCount} images`);
  
  // WORKAROUND: AI is generating too little content per section
  // Instead of asking for more paragraphs (which AI ignores), ask for MORE SECTIONS
  // Strategy: Multiply section count by 1.5-1.6x (tuned based on testing)
  // Testing shows: 1.3x = 85%, 1.5x = 95%, 1.7x = 121% pages → optimal is 1.55x
  const contentMultiplier = requestedPages && requestedPages > templatePageCount ? 1.6 : 1.4;
  const targetPages = requestedPages || templatePageCount;
  const minSections = Math.min(templateHeadingCount * contentMultiplier, 150); // 1.4-1.6x sections
  const targetSections = requestedPages 
    ? Math.max(minSections, Math.ceil(requestedPages * 2.1)) // 2.1 sections per page
    : minSections;
  
  const paragraphsPerSection = Math.max(3, Math.ceil((targetPages * 2.5) / targetSections)); // Request more
  
  logger.info(`📊 Generation Plan:`);
  logger.info(`   Template: ${templatePageCount} pages with ${templateHeadingCount} headings`);
  logger.info(`   Target: ${targetPages} pages with ${targetSections} sections`);
  logger.info(`   Content: ${paragraphsPerSection} paragraphs per section`);
  logger.info(`   Expected: ~${targetPages * 400} words total`);
  if (requestedPages) {
    logger.info(`   Mode: EXPAND template to ${requestedPages} pages`);
  } else {
    logger.info(`   Mode: PRESERVE template structure (all ${templateHeadingCount} headings)`);
  }
  logger.info(``);
  
  // Build detailed page-by-page structure description
  let pageDescriptions = '';
  if (templateStructure.pageStructures && templateStructure.pageStructures.length > 0) {
    pageDescriptions = templateStructure.pageStructures.map(page => {
      const elements = [];
      
      // Page type
      elements.push(`Type: ${page.pageType || 'content'}`);
      
      // Headings
      if (page.headingCount > 0) elements.push(`${page.headingCount} heading(s)`);
      
      // Content density
      if (page.textDensity) elements.push(`density: ${page.textDensity}`);
      
      // Visual elements
      if (page.tableCount > 0) elements.push(`${page.tableCount} table(s)`);
      if (page.imageCount > 0) elements.push(`${page.imageCount} image(s)`);
      if (page.hasList) elements.push('lists');
      
      // Formatting
      if (page.hasCenteredText) elements.push('centered text');
      if (page.hasLargeSpacing) elements.push('large spacing');
      if (page.boldCount > 5) elements.push('bold formatting');
      
      return `📄 Page ${page.pageNumber}: ${elements.join(', ')}`;
    }).join('\n');
  }
  
  // Build table structure descriptions
  let tableDescriptions = '';
  if (templateStructure.structure.tableStructures && templateStructure.structure.tableStructures.length > 0) {
    tableDescriptions = templateStructure.structure.tableStructures.map(table => {
      return `Table ${table.index}: ${table.rows} rows × ${table.cols} columns
  Headers: ${table.headers.join(' | ')}
  Has shading: ${table.hasShading ? 'Yes' : 'No'}`;
    }).join('\n\n');
  }
  
  // Determine target
  
  // TWO-STEP GENERATION PROCESS
  // Step 1: Generate section outline (lightweight, forces AI to create many sections)
  // Step 2: Fill content in batches (prevents AI from getting lazy/overwhelmed)
  
  logger.info(`🎯 STEP 1: Generating section outline...`);
  
  const outlinePrompt = `
You are a professional document structure designer.

TEMPLATE ANALYSIS:
- Template has ${templatePageCount} pages with ${templateHeadingCount} sections
- User wants ${targetPages} pages about "${userPrompt}"
- Target: ${Math.round(targetSections)} sections minimum

SAMPLE TEMPLATE HEADINGS:
${templateStructure.structure.headings.slice(0, 15).join('\n')}
${templateStructure.structure.headings.length > 15 ? `... and ${templateStructure.structure.headings.length - 15} more` : ''}

YOUR TASK: Create a DETAILED section outline for "${userPrompt}"

REQUIREMENTS:
- Generate ${Math.round(targetSections)}+ section titles
- Follow template's hierarchy pattern (use numbering: 1., 1.1, 1.2, 2., 2.1, etc.)
- Break down the topic into detailed subsections
- Match professional document structure
- Translate foreign terms to English

RETURN JSON:
{
  "title": "Document title",
  "subtitle": "Document subtitle",
  "outline": [
    { "id": 1, "heading": "Executive Summary", "level": 1 },
    { "id": 2, "heading": "Key Points", "level": 2 },
    { "id": 3, "heading": "1. Introduction", "level": 1 },
    { "id": 4, "heading": "1.1 Purpose", "level": 2 },
    ... generate ${Math.round(targetSections)}+ sections
  ]
}

CRITICAL: Generate AT LEAST ${Math.round(targetSections)} sections. More is better than less.
`;

  const outlineResponse = await callGemini(outlinePrompt, true, 4096);
  const sectionOutline = outlineResponse.outline || [];
  
  logger.info(`✅ STEP 1 Complete: ${sectionOutline.length} sections outlined`);
  
  if (sectionOutline.length < targetSections * 0.7) {
    logger.warn(`⚠️  Only ${sectionOutline.length} sections generated (target: ${Math.round(targetSections)})`);
  }
  
  // Step 2: Fill content in batches
  logger.info(`🎯 STEP 2: Filling content in batches...`);
  
  const batchSize = 15; // Process 15 sections at a time
  const batches = Math.ceil(sectionOutline.length / batchSize);
  const filledSections = [];
  
  for (let batchIndex = 0; batchIndex < batches; batchIndex++) {
    const startIdx = batchIndex * batchSize;
    const endIdx = Math.min(startIdx + batchSize, sectionOutline.length);
    const batchSections = sectionOutline.slice(startIdx, endIdx);
    
    logger.info(`   Batch ${batchIndex + 1}/${batches}: Filling sections ${startIdx + 1}-${endIdx}...`);
    
    const fillPrompt = `
You are filling content for a professional document about "${userPrompt}".

SECTIONS TO FILL (${batchSections.length} sections):
${batchSections.map(s => `${s.id}. ${s.heading} (Level ${s.level})`).join('\n')}

TEMPLATE INFO:
- Tables available: ${templateStructure.structure.tableCount}
- Uses bold and italic formatting
- Professional business style

CONTENT REQUIREMENTS:
- Each section: ${paragraphsPerSection}+ paragraphs (60-100 words each)
- Use **bold** for key terms
- Use *italic* for emphasis
- Add tables where data makes sense
- Add lists for enumeration

RETURN JSON:
{
  "sections": [
    {
      "id": ${batchSections[0]?.id},
      "heading": "${batchSections[0]?.heading}",
      "level": ${batchSections[0]?.level},
      "content": [
        "Paragraph 1 with **bold** and *italic* (60-100 words)...",
        "Paragraph 2 (60-100 words)...",
        "Paragraph 3 (60-100 words)..."
      ],
      "hasList": false,
      "hasTable": false
    }
    ... fill all ${batchSections.length} sections
  ]
}
`;

    const batchResponse = await callGemini(fillPrompt, true, 8192);
    filledSections.push(...(batchResponse.sections || []));
    
    logger.info(`   ✅ Batch ${batchIndex + 1} complete: ${batchResponse.sections?.length || 0} sections filled`);
  }
  
  logger.info(`✅ STEP 2 Complete: ${filledSections.length} sections filled with content`);
  
  // Build final structured content
  const structuredContent = {
    title: outlineResponse.title || 'Generated Document',
    subtitle: outlineResponse.subtitle || 'Professional Report',
    sections: filledSections,
    metadata: {
      author: 'AI Documentation Engine',
      keywords: userPrompt.split(' ').slice(0, 5).join(', ')
    }
  };
  
  const sectionsGenerated = structuredContent.sections?.length || 0;
  const preservationRatio = templateHeadingCount > 0 ? (sectionsGenerated / templateHeadingCount) : 1;
  
  // Calculate estimated word count
  let estimatedWords = 0;
  let paragraphCount = 0;
  let tableCount = 0;
  let listCount = 0;
  let diagramCount = 0;
  
  structuredContent.sections?.forEach(section => {
    section.content?.forEach(para => {
      estimatedWords += para.split(' ').length;
      paragraphCount++;
    });
    section.listItems?.forEach(item => {
      estimatedWords += item.split(' ').length;
    });
    if (section.hasTable) tableCount++;
    if (section.hasList) listCount++;
    if (section.needsDiagram) diagramCount++;
  });
  
  const targetWords = targetPages * 400; // 400 words per page
  const contentRatio = targetWords > 0 ? (estimatedWords / targetWords) : 1;
  const estimatedPages = Math.round(estimatedWords / 400);
  
  logger.info(`\n═══════════════════════════════════════════════════════════`);
  logger.info(`📊 FILL MODE GENERATION ANALYSIS`);
  logger.info(`═══════════════════════════════════════════════════════════`);
  
  // Template vs Output comparison
  logger.info(`\n📄 STRUCTURE COMPARISON:`);
  logger.info(`   Template: ${templatePageCount} pages, ${templateHeadingCount} headings, ${templateStructure.structure.tableCount} tables`);
  logger.info(`   Generated: ${estimatedPages} pages (est), ${sectionsGenerated} sections, ${tableCount} tables`);
  logger.info(`   Match: ${Math.round(preservationRatio * 100)}% structure depth preserved`);
  
  // Content volume analysis
  logger.info(`\n📝 CONTENT VOLUME:`);
  logger.info(`   Target: ~${targetWords} words (${targetPages} pages × 400 words/page)`);
  logger.info(`   Generated: ~${estimatedWords} words`);
  logger.info(`   Ratio: ${Math.round(contentRatio * 100)}% of target`);
  logger.info(`   Estimated Pages: ${estimatedPages} (vs ${targetPages} target)`);
  
  // Content breakdown
  logger.info(`\n🔍 CONTENT BREAKDOWN:`);
  logger.info(`   Sections: ${sectionsGenerated}`);
  logger.info(`   Paragraphs: ${paragraphCount} (avg ${(paragraphCount / sectionsGenerated).toFixed(1)} per section)`);
  logger.info(`   Tables: ${tableCount}`);
  logger.info(`   Lists: ${listCount}`);
  logger.info(`   Diagrams: ${diagramCount}`);
  logger.info(`   Avg words/section: ${(estimatedWords / sectionsGenerated).toFixed(0)}`);
  
  // Structure validation
  logger.info(`\n✅ VALIDATION RESULTS:`);
  if (preservationRatio < 0.7) {
    logger.warn(`   ⚠️  STRUCTURE: Too condensed (${Math.round(preservationRatio * 100)}% < 70%)`);
    logger.warn(`      → AI simplified template structure - missing ${templateHeadingCount - sectionsGenerated} sections`);
  } else if (preservationRatio >= 0.9 && preservationRatio <= 1.1) {
    logger.info(`   ✅ STRUCTURE: Excellent match (${Math.round(preservationRatio * 100)}%)`);
  } else if (preservationRatio > 1.1) {
    logger.info(`   ✅ STRUCTURE: Expanded appropriately (${Math.round(preservationRatio * 100)}%)`);
  } else {
    logger.info(`   ✅ STRUCTURE: Good match (${Math.round(preservationRatio * 100)}%)`);
  }
  
  // Content density validation
  if (contentRatio < 0.7) {
    logger.warn(`   ⚠️  CONTENT: Too sparse (${Math.round(contentRatio * 100)}% < 70%)`);
    logger.warn(`      → Document will be ${Math.abs(estimatedPages - targetPages)} pages shorter than target`);
    logger.warn(`      → Need ~${targetWords - estimatedWords} more words`);
  } else if (contentRatio > 1.3) {
    logger.warn(`   ⚠️  CONTENT: Too dense (${Math.round(contentRatio * 100)}% > 130%)`);
    logger.warn(`      → Document will be ${Math.abs(estimatedPages - targetPages)} pages longer than target`);
    logger.warn(`      → Consider reducing content by ~${estimatedWords - targetWords} words`);
  } else if (contentRatio < 0.85 || contentRatio > 1.15) {
    logger.info(`   ⚠️  CONTENT: Slightly off target (${Math.round(contentRatio * 100)}%)`);
    logger.info(`      → Expected ${targetPages} pages, got ~${estimatedPages} pages`);
  } else {
    logger.info(`   ✅ CONTENT: Well balanced (${Math.round(contentRatio * 100)}%)`);
    logger.info(`      → Page count matches target (${estimatedPages} ≈ ${targetPages})`);
  }
  
  // Page count specific warning
  const pageDifference = Math.abs(estimatedPages - targetPages);
  if (pageDifference >= 3) {
    logger.warn(`\n⚠️  PAGE COUNT ISSUE:`);
    logger.warn(`   Template: ${templatePageCount} pages → Output: ~${estimatedPages} pages`);
    logger.warn(`   Difference: ${pageDifference} pages (${estimatedPages < targetPages ? 'TOO SHORT' : 'TOO LONG'})`);
    if (estimatedPages < targetPages) {
      logger.warn(`   💡 SOLUTION: AI needs to generate more paragraphs per section`);
      logger.warn(`      → Each section should have ${Math.ceil((targetWords / sectionsGenerated) / 80)} paragraphs`);
    } else {
      logger.warn(`   💡 SOLUTION: AI generated too much content per section`);
    }
  }
  
  logger.info(`═══════════════════════════════════════════════════════════\n`);
  
  // Step 2: Build document JSON matching our docx builder format
  const documentData = {
    title: structuredContent.title || 'Generated Document',
    subtitle: structuredContent.subtitle || 'Based on uploaded template',
    author: structuredContent.metadata?.author || 'AI Documentation Engine',
    keywords: structuredContent.metadata?.keywords || 'document, template, generated',
    mode: 'fill',
    requestedPages: requestedPages || null,
    templatePages: templatePageCount,
    sectionsGenerated: sectionsGenerated,
    
    // Pass template structure for header/footer adaptation
    templateStructure: templateStructure,
    
    // Pass template styling to builder
    templateStyle: {
      colors: templateStructure.structure.colors || [],
      fontSizes: templateStructure.structure.fontSizes || [],
      shadingColors: templateStructure.structure.shadingColors || [],
      hasBold: templateStructure.structure.hasBoldText,
      hasItalic: templateStructure.structure.hasItalicText
    },
    
    // Theme - professional and clean
    theme: {
      primaryColor: '1E40AF',
      secondaryColor: '3B82F6',
      accentColor: '60A5FA',
      lightBgColor: 'F3F4F6',
      textColor: '1F2937'
    },
    
    // Convert sections to our page format
    pages: structuredContent.sections?.map((section, idx) => ({
      heading: section.heading,
      headingLevel: section.level || 2,
      headingAlignment: section.level === 1 ? 'center' : 'left',
      
      paragraphs: section.content || [],
      
      bulletList: section.hasList ? section.listItems : null,
      listType: section.hasList ? (section.listType || 'bullet') : null,
      
      table: section.hasTable ? {
        title: section.table?.title || `${section.heading} - Data`,
        headers: section.table?.headers || [],
        rows: section.table?.rows || []
      } : null,
      
      visualNeed: section.needsDiagram ? {
        type: 'diagram',
        diagram: {
          title: section.diagram?.title || 'Process Overview',
          steps: section.diagram?.steps || []
        }
      } : { type: 'none' },
      
      horizontalDivider: idx < (structuredContent.sections?.length - 1)
    })) || [],
    
    sections: structuredContent.sections || []
  };
  
  logger.info(`📄 Fill Mode Complete -> ${documentData.pages.length} sections structured`);
  logger.info(`   - Tables: ${documentData.pages.filter(p => p.table).length}`);
  logger.info(`   - Lists: ${documentData.pages.filter(p => p.bulletList).length}`);
  logger.info(`   - Diagrams: ${documentData.pages.filter(p => p.visualNeed?.type === 'diagram').length}`);
  logger.info(`   - ✅ AI used judgment to create optimal structure`);
  
  return documentData;
};

export default {
  generateFillModeDocument
};
