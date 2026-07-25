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
  
  logger.info(`Template Structure: ${templateStructure.pages} pages, ${templateStructure.structure.headingCount} headings, ${templateStructure.structure.tableCount} tables, ${templateStructure.structure.imageCount} images`);
  
  // Build visual structure summary for each page
  let pageVisualSummary = '';
  if (templateStructure.pageStructures && templateStructure.pageStructures.length > 0) {
    pageVisualSummary = templateStructure.pageStructures.map(page => {
      const elements = [];
      if (page.headings > 0) elements.push(`${page.headings} heading(s)`);
      if (page.tableCount > 0) elements.push(`${page.tableCount} table(s)`);
      if (page.imageCount > 0) elements.push(`${page.imageCount} image(s)`);
      if (page.hasList) elements.push('lists');
      if (page.boldCount > 5) elements.push('bold formatting');
      if (page.italicCount > 5) elements.push('italic formatting');
      
      return `Page ${page.pageNumber}: ${elements.length > 0 ? elements.join(', ') : 'text content'}`;
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
  const templatePageCount = templateStructure.pages || 2;
  const targetPages = requestedPages || templatePageCount;
  
  logger.info(`📊 Template: ${templatePageCount} pages → Target: ${targetPages} pages`);
  if (requestedPages) {
    logger.info(`   User requested ${requestedPages} pages - AI will expand/adapt structure`);
  } else {
    logger.info(`   No page count specified - AI will match template depth`);
  }
  
  // Step 1: Universal Dynamic AI Judgment - Gemini decides everything
  const analysisPrompt = `
You are a professional document designer with COMPLETE CREATIVE FREEDOM and HUMAN JUDGMENT.

═══════════════════════════════════════════════════════════
TEMPLATE VISUAL STRUCTURE ANALYSIS
═══════════════════════════════════════════════════════════

Template Pages: ${templatePageCount}
Visual Elements Found:
  • Headings: ${templateStructure.structure.headingCount}
  • Tables: ${templateStructure.structure.tableCount}
  • Images/Diagrams: ${templateStructure.structure.imageCount}
  • Formatting: ${templateStructure.structure.hasBoldText ? '✓ Bold' : ''} ${templateStructure.structure.hasItalicText ? '✓ Italic' : ''}
  • Colors: ${templateStructure.structure.colors.length > 0 ? templateStructure.structure.colors.join(', ') : 'Default'}
  • Lists: ${templateStructure.structure.hasLists ? '✓ Present' : 'None'}

Page-by-Page Layout:
${pageVisualSummary || 'Visual structure not available'}

CRITICAL: Template Table Structures (MUST MATCH):
${tableDescriptions || 'No tables in template'}

Template Headings Found:
${templateStructure.structure.headings.slice(0, 20).join('\n')}

Template Style Sample (first 1500 chars):
"""
${templateText.substring(0, 1500)}
"""

═══════════════════════════════════════════════════════════
USER'S REQUEST
═══════════════════════════════════════════════════════════

Topic: "${userPrompt}"
${requestedPages ? `Target Length: ${requestedPages} pages` : `Target: Match template depth (${templatePageCount} pages)`}

LANGUAGE & TERMINOLOGY:
- Generate the document in English (professional business English)
- If user uses foreign terms (e.g., "sommaire" = French for summary), translate to English equivalents
- Use clear, professional terminology appropriate for business documents

═══════════════════════════════════════════════════════════
YOUR TASK: USE HUMAN JUDGMENT TO DESIGN THE DOCUMENT
═══════════════════════════════════════════════════════════

You are NOT following rigid rules. You are making intelligent design decisions:

1️⃣ ANALYZE THE TEMPLATE'S INTENT
   - What type of document is this? (Report, proposal, manual, study?)
   - What's the professional level? (Corporate, academic, technical, creative?)
   - What visual patterns exist? (Tables for data, images for concepts, lists for features?)

2️⃣ DECIDE CONTENT STRUCTURE FOR USER'S TOPIC
   ${requestedPages 
     ? `- Template has ${templatePageCount} pages, but user wants ${requestedPages} pages`
     : `- Template has ${templatePageCount} pages, match that depth`}
   - How many major sections make sense for "${userPrompt}"?
   - Which sections need more depth? Which need less?
   - Where should tables go? (Data comparisons, metrics, specifications)
   - Where should diagrams go? (Processes, workflows, architectures)
   - IMPORTANT: If user mentions foreign words (e.g., "sommaire"), translate them to English (e.g., "Executive Summary")

3️⃣ MAKE SMART FORMATTING DECISIONS
   - **Bold**: Key terms, metrics, important concepts, emphasis
   - *Italic*: Definitions, technical terms, subtle emphasis
   - \`Code\`: Technical commands, APIs, file names
   - ==Highlight==: Critical warnings, key takeaways, important notes
   - Tables: Comparisons, data, specifications, metrics
   - Lists: Features, steps, requirements, bullet points
   - Diagrams: Processes, workflows, architectures, frameworks

4️⃣ DECIDE CONTENT DEPTH INTELLIGENTLY
   - Complex sections = more paragraphs (4-6)
   - Simple sections = fewer paragraphs (2-3)
   - Technical details = tables + diagrams
   - Narrative sections = more text, fewer visuals
   - Summary sections = lists + key points

🚨 CRITICAL: TABLE STRUCTURE MATCHING
${templateStructure.structure.tableStructures && templateStructure.structure.tableStructures.length > 0 
  ? `Template has ${templateStructure.structure.tableStructures.length} tables. You MUST generate tables with THE EXACT SAME STRUCTURE:

${templateStructure.structure.tableStructures.map(t => 
  `   Table ${t.index}: ${t.cols} columns with headers: ${t.headers.join(' | ')}`
).join('\n')}

For EACH table above:
- Use those EXACT column headers (adapt wording to fit user's topic if needed)
- Generate ${templateStructure.structure.tableStructures.map(t => t.rows - 1).join(', ')} data rows respectively
- Fill with relevant data for "${userPrompt}"
- Example: If template has "Milestone | Deliverables", adapt to topic: "Phase | Key Outcomes"
` 
  : 'Template has no tables - you can add tables where they make sense for the content.'
}

═══════════════════════════════════════════════════════════
RETURN FORMAT
═══════════════════════════════════════════════════════════

Return JSON with as many sections as YOU decide are needed (minimum 3, maximum 150):

{
  "title": "Professional title for user's topic",
  "subtitle": "Compelling subtitle",
  "sections": [
    {
      "heading": "Section heading (you decide based on user's topic)",
      "level": 1-3 (1=major, 2=sub, 3=detail),
      "content": [
        "Paragraph 1 with **bold emphasis** and *italic nuance*...",
        "Paragraph 2 with more \`technical terms\` and ==highlights==...",
        "Paragraph 3...",
        "Add as many paragraphs as THIS SECTION needs (2-6 typically)"
      ],
      "hasList": true/false (your decision - does this section need a list?),
      "listType": "bullet" or "numbered" (if hasList is true),
      "listItems": [
        "**Key point 1**: Explanation with formatting",
        "**Key point 2**: More details"
      ],
      "hasTable": true/false (your decision - does data/comparison make sense here?),
      "table": {
        "title": "Descriptive table title",
        "headers": ["Column 1", "Column 2", "Column 3"],
        "rows": [
          ["Data", "Value", "Status"],
          ["More data", "Another value", "Another status"]
        ]
      },
      "needsDiagram": true/false (your decision - would a visual help explain this?),
      "diagram": {
        "title": "Process/Workflow/Architecture name",
        "steps": ["Step 1", "Step 2", "Step 3", "Step 4"]
      }
    }
  ],
  "metadata": {
    "author": "AI Documentation Engine",
    "keywords": "relevant, professional, keywords"
  }
}

═══════════════════════════════════════════════════════════
CRITICAL GUIDELINES
═══════════════════════════════════════════════════════════

✓ Generate enough sections to properly cover "${userPrompt}" ${requestedPages ? `in ${requestedPages} pages` : ''}
✓ Each paragraph should be 3-5 sentences (substantial, not thin)
✓ Use formatting intelligently (bold/italic/code/highlight) where it adds clarity
✓ Add tables where data/comparisons make the content clearer
✓ Add lists where enumeration helps (features, steps, requirements)
✓ Add diagrams where visual explanation helps (processes, workflows)
✓ Match the template's professional tone and style
✓ Ensure content directly addresses: "${userPrompt}"
✓ Make the document feel cohesive and well-structured

YOU DECIDE: How many sections, how deep each section, where visuals go, how much content each needs.
`;
  
  logger.info(`🤖 Calling AI with human judgment approach...`);
  const structuredContent = await callGemini(analysisPrompt, true, 8192);
  
  const sectionsGenerated = structuredContent.sections?.length || 0;
  logger.info(`✅ AI Generated: ${sectionsGenerated} sections (AI decided this was optimal)`);
  
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
