import { GoogleGenerativeAI } from '@google/generative-ai';
import { config } from '../../../config/env.js';
import { logger } from '../../../utils/logger.js';

const genAI = new GoogleGenerativeAI(config.geminiApiKey);

const candidateModels = [
  config.geminiModel,
  'gemini-flash-lite-latest',
  'gemini-2.0-flash-lite-preview-02-05',
  'gemini-1.5-flash-latest'
];

/**
 * Call Gemini API with model fallback
 */
async function callGemini(promptInput, isJson = true, maxTokens = 8192) {
  let lastError = null;

  for (const modelName of candidateModels) {
    if (!modelName) continue;
    try {
      const model = genAI.getGenerativeModel({
        model: modelName,
        generationConfig: {
          temperature: 1.3, // Very high creativity to prevent generic patterns
          maxOutputTokens: maxTokens,
          ...(isJson ? { responseMimeType: 'application/json' } : {})
        }
      });

      const result = await model.generateContent(promptInput);
      const responseText = result.response.text();
      return isJson ? JSON.parse(responseText) : responseText;
    } catch (err) {
      logger.warn(`PPTX Gemini call with model "${modelName}" failed: ${err.message}`);
      lastError = err;
    }
  }
  throw lastError || new Error('Gemini API calls failed on all candidate models');
}

/**
 * Generate structured presentation JSON content using Gemini AI
 * WITH PROFESSIONAL DESIGN INTELLIGENCE
 */
export const generatePptxContent = async (userPrompt) => {
  if (!config.geminiApiKey || config.geminiApiKey === 'your_gemini_api_key_here') {
    throw new Error('GEMINI_API_KEY is not configured in .env file');
  }

  // Detect slide count request or use intelligent default
  const slideMatch = userPrompt.match(/\b(\d+)\s*(?:-| )?slide/i);
  const targetSlideCount = slideMatch ? Math.min(Math.max(parseInt(slideMatch[1], 10), 3), 20) : null;
  
  // Generate variation constraints to force diversity
  const variationStrategies = [
    { 
      startLayout: 'hero',  // ALWAYS START WITH HERO for title slide
      imageStyle: 'abstract concepts', 
      slideRange: '5-8',
      forbidden: 'NEVER use timeline on slide 5, NEVER end with split',
      imageSlots: [1, 3, 6],
      heroHasImage: true  // Hero slide should have image
    },
    { 
      startLayout: 'hero',  // ALWAYS START WITH HERO for title slide
      imageStyle: 'photographic realism', 
      slideRange: '6-9',
      forbidden: 'NEVER use split on slide 2, NEVER put images on consecutive slides',
      imageSlots: [1, 4, 7],
      heroHasImage: true  // Hero slide should have image
    },
    { 
      startLayout: 'hero',  // ALWAYS START WITH HERO for title slide
      imageStyle: 'close-up details', 
      slideRange: '5-8',
      forbidden: 'NEVER follow hero→split→cards→stat pattern, NEVER use exactly 6 slides',
      imageSlots: [1, 3, 5, 8],
      heroHasImage: true  // Hero slide should have image
    },
    { 
      startLayout: 'hero',  // ALWAYS START WITH HERO for title slide
      imageStyle: 'environmental context', 
      slideRange: '6-10',
      forbidden: 'NEVER put 3 images exactly, NEVER use timeline twice',
      imageSlots: [1, 4, 6],
      heroHasImage: true  // Hero slide should have image
    },
    { 
      startLayout: 'hero',  // ALWAYS START WITH HERO for title slide
      imageStyle: 'action scenes', 
      slideRange: '5-9',
      forbidden: 'NEVER have timeline as 5th slide, NEVER repeat layouts',
      imageSlots: [1, 3, 7],
      heroHasImage: true  // Hero slide should have image
    },
    { 
      startLayout: 'hero',  // ALWAYS START WITH HERO for title slide
      imageStyle: 'before/after sequences', 
      slideRange: '5-7',
      forbidden: 'NEVER generate exactly 7 slides, NEVER use comparison twice',
      imageSlots: [1, 2, 5],
      heroHasImage: true  // Hero slide should have image
    }
  ];
  const randomStrategy = variationStrategies[Math.floor(Math.random() * variationStrategies.length)];
  
  // Randomize image density and placement
  const imageDensity = Math.floor(Math.random() * 4) + 2; // 2-5 images
  const avoidPatterns = [
    'hero → split → cards → stat → timeline → hero',
    'always 6 slides',
    'always images on slides 1, 2, 5',
    'always 3 images total',
    'always start with hero slide',
    'always end with hero slide'
  ];
  
  logger.info(`Generating Professional Presentation for prompt: "${userPrompt}"${targetSlideCount ? ` (Target: ${targetSlideCount} slides)` : ' (AI-determined count)'} [Strategy: ${randomStrategy.startLayout} start, ${imageDensity} images]`);

  const systemPrompt = `
You are an experienced PRESENTATION DESIGNER, not a slide generator.
Your job is to transform ideas into visually compelling, professional presentations that feel handcrafted by an expert.

USER REQUEST: "${userPrompt}"
${targetSlideCount ? `TARGET SLIDE COUNT: ${targetSlideCount} slides` : `SLIDE COUNT: ${randomStrategy.slideRange} slides (choose within this range based on content complexity)`}

🚨 CRITICAL ANTI-PATTERN RULES (MUST FOLLOW):
${randomStrategy.forbidden}

⚠️ FORBIDDEN PATTERNS (these patterns are BANNED):
${avoidPatterns.map(p => `• ${p}`).join('\n')}

🎲 REQUIRED VARIATION FOR THIS GENERATION:
• MUST start with layout: "${randomStrategy.startLayout}" (slide 1)
• Image style theme: "${randomStrategy.imageStyle}"
• Target EXACTLY ${imageDensity} images (no more, no less)
• Suggested image placement: slides ${randomStrategy.imageSlots.join(', ')} (but adjust if better placement exists)
• If slide count is 6, make it 5 or 7 instead
• NEVER repeat the sequence: hero → split → cards → stat → timeline
• Mix up layout order UNPREDICTABLY - surprise me!

═══════════════════════════════════════════════════════════════════
PROFESSIONAL PRESENTATION DESIGN RULES
═══════════════════════════════════════════════════════════════════

🎨 DESIGN PHILOSOPHY:
• Design each slide based on its PURPOSE, not by filling templates
• Ask: What is the single message this slide should communicate?
• Ask: What visual structure best communicates this idea?
• NEVER default to bullet lists unless they're the best choice

📖 STORYTELLING FIRST:
• Build a logical narrative with smooth transitions
• Alternate between information-heavy and visually impactful slides
• Use pacing - sometimes one sentence is enough

🎯 AVAILABLE SLIDE LAYOUTS:
- "hero" - Big statement or title slide
- "stat" - Large number with context
- "split" - Image + key points (left/right split)
- "cards" - 2-3 feature cards or pillars
- "comparison" - Side-by-side A vs B
- "timeline" - Chronological flow
- "process" - Step-by-step sequence
- "quote" - Impactful quote
- "minimal" - Single powerful sentence
- "list" - Bullet points
- "table" - Data grid

🖼️ CRITICAL IMAGE PROMPT RULES (MUST FOLLOW EXACTLY):

🚨 MANDATORY REQUIREMENTS FOR EVERY IMAGE PROMPT:
1. **MUST include 3+ EXACT keywords from the slide's title AND subtitle** (copy them word-for-word!)
2. **MUST describe a SPECIFIC scene matching THIS SLIDE'S EXACT CONTENT** (not general topic)
3. **MUST avoid these BANNED words**: "glowing", "futuristic", "abstract", "concept", "visualization", "neural network", "circuit board", "digital illustration", "technology", "modern", "innovative"
4. **MUST be 80-150 characters** (not too short, not too long)
5. **MUST use photographic realism style** (not artistic/abstract)
6. **Image style theme**: ${randomStrategy.imageStyle}
7. **CRITICAL**: If slide is about "Marine Conservation", image MUST show conservation activities (NOT random ocean). If slide is about "AI Safety", image MUST show safety measures (NOT generic robots).

IMAGE PROMPT FORMULA:
[Subject with action relating to EXACT SLIDE TITLE] + [specific keywords from slide title] + [setting] + [realistic photo]

EXAMPLES (NOTICE HOW KEYWORDS FROM TITLE ARE COPIED):

For slide titled "Acoustic Engineering Principles":
❌ BAD: "abstract sound waves glowing" (no keywords from title, banned words)
❌ BAD: "person in studio" (too generic, missing "acoustic engineering" keywords)
❌ BAD: "microphone recording" (missing "principles" context)
✅ GOOD: "Acoustic engineering lab showing sound principles demonstration with frequency analyzer and panel testing setup"
✅ GOOD: "Engineer demonstrating acoustic principles by measuring room response in engineering test chamber"

For slide titled "Marine Conservation Strategies":
❌ BAD: "beautiful ocean underwater" (doesn't mention conservation!)
❌ BAD: "fish swimming in sea" (misses the conservation focus completely!)
✅ GOOD: "Marine biologist implementing conservation strategies tagging endangered sea turtles for population monitoring"
✅ GOOD: "Conservation team deploying marine protection strategies installing artificial reef structures for habitat restoration"

For slide titled "Neural Network Training Process":
❌ BAD: "abstract neural network glowing" (banned words, not about training)
❌ BAD: "AI concept futuristic" (banned words, too vague)
✅ GOOD: "Data scientist monitoring neural network training process on multiple screens showing loss curves and accuracy metrics"
✅ GOOD: "Research team reviewing training process parameters for neural network optimization in collaborative workspace"

For slide titled "The Foundations of Inner Peace":
❌ BAD: "peaceful meditation concept" (too generic, missing "foundations")
❌ BAD: "zen abstract visualization" (banned words)
✅ GOOD: "Meditation instructor teaching foundations of inner peace breath work techniques to students in morning session"
✅ GOOD: "Person practicing foundational inner peace meditation routine in serene natural garden setting at sunrise"

For slide titled "Economic Impact Analysis":
❌ BAD: "charts and graphs" (no "economic impact" keywords!)
❌ BAD: "business meeting" (too vague)
✅ GOOD: "Economist presenting economic impact analysis data charts to policy makers in government briefing room"
✅ GOOD: "Financial analysts reviewing economic impact metrics and analysis reports comparing regional growth trends"

🎯 STRICT CHECKLIST BEFORE WRITING EACH IMAGE PROMPT:
□ Read the slide's EXACT title and subtitle
□ Identify 3+ specific keywords from title (write them down!)
□ COPY those exact keywords into your image prompt
□ Describe what's happening that relates to THOSE SPECIFIC KEYWORDS
□ NO generic scenes - must match the slide's precise topic
□ Avoid ALL 13 banned words completely
□ Keep 80-150 characters
□ Double-check: Does this image match THIS EXACT SLIDE or just the general topic?

⚠️ COMMON MISTAKES TO AVOID:
• Slide: "Ocean Acidification" → Prompt: "beautiful coral reef" ❌ (doesn't mention acidification!)
• Slide: "Machine Learning Bias" → Prompt: "robots in factory" ❌ (doesn't mention bias!)
• Slide: "Renewable Energy Storage" → Prompt: "solar panels" ❌ (doesn't mention storage!)

✅ CORRECT APPROACH:
• Slide: "Ocean Acidification Effects" → "Marine researcher testing ocean acidification pH levels on damaged coral samples"
• Slide: "Machine Learning Bias Detection" → "Data scientist reviewing machine learning bias audit results highlighting demographic disparities"
• Slide: "Renewable Energy Storage Solutions" → "Engineers inspecting large-scale battery storage facility for renewable energy grid stabilization"

🎨 DYNAMIC COLOR THEMES (pick ONE that matches topic):
- Technology: #0F172A bg, #38B6FF accent
- Nature: #0A2818 bg, #4CAF50 accent
- Business: #1A1A2E bg, #FFD700 accent
- Medical: #001E3C bg, #00BCD4 accent
- Energy: #2C0A00 bg, #FF6F00 accent
- Luxury: #1A0E2E bg, #9C27B0 accent
- Music: #1A0505 bg, #E91E63 accent
- Education: #001F24 bg, #00ACC1 accent

═══════════════════════════════════════════════════════════════════

JSON SCHEMA (return pure JSON, no markdown):
{
  "title": "Presentation Title",
  "subtitle": "Value proposition",
  "theme": {
    "primaryColor": "FFFFFF",
    "secondaryColor": "CBD5E1",
    "accentColor": "E91E63",
    "backgroundColor": "1A0505",
    "cardBgColor": "2A0A0A",
    "textColor": "FFFFFF",
    "fontFamily": "Poppins"
  },
  "slides": [
    {
      "slideNumber": 1,
      "type": "content",
      "layout": "${randomStrategy.startLayout}",
      "title": "Slide Title",
      "subtitle": "Context",
      "hasImage": ${randomStrategy.imageSlots.includes(1)},
      ${randomStrategy.imageSlots.includes(1) ? '"imagePrompt": "HYPER-SPECIFIC prompt matching this exact slide",' : ''}
      // Add required fields for this layout type (see requirements below)
    }
    // ... more slides with VARIED layouts
  ]
}

LAYOUT CONTENT REQUIREMENTS (provide ALL fields for chosen layout):
• hero: title + subtitle (+ optional hasImage/imagePrompt)
• split: title + subtitle + bullets (3-5 items, 15+ words each) + hasImage/imagePrompt
• cards: title + cards array (2-3 cards with title + description 20+ words)
• stat: title + statNumber + statLabel + bullets (context points)
• comparison: title + leftSide object (title + bullets) + rightSide object (title + bullets)
• timeline/process: title + steps array (4-5 steps with title + subtitle + description)
• quote: title + quote (string) + author (string)
• minimal: title + statement (powerful sentence) + support (optional)
• table: title + tableData (headers array + rows 2D array)
• list: title + bullets (3-5 items)

CRITICAL RULES:
1. EVERY slide MUST have complete content for its layout type
2. Write RICH, SPECIFIC content - no placeholders
3. Target EXACTLY ${imageDensity} images total (mark hasImage: true on ${imageDensity} slides)
4. **IMAGE PROMPTS ARE CRITICAL**: 
   - MUST include 2+ words from the slide's title in the prompt
   - MUST describe specific people/objects/actions (not abstract concepts)
   - BANNED WORDS: "glowing", "futuristic", "abstract", "concept", "visualization", "neural network", "circuit board"
   - Use style: ${randomStrategy.imageStyle}
   - Length: 80-150 characters
   - Example: For slide "Acoustic Engineering" → "Audio engineer adjusting acoustic panels in recording studio measuring frequency response"
5. NEVER use forbidden patterns listed above
6. NEVER leave slides empty
7. BE UNPREDICTABLE - vary your choices

Generate NOW with maximum creativity and HYPER-SPECIFIC image prompts.
`;

  const presentationData = await callGemini(systemPrompt, true, 12000);
  
  // Log AI decisions for debugging
  const slideCount = presentationData.slides ? presentationData.slides.length : 0;
  const slidesWithImages = presentationData.slides ? presentationData.slides.filter(s => s.hasImage).length : 0;
  const layouts = presentationData.slides ? presentationData.slides.map(s => s.layout || 'list').join(', ') : '';
  
  logger.info(`📋 AI Design Decisions:`);
  logger.info(`   → Slide Count: ${slideCount}`);
  logger.info(`   → Images Requested: ${slidesWithImages} slides`);
  logger.info(`   → Layouts: ${layouts}`);
  logger.info(`   → Theme: ${presentationData.theme?.backgroundColor || 'default'} BG, ${presentationData.theme?.accentColor || 'default'} accent`);
  
  // Debug: Log slide 5 content if it's a timeline
  if (presentationData.slides && presentationData.slides[4]) {
    const slide5 = presentationData.slides[4];
    if (slide5.layout === 'timeline' || slide5.layout === 'process') {
      logger.info(`🔍 Debug Slide 5 (${slide5.layout}):`, JSON.stringify(slide5, null, 2).substring(0, 300));
    }
  }
  
  return presentationData;
};
