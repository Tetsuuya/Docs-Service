import { generatePptxContent } from '../../../src/services/pptx/scratch/scratchPptxService.js';

/**
 * Quick test to verify image prompt improvements
 */

const TEST_PROMPTS = [
  'make a ppt about space exploration',
  'make a ppt about artificial intelligence'
];

async function quickTest() {
  console.log('🧪 QUICK IMAGE PROMPT TEST\n');

  for (const prompt of TEST_PROMPTS) {
    console.log(`\nTesting: "${prompt}"`);
    console.log('─────────────────────────────────────────────────────────');

    try {
      const result = await generatePptxContent(prompt);
      const slidesWithImages = result.slides.filter(s => s.hasImage);

      console.log(`✅ Generated: "${result.title}"`);
      console.log(`   ${slidesWithImages.length} images:\n`);

      slidesWithImages.forEach(slide => {
        const titleWords = slide.title.toLowerCase().split(' ').filter(w => w.length > 4);
        const promptLower = slide.imagePrompt.toLowerCase();
        const matchedWords = titleWords.filter(word => promptLower.includes(word));
        
        // Check for banned words
        const bannedWords = ['glowing', 'futuristic', 'abstract', 'concept', 'visualization', 'neural network', 'circuit board'];
        const foundBanned = bannedWords.filter(word => promptLower.includes(word));

        console.log(`   Slide ${slide.slideNumber}: ${slide.title}`);
        console.log(`   Prompt: "${slide.imagePrompt}"`);
        console.log(`   ✓ Title matches: ${matchedWords.length > 0 ? matchedWords.join(', ') : 'NONE ⚠️'}`);
        console.log(`   ✓ Banned words: ${foundBanned.length > 0 ? foundBanned.join(', ') + ' ❌' : 'None ✅'}`);
        console.log(`   ✓ Length: ${slide.imagePrompt.length} chars ${slide.imagePrompt.length >= 80 && slide.imagePrompt.length <= 150 ? '✅' : '⚠️'}\n`);
      });

    } catch (error) {
      console.error(`❌ FAILED: ${error.message}`);
    }
  }
}

quickTest().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
