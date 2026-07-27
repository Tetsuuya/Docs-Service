import { generatePptxContent } from '../../../src/services/pptx/scratch/scratchPptxService.js';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test script to verify Gemini generates varied presentation architectures
 * Tests multiple prompts and analyzes patterns
 */

const TEST_PROMPTS = [
  'make a ppt about music',
  'make a ppt about space exploration',
  'make a ppt about artificial intelligence',
  'make a ppt about healthy cooking',
  'make a ppt about climate change',
  'make a ppt about entrepreneurship',
  'make a ppt about yoga and meditation',
  'make a ppt about cybersecurity'
];

async function runTest() {
  console.log('═══════════════════════════════════════════════════════════');
  console.log('🧪 GEMINI ARCHITECTURE VARIETY TEST');
  console.log('═══════════════════════════════════════════════════════════\n');

  const results = [];
  const outputDir = path.join(__dirname, 'output');
  
  // Create output directory
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // Run generation for each prompt
  for (let i = 0; i < TEST_PROMPTS.length; i++) {
    const prompt = TEST_PROMPTS[i];
    console.log(`\n[${i + 1}/${TEST_PROMPTS.length}] Testing: "${prompt}"`);
    console.log('─────────────────────────────────────────────────────────');

    try {
      const startTime = Date.now();
      const result = await generatePptxContent(prompt);
      const duration = ((Date.now() - startTime) / 1000).toFixed(2);

      // Extract key metrics
      const slideCount = result.slides.length;
      const layouts = result.slides.map(s => s.layout || 'list');
      const layoutSequence = layouts.join(' → ');
      const imagesCount = result.slides.filter(s => s.hasImage).length;
      const imageSlides = result.slides
        .map((s, idx) => s.hasImage ? idx + 1 : null)
        .filter(x => x !== null);
      const layoutDistribution = {};
      layouts.forEach(layout => {
        layoutDistribution[layout] = (layoutDistribution[layout] || 0) + 1;
      });

      // Store result
      const testResult = {
        prompt,
        duration: `${duration}s`,
        title: result.title,
        slideCount,
        layouts: layoutSequence,
        layoutDistribution,
        imagesCount,
        imageSlides,
        firstLayout: layouts[0],
        lastLayout: layouts[layouts.length - 1],
        theme: {
          bg: result.theme.backgroundColor,
          accent: result.theme.accentColor
        }
      };

      results.push(testResult);

      // Save full JSON
      const filename = `${i + 1}_${prompt.replace(/[^a-z0-9]/gi, '_').substring(0, 30)}.json`;
      fs.writeFileSync(
        path.join(outputDir, filename),
        JSON.stringify(result, null, 2)
      );

      // Print summary
      console.log(`✅ Generated: "${result.title}"`);
      console.log(`   Slides: ${slideCount} | Images: ${imagesCount} on slides ${imageSlides.join(', ')}`);
      console.log(`   Layouts: ${layoutSequence}`);
      console.log(`   Theme: BG ${result.theme.backgroundColor}, Accent ${result.theme.accentColor}`);
      console.log(`   Time: ${duration}s`);

    } catch (error) {
      console.error(`❌ FAILED: ${error.message}`);
      results.push({
        prompt,
        error: error.message
      });
    }
  }

  // Analyze patterns
  console.log('\n\n═══════════════════════════════════════════════════════════');
  console.log('📊 PATTERN ANALYSIS');
  console.log('═══════════════════════════════════════════════════════════\n');

  const successfulResults = results.filter(r => !r.error);

  if (successfulResults.length === 0) {
    console.log('❌ No successful generations to analyze');
    return;
  }

  // Slide count analysis
  const slideCounts = successfulResults.map(r => r.slideCount);
  const uniqueSlideCounts = [...new Set(slideCounts)];
  console.log('📏 SLIDE COUNTS:');
  console.log(`   Range: ${Math.min(...slideCounts)} - ${Math.max(...slideCounts)}`);
  console.log(`   Unique values: ${uniqueSlideCounts.join(', ')}`);
  console.log(`   Variety score: ${uniqueSlideCounts.length}/${successfulResults.length} (${(uniqueSlideCounts.length / successfulResults.length * 100).toFixed(0)}%)`);

  // Image count analysis
  const imageCounts = successfulResults.map(r => r.imagesCount);
  const uniqueImageCounts = [...new Set(imageCounts)];
  console.log('\n🖼️  IMAGE COUNTS:');
  console.log(`   Range: ${Math.min(...imageCounts)} - ${Math.max(...imageCounts)}`);
  console.log(`   Unique values: ${uniqueImageCounts.join(', ')}`);
  console.log(`   Variety score: ${uniqueImageCounts.length}/${successfulResults.length} (${(uniqueImageCounts.length / successfulResults.length * 100).toFixed(0)}%)`);

  // Layout sequence analysis
  const layoutSequences = successfulResults.map(r => r.layouts);
  const uniqueSequences = [...new Set(layoutSequences)];
  console.log('\n🎨 LAYOUT SEQUENCES:');
  console.log(`   Unique sequences: ${uniqueSequences.length}/${successfulResults.length}`);
  if (uniqueSequences.length < successfulResults.length) {
    const duplicates = layoutSequences.filter((seq, idx) => layoutSequences.indexOf(seq) !== idx);
    console.log(`   ⚠️  DUPLICATES FOUND: ${duplicates.length}`);
    duplicates.forEach(dup => console.log(`      - ${dup}`));
  } else {
    console.log(`   ✅ All sequences are unique!`);
  }

  // First layout analysis
  const firstLayouts = successfulResults.map(r => r.firstLayout);
  const firstLayoutCounts = {};
  firstLayouts.forEach(layout => {
    firstLayoutCounts[layout] = (firstLayoutCounts[layout] || 0) + 1;
  });
  console.log('\n🚀 FIRST SLIDE LAYOUTS:');
  Object.entries(firstLayoutCounts).forEach(([layout, count]) => {
    const percent = (count / successfulResults.length * 100).toFixed(0);
    console.log(`   ${layout}: ${count}x (${percent}%)`);
  });

  // Image placement analysis
  const imagePlacementPatterns = successfulResults.map(r => r.imageSlides.join(','));
  const uniquePlacements = [...new Set(imagePlacementPatterns)];
  console.log('\n📍 IMAGE PLACEMENT PATTERNS:');
  console.log(`   Unique patterns: ${uniquePlacements.length}/${successfulResults.length}`);
  if (imagePlacementPatterns.filter(p => p === '1,2,5').length > 0) {
    console.log(`   ⚠️  Found "1,2,5" pattern (old pattern): ${imagePlacementPatterns.filter(p => p === '1,2,5').length}x`);
  } else {
    console.log(`   ✅ No "1,2,5" pattern detected!`);
  }

  // Overall layout distribution
  const allLayouts = {};
  successfulResults.forEach(r => {
    Object.entries(r.layoutDistribution).forEach(([layout, count]) => {
      allLayouts[layout] = (allLayouts[layout] || 0) + count;
    });
  });
  console.log('\n📊 OVERALL LAYOUT USAGE:');
  Object.entries(allLayouts)
    .sort((a, b) => b[1] - a[1])
    .forEach(([layout, count]) => {
      console.log(`   ${layout}: ${count}x`);
    });

  // Save analysis report
  const report = {
    testDate: new Date().toISOString(),
    totalTests: TEST_PROMPTS.length,
    successful: successfulResults.length,
    failed: results.length - successfulResults.length,
    analysis: {
      slideCounts: {
        range: [Math.min(...slideCounts), Math.max(...slideCounts)],
        unique: uniqueSlideCounts,
        varietyScore: uniqueSlideCounts.length / successfulResults.length
      },
      imageCounts: {
        range: [Math.min(...imageCounts), Math.max(...imageCounts)],
        unique: uniqueImageCounts,
        varietyScore: uniqueImageCounts.length / successfulResults.length
      },
      layoutSequences: {
        unique: uniqueSequences.length,
        total: successfulResults.length,
        allUnique: uniqueSequences.length === successfulResults.length
      },
      firstLayouts: firstLayoutCounts,
      layoutDistribution: allLayouts,
      imagePlacementPatterns: uniquePlacements
    },
    results
  };

  fs.writeFileSync(
    path.join(outputDir, '_ANALYSIS_REPORT.json'),
    JSON.stringify(report, null, 2)
  );

  console.log('\n═══════════════════════════════════════════════════════════');
  console.log('✅ Test complete! Results saved to:');
  console.log(`   ${outputDir}`);
  console.log('═══════════════════════════════════════════════════════════\n');
}

// Run test
runTest().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
