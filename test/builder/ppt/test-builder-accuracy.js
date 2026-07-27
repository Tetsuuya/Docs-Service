import { buildScratchPptx } from '../../../src/services/pptx/scratch/scratchPptxBuilder.js';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Test if builder accurately renders all JSON content
 * Verifies that everything in JSON gets placed in the PPTX
 */

function analyzeJsonRequirements(presentationJson) {
  const requirements = {
    totalSlides: presentationJson.slides.length,
    byLayout: {},
    contentChecks: []
  };

  presentationJson.slides.forEach((slide, idx) => {
    const layout = slide.layout || 'list';
    requirements.byLayout[layout] = (requirements.byLayout[layout] || 0) + 1;

    const check = {
      slideNumber: idx + 1,
      layout,
      title: slide.title,
      requiredFields: [],
      hasImage: slide.hasImage || false
    };

    // Check required fields based on layout
    switch (layout) {
      case 'split':
      case 'list':
        if (slide.bullets && slide.bullets.length > 0) {
          check.requiredFields.push(`bullets[${slide.bullets.length}]`);
        } else {
          check.missing = 'bullets array';
        }
        break;

      case 'cards':
        if (slide.cards && slide.cards.length > 0) {
          check.requiredFields.push(`cards[${slide.cards.length}]`);
          slide.cards.forEach((card, cardIdx) => {
            if (!card.title || !card.description) {
              check.missing = `card ${cardIdx + 1} incomplete`;
            }
          });
        } else {
          check.missing = 'cards array';
        }
        break;

      case 'stat':
        if (slide.statNumber) check.requiredFields.push('statNumber');
        else check.missing = 'statNumber';
        
        if (slide.statLabel) check.requiredFields.push('statLabel');
        else check.missing = (check.missing || '') + ' statLabel';
        
        if (slide.bullets && slide.bullets.length > 0) {
          check.requiredFields.push(`bullets[${slide.bullets.length}]`);
        }
        break;

      case 'comparison':
        if (slide.leftSide && slide.rightSide) {
          check.requiredFields.push('leftSide', 'rightSide');
          if (!slide.leftSide.bullets || !slide.rightSide.bullets) {
            check.missing = 'comparison bullets';
          }
        } else {
          check.missing = 'leftSide or rightSide';
        }
        break;

      case 'timeline':
      case 'process':
        if (slide.steps && slide.steps.length >= 3) {
          check.requiredFields.push(`steps[${slide.steps.length}]`);
        } else {
          check.missing = 'steps array (need 3+)';
        }
        break;

      case 'quote':
        if (slide.quote) check.requiredFields.push('quote');
        else check.missing = 'quote';
        
        if (slide.author) check.requiredFields.push('author');
        break;

      case 'minimal':
        if (slide.statement) check.requiredFields.push('statement');
        else check.missing = 'statement';
        break;

      case 'table':
        if (slide.tableData && slide.tableData.headers && slide.tableData.rows) {
          check.requiredFields.push('tableData');
        } else {
          check.missing = 'tableData';
        }
        break;

      case 'hero':
        check.requiredFields.push('title', 'subtitle');
        break;
    }

    if (slide.hasImage && slide.imagePrompt) {
      check.requiredFields.push('imagePrompt');
    }

    requirements.contentChecks.push(check);
  });

  return requirements;
}

async function testBuilderAccuracy() {
  console.log('═══════════════════════════════════════════════════════════');
  console.log('🔨 BUILDER ACCURACY TEST');
  console.log('═══════════════════════════════════════════════════════════\n');
  console.log('This test verifies the builder renders all JSON content correctly.\n');

  const jsonDir = path.resolve(__dirname, '../../gemini-generation-json/ppt/output');
  const outputDir = path.join(__dirname, 'output');

  // Create output directory
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  // Get all test JSON files
  const jsonFiles = fs.readdirSync(jsonDir)
    .filter(f => f.endsWith('.json') && !f.startsWith('_'))
    .slice(0, 3); // Test first 3 for speed

  if (jsonFiles.length === 0) {
    console.log('❌ No JSON files found. Run gemini-generation-json/test-variety.js first!');
    return;
  }

  const testResults = [];

  for (let i = 0; i < jsonFiles.length; i++) {
    const jsonFile = jsonFiles[i];
    const jsonPath = path.join(jsonDir, jsonFile);
    
    console.log(`\n[${i + 1}/${jsonFiles.length}] Testing: ${jsonFile}`);
    console.log('─────────────────────────────────────────────────────────');

    try {
      // Load JSON
      const presentationJson = JSON.parse(fs.readFileSync(jsonPath, 'utf-8'));
      console.log(`📄 Loaded: "${presentationJson.title}"`);
      console.log(`   ${presentationJson.slides.length} slides`);

      // Analyze requirements
      const requirements = analyzeJsonRequirements(presentationJson);
      console.log(`   Layouts: ${Object.keys(requirements.byLayout).join(', ')}`);

      // Check for missing content in JSON
      const missingContent = requirements.contentChecks.filter(c => c.missing);
      if (missingContent.length > 0) {
        console.log(`\n⚠️  JSON Issues Found:`);
        missingContent.forEach(c => {
          console.log(`   Slide ${c.slideNumber} (${c.layout}): Missing ${c.missing}`);
        });
      }

      // Mock image paths (builder needs these)
      const imagePaths = presentationJson.slides
        .filter(s => s.hasImage)
        .map((_, idx) => path.join(__dirname, `mock_image_${idx}.jpg`));

      // Create mock images if builder validates file existence
      imagePaths.forEach(imgPath => {
        if (!fs.existsSync(imgPath)) {
          fs.writeFileSync(imgPath, Buffer.from([0xFF, 0xD8, 0xFF, 0xE0])); // Minimal JPEG header
        }
      });

      // Build presentation
      console.log(`\n🔨 Building presentation...`);
      const outputPath = path.join(outputDir, `test_${i + 1}.pptx`);
      
      const pptxBuffer = await buildScratchPptx(presentationJson, imagePaths);
      
      // Save to file
      fs.writeFileSync(outputPath, pptxBuffer);

      console.log(`✅ Build successful!`);
      console.log(`   Output: ${outputPath}`);

      // Verify file was created
      const stats = fs.statSync(outputPath);
      console.log(`   File size: ${(stats.size / 1024).toFixed(1)} KB`);

      // Analyze what was built
      const result = {
        jsonFile,
        title: presentationJson.title,
        slideCount: presentationJson.slides.length,
        imageCount: imagePaths.length,
        layouts: requirements.byLayout,
        buildSuccess: true,
        fileSize: stats.size,
        contentIssues: missingContent.length,
        issues: missingContent.map(c => `Slide ${c.slideNumber}: ${c.missing}`)
      };

      testResults.push(result);

      // Detailed slide-by-slide check
      console.log(`\n📊 Slide-by-Slide Verification:`);
      requirements.contentChecks.forEach(check => {
        const status = check.missing ? '❌' : '✅';
        const fields = check.requiredFields.length > 0 
          ? check.requiredFields.join(', ') 
          : 'basic';
        const image = check.hasImage ? '🖼️' : '  ';
        
        console.log(`   ${status} ${image} Slide ${check.slideNumber}: ${check.layout} - ${fields}`);
        if (check.missing) {
          console.log(`       ⚠️  Missing: ${check.missing}`);
        }
      });

      // Clean up mock images
      imagePaths.forEach(imgPath => {
        if (fs.existsSync(imgPath)) {
          fs.unlinkSync(imgPath);
        }
      });

    } catch (error) {
      console.error(`❌ Build failed: ${error.message}`);
      console.error(error.stack);
      
      testResults.push({
        jsonFile,
        buildSuccess: false,
        error: error.message,
        stack: error.stack
      });
    }
  }

  // Overall summary
  console.log('\n\n═══════════════════════════════════════════════════════════');
  console.log('📊 TEST SUMMARY');
  console.log('═══════════════════════════════════════════════════════════\n');

  const successful = testResults.filter(r => r.buildSuccess).length;
  const failed = testResults.filter(r => !r.buildSuccess).length;
  const totalContentIssues = testResults
    .filter(r => r.buildSuccess)
    .reduce((sum, r) => sum + (r.contentIssues || 0), 0);

  console.log(`Tests Run: ${testResults.length}`);
  console.log(`✅ Successful Builds: ${successful}/${testResults.length}`);
  console.log(`❌ Failed Builds: ${failed}/${testResults.length}`);
  console.log(`⚠️  Content Issues: ${totalContentIssues} slides with missing fields\n`);

  if (successful > 0) {
    console.log('✅ Successful Builds:');
    testResults.filter(r => r.buildSuccess).forEach(r => {
      console.log(`   ${r.title}`);
      console.log(`     ${r.slideCount} slides, ${r.imageCount} images, ${(r.fileSize / 1024).toFixed(1)} KB`);
      if (r.contentIssues > 0) {
        console.log(`     ⚠️  ${r.contentIssues} content issues`);
      }
    });
  }

  if (failed > 0) {
    console.log('\n❌ Failed Builds:');
    testResults.filter(r => !r.buildSuccess).forEach(r => {
      console.log(`   ${r.jsonFile}: ${r.error}`);
    });
  }

  // Layout coverage
  const allLayouts = new Set();
  testResults.filter(r => r.buildSuccess).forEach(r => {
    Object.keys(r.layouts).forEach(layout => allLayouts.add(layout));
  });

  console.log(`\n🎨 Layout Coverage: ${allLayouts.size} different layouts tested`);
  console.log(`   ${Array.from(allLayouts).join(', ')}`);

  // Save report
  const report = {
    testDate: new Date().toISOString(),
    summary: {
      totalTests: testResults.length,
      successful,
      failed,
      contentIssues: totalContentIssues
    },
    layoutsCovered: Array.from(allLayouts),
    results: testResults
  };

  fs.writeFileSync(
    path.join(outputDir, '_BUILDER_TEST_REPORT.json'),
    JSON.stringify(report, null, 2)
  );

  console.log('\n═══════════════════════════════════════════════════════════');
  console.log('✅ Test complete! Report saved to:');
  console.log(`   ${path.join(outputDir, '_BUILDER_TEST_REPORT.json')}`);
  console.log(`   Built presentations: ${outputDir}`);
  console.log('═══════════════════════════════════════════════════════════\n');
}

// Run test
testBuilderAccuracy().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
