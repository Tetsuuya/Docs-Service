import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Analyze image prompt quality and relevance to slide content
 */

// Quality criteria
const GENERIC_KEYWORDS = [
  'neural network glowing',
  'technology background',
  'abstract concept',
  'digital illustration',
  'futuristic',
  'glowing',
  'blue light',
  'circuit board',
  'spinning vinyl' // overused for music
];

const QUALITY_INDICATORS = {
  specific: {
    keywords: ['close-up', 'showing', 'displaying', 'analyzing', 'during', 'with', 'while', 'consultant', 'engineer', 'doctor', 'professional'],
    weight: 2
  },
  contextual: {
    keywords: ['studio', 'office', 'hospital', 'setting', 'environment', 'workspace', 'facility', 'kitchen', 'laboratory'],
    weight: 1.5
  },
  action: {
    keywords: ['adjusting', 'reviewing', 'operating', 'mixing', 'analyzing', 'performing', 'consulting', 'preparing', 'examining'],
    weight: 2
  },
  technical: {
    keywords: ['high-resolution', 'realistic photography', 'photographic realism', 'macro', 'detail', 'professional'],
    weight: 1
  }
};

function analyzeImagePrompt(prompt, slideTitle, slideSubtitle, slideContent) {
  const analysis = {
    prompt,
    length: prompt.length,
    score: 0,
    maxScore: 10,
    quality: 'Unknown',
    issues: [],
    strengths: []
  };

  // Check for generic keywords (negative score)
  GENERIC_KEYWORDS.forEach(keyword => {
    if (prompt.toLowerCase().includes(keyword.toLowerCase())) {
      analysis.issues.push(`Generic keyword: "${keyword}"`);
      analysis.score -= 2;
    }
  });

  // Check for quality indicators (positive score)
  Object.entries(QUALITY_INDICATORS).forEach(([category, { keywords, weight }]) => {
    keywords.forEach(keyword => {
      if (prompt.toLowerCase().includes(keyword.toLowerCase())) {
        analysis.strengths.push(`${category}: "${keyword}"`);
        analysis.score += weight;
      }
    });
  });

  // Check if prompt is too short (likely generic)
  if (prompt.length < 50) {
    analysis.issues.push('Too short (< 50 chars) - likely generic');
    analysis.score -= 1;
  } else if (prompt.length > 100) {
    analysis.strengths.push('Detailed length (> 100 chars)');
    analysis.score += 1;
  }

  // Check for slide-specific keywords in prompt
  const slideKeywords = [
    ...slideTitle.toLowerCase().split(' '),
    ...(slideSubtitle || '').toLowerCase().split(' ')
  ].filter(word => word.length > 4); // Only words > 4 chars

  let matchCount = 0;
  slideKeywords.forEach(keyword => {
    if (prompt.toLowerCase().includes(keyword)) {
      matchCount++;
    }
  });

  if (matchCount > 0) {
    analysis.strengths.push(`Matches ${matchCount} slide keywords`);
    analysis.score += matchCount * 0.5;
  } else {
    analysis.issues.push('No keywords match slide title/subtitle');
    analysis.score -= 2;
  }

  // Normalize score to 0-10
  analysis.score = Math.max(0, Math.min(10, analysis.score));

  // Determine quality rating
  if (analysis.score >= 8) {
    analysis.quality = '🟢 EXCELLENT';
  } else if (analysis.score >= 6) {
    analysis.quality = '🟡 GOOD';
  } else if (analysis.score >= 4) {
    analysis.quality = '🟠 FAIR';
  } else {
    analysis.quality = '🔴 POOR';
  }

  return analysis;
}

function analyzeAllImagePrompts() {
  console.log('═══════════════════════════════════════════════════════════');
  console.log('🖼️  IMAGE PROMPT QUALITY ANALYSIS');
  console.log('═══════════════════════════════════════════════════════════\n');

  const outputDir = path.join(__dirname, 'output');
  const files = fs.readdirSync(outputDir).filter(f => f.endsWith('.json') && !f.startsWith('_'));

  if (files.length === 0) {
    console.log('❌ No JSON files found. Run test-variety.js first!');
    return;
  }

  const allAnalyses = [];
  let totalImages = 0;

  files.forEach((filename, idx) => {
    const filePath = path.join(outputDir, filename);
    const data = JSON.parse(fs.readFileSync(filePath, 'utf-8'));

    console.log(`\n[${ idx + 1}/${files.length}] ${data.title}`);
    console.log('─────────────────────────────────────────────────────────');

    const slidesWithImages = data.slides.filter(s => s.hasImage);
    console.log(`Found ${slidesWithImages.length} slides with images:\n`);

    slidesWithImages.forEach(slide => {
      totalImages++;
      const analysis = analyzeImagePrompt(
        slide.imagePrompt,
        slide.title,
        slide.subtitle || '',
        JSON.stringify(slide)
      );

      allAnalyses.push({
        presentation: data.title,
        slideNumber: slide.slideNumber,
        slideTitle: slide.title,
        ...analysis
      });

      console.log(`Slide ${slide.slideNumber}: ${slide.title}`);
      console.log(`  Quality: ${analysis.quality} (${analysis.score.toFixed(1)}/10)`);
      console.log(`  Prompt: "${analysis.prompt.substring(0, 80)}${analysis.prompt.length > 80 ? '...' : ''}"`);
      
      if (analysis.strengths.length > 0) {
        console.log(`  ✅ Strengths:`);
        analysis.strengths.forEach(s => console.log(`     • ${s}`));
      }
      
      if (analysis.issues.length > 0) {
        console.log(`  ⚠️  Issues:`);
        analysis.issues.forEach(i => console.log(`     • ${i}`));
      }
      console.log();
    });
  });

  // Overall statistics
  console.log('\n═══════════════════════════════════════════════════════════');
  console.log('📊 OVERALL STATISTICS');
  console.log('═══════════════════════════════════════════════════════════\n');

  const scores = allAnalyses.map(a => a.score);
  const avgScore = scores.reduce((a, b) => a + b, 0) / scores.length;
  const excellentCount = allAnalyses.filter(a => a.score >= 8).length;
  const goodCount = allAnalyses.filter(a => a.score >= 6 && a.score < 8).length;
  const fairCount = allAnalyses.filter(a => a.score >= 4 && a.score < 6).length;
  const poorCount = allAnalyses.filter(a => a.score < 4).length;

  console.log(`Total Images Analyzed: ${totalImages}`);
  console.log(`Average Score: ${avgScore.toFixed(2)}/10\n`);

  console.log('Quality Distribution:');
  console.log(`  🟢 EXCELLENT (8-10): ${excellentCount} (${(excellentCount / totalImages * 100).toFixed(0)}%)`);
  console.log(`  🟡 GOOD (6-8):       ${goodCount} (${(goodCount / totalImages * 100).toFixed(0)}%)`);
  console.log(`  🟠 FAIR (4-6):       ${fairCount} (${(fairCount / totalImages * 100).toFixed(0)}%)`);
  console.log(`  🔴 POOR (0-4):       ${poorCount} (${(poorCount / totalImages * 100).toFixed(0)}%)\n`);

  // Most common issues
  const allIssues = allAnalyses.flatMap(a => a.issues);
  const issueCounts = {};
  allIssues.forEach(issue => {
    const key = issue.split(':')[0]; // Get issue type
    issueCounts[key] = (issueCounts[key] || 0) + 1;
  });

  if (Object.keys(issueCounts).length > 0) {
    console.log('Most Common Issues:');
    Object.entries(issueCounts)
      .sort((a, b) => b[1] - a[1])
      .forEach(([issue, count]) => {
        console.log(`  • ${issue}: ${count}x`);
      });
    console.log();
  }

  // Top and bottom performers
  const sorted = [...allAnalyses].sort((a, b) => b.score - a.score);
  
  console.log('\n🏆 TOP 3 BEST IMAGE PROMPTS:');
  sorted.slice(0, 3).forEach((a, idx) => {
    console.log(`${idx + 1}. Slide ${a.slideNumber} - ${a.slideTitle} (${a.score.toFixed(1)}/10)`);
    console.log(`   "${a.prompt.substring(0, 100)}..."`);
  });

  console.log('\n⚠️  TOP 3 WORST IMAGE PROMPTS:');
  sorted.slice(-3).reverse().forEach((a, idx) => {
    console.log(`${idx + 1}. Slide ${a.slideNumber} - ${a.slideTitle} (${a.score.toFixed(1)}/10)`);
    console.log(`   "${a.prompt.substring(0, 100)}..."`);
  });

  // Save detailed report
  const report = {
    analysisDate: new Date().toISOString(),
    totalImages,
    averageScore: avgScore,
    qualityDistribution: {
      excellent: excellentCount,
      good: goodCount,
      fair: fairCount,
      poor: poorCount
    },
    commonIssues: issueCounts,
    detailedAnalyses: allAnalyses
  };

  fs.writeFileSync(
    path.join(outputDir, '_IMAGE_PROMPT_ANALYSIS.json'),
    JSON.stringify(report, null, 2)
  );

  console.log('\n═══════════════════════════════════════════════════════════');
  console.log('✅ Analysis complete! Detailed report saved to:');
  console.log(`   ${path.join(outputDir, '_IMAGE_PROMPT_ANALYSIS.json')}`);
  console.log('═══════════════════════════════════════════════════════════\n');
}

// Run analysis
analyzeAllImagePrompts();
