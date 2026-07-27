import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Analyze content quality - check if content is detailed, not generic/placeholder
 */

// Red flags for poor content
const PLACEHOLDER_INDICATORS = [
  'lorem ipsum',
  'placeholder',
  'sample text',
  'insert text here',
  'add content',
  'example text',
  'tbd',
  'to be determined'
];

const GENERIC_PHRASES = [
  'key benefit',
  'important feature',
  'main advantage',
  'critical component',
  'essential element',
  'primary focus',
  'core value'
];

function analyzeContentQuality(slide, slideIdx, presentationTitle) {
  const analysis = {
    slideNumber: slideIdx + 1,
    layout: slide.layout,
    title: slide.title,
    issues: [],
    warnings: [],
    strengths: [],
    score: 10,
    details: {}
  };

  // Check based on layout type
  switch (slide.layout) {
    case 'split':
    case 'list':
      if (!slide.bullets || slide.bullets.length === 0) {
        analysis.issues.push('Missing bullets array');
        analysis.score -= 5;
      } else {
        analysis.details.bulletCount = slide.bullets.length;
        
        // Check bullet quality
        const avgBulletLength = slide.bullets.reduce((sum, b) => sum + b.length, 0) / slide.bullets.length;
        analysis.details.avgBulletLength = Math.round(avgBulletLength);
        
        if (avgBulletLength < 50) {
          analysis.warnings.push(`Short bullets (avg ${Math.round(avgBulletLength)} chars, target: 80+)`);
          analysis.score -= 2;
        } else if (avgBulletLength >= 80) {
          analysis.strengths.push(`Detailed bullets (avg ${Math.round(avgBulletLength)} chars)`);
          analysis.score += 1;
        }
        
        // Check for generic content
        slide.bullets.forEach((bullet, idx) => {
          const lower = bullet.toLowerCase();
          PLACEHOLDER_INDICATORS.forEach(ph => {
            if (lower.includes(ph)) {
              analysis.issues.push(`Bullet ${idx + 1} has placeholder: "${ph}"`);
              analysis.score -= 3;
            }
          });
          
          GENERIC_PHRASES.forEach(gp => {
            if (lower.includes(gp)) {
              analysis.warnings.push(`Bullet ${idx + 1} uses generic phrase: "${gp}"`);
              analysis.score -= 0.5;
            }
          });
        });
      }
      break;

    case 'cards':
      if (!slide.cards || slide.cards.length === 0) {
        analysis.issues.push('Missing cards array');
        analysis.score -= 5;
      } else {
        analysis.details.cardCount = slide.cards.length;
        
        // Check each card
        slide.cards.forEach((card, idx) => {
          if (!card.title) {
            analysis.issues.push(`Card ${idx + 1} missing title`);
            analysis.score -= 2;
          }
          if (!card.description) {
            analysis.issues.push(`Card ${idx + 1} missing description`);
            analysis.score -= 2;
          } else {
            const descLength = card.description.length;
            if (descLength < 50) {
              analysis.warnings.push(`Card ${idx + 1} description too short (${descLength} chars, target: 80+)`);
              analysis.score -= 1;
            } else if (descLength >= 80) {
              analysis.strengths.push(`Card ${idx + 1} detailed (${descLength} chars)`);
            }
            
            // Check for placeholders
            const lower = card.description.toLowerCase();
            PLACEHOLDER_INDICATORS.forEach(ph => {
              if (lower.includes(ph)) {
                analysis.issues.push(`Card ${idx + 1} has placeholder: "${ph}"`);
                analysis.score -= 3;
              }
            });
          }
        });
      }
      break;

    case 'stat':
      if (!slide.statNumber) {
        analysis.issues.push('Missing statNumber');
        analysis.score -= 3;
      } else {
        analysis.details.statNumber = slide.statNumber;
        analysis.strengths.push(`Has stat: ${slide.statNumber}`);
      }
      
      if (!slide.statLabel) {
        analysis.issues.push('Missing statLabel');
        analysis.score -= 2;
      }
      
      if (!slide.bullets || slide.bullets.length < 2) {
        analysis.warnings.push('Stat slide should have 2-4 supporting bullets');
        analysis.score -= 1;
      }
      break;

    case 'comparison':
      if (!slide.leftSide || !slide.rightSide) {
        analysis.issues.push('Missing leftSide or rightSide');
        analysis.score -= 5;
      } else {
        if (!slide.leftSide.bullets || slide.leftSide.bullets.length < 2) {
          analysis.warnings.push('Left side needs more bullets');
          analysis.score -= 1;
        }
        if (!slide.rightSide.bullets || slide.rightSide.bullets.length < 2) {
          analysis.warnings.push('Right side needs more bullets');
          analysis.score -= 1;
        }
        
        if (slide.leftSide.bullets && slide.rightSide.bullets) {
          analysis.details.comparisonBalance = `${slide.leftSide.bullets.length} vs ${slide.rightSide.bullets.length}`;
          if (Math.abs(slide.leftSide.bullets.length - slide.rightSide.bullets.length) > 1) {
            analysis.warnings.push('Comparison sides unbalanced');
          } else {
            analysis.strengths.push('Balanced comparison');
          }
        }
      }
      break;

    case 'timeline':
    case 'process':
      if (!slide.steps || slide.steps.length < 3) {
        analysis.issues.push('Timeline/process needs 3+ steps');
        analysis.score -= 3;
      } else {
        analysis.details.stepCount = slide.steps.length;
        
        slide.steps.forEach((step, idx) => {
          if (!step.title) {
            analysis.issues.push(`Step ${idx + 1} missing title`);
            analysis.score -= 1;
          }
          if (!step.description) {
            analysis.warnings.push(`Step ${idx + 1} missing description`);
            analysis.score -= 0.5;
          } else if (step.description.length < 30) {
            analysis.warnings.push(`Step ${idx + 1} description too short`);
          }
        });
        
        if (slide.steps.length >= 4) {
          analysis.strengths.push(`Detailed process (${slide.steps.length} steps)`);
        }
      }
      break;

    case 'quote':
      if (!slide.quote) {
        analysis.issues.push('Missing quote text');
        analysis.score -= 4;
      } else {
        analysis.details.quoteLength = slide.quote.length;
        if (slide.quote.length < 30) {
          analysis.warnings.push('Quote too short');
          analysis.score -= 1;
        } else {
          analysis.strengths.push('Substantive quote');
        }
      }
      
      if (!slide.author) {
        analysis.warnings.push('Missing quote author');
        analysis.score -= 1;
      }
      break;

    case 'minimal':
      if (!slide.statement) {
        analysis.issues.push('Missing statement');
        analysis.score -= 4;
      } else {
        analysis.details.statementLength = slide.statement.length;
        if (slide.statement.length < 20) {
          analysis.warnings.push('Statement too short');
        }
      }
      break;

    case 'table':
      if (!slide.tableData || !slide.tableData.headers || !slide.tableData.rows) {
        analysis.issues.push('Missing table data');
        analysis.score -= 5;
      } else {
        analysis.details.tableSize = `${slide.tableData.headers.length} cols × ${slide.tableData.rows.length} rows`;
        analysis.strengths.push('Has table data');
      }
      break;

    case 'hero':
      // Hero just needs title + subtitle
      if (!slide.subtitle) {
        analysis.warnings.push('Hero slide missing subtitle');
        analysis.score -= 0.5;
      }
      analysis.strengths.push('Hero slide (title-focused)');
      break;
  }

  // Normalize score
  analysis.score = Math.max(0, Math.min(10, analysis.score));
  
  // Quality rating
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

function analyzeAll() {
  console.log('═══════════════════════════════════════════════════════════');
  console.log('📝 CONTENT QUALITY ANALYSIS');
  console.log('═══════════════════════════════════════════════════════════\n');

  const outputDir = path.join(__dirname, 'output');
  const files = fs.readdirSync(outputDir).filter(f => f.endsWith('.json') && !f.startsWith('_'));

  if (files.length === 0) {
    console.log('❌ No JSON files found. Run test-variety.js first!');
    return;
  }

  const allAnalyses = [];
  let totalSlides = 0;

  files.forEach((filename, idx) => {
    const filePath = path.join(outputDir, filename);
    const data = JSON.parse(fs.readFileSync(filePath, 'utf-8'));

    console.log(`\n[${idx + 1}/${files.length}] ${data.title}`);
    console.log('─────────────────────────────────────────────────────────');

    data.slides.forEach((slide, slideIdx) => {
      totalSlides++;
      const analysis = analyzeContentQuality(slide, slideIdx, data.title);
      allAnalyses.push({
        presentation: data.title,
        ...analysis
      });

      console.log(`Slide ${analysis.slideNumber}: ${slide.title} (${slide.layout})`);
      console.log(`  Quality: ${analysis.quality} (${analysis.score.toFixed(1)}/10)`);
      
      if (Object.keys(analysis.details).length > 0) {
        console.log(`  Details: ${JSON.stringify(analysis.details)}`);
      }
      
      if (analysis.strengths.length > 0) {
        console.log(`  ✅ ${analysis.strengths.join(', ')}`);
      }
      
      if (analysis.warnings.length > 0) {
        console.log(`  ⚠️  ${analysis.warnings.join(', ')}`);
      }
      
      if (analysis.issues.length > 0) {
        console.log(`  🔴 ${analysis.issues.join(', ')}`);
      }
      console.log();
    });
  });

  // Overall statistics
  console.log('\n═══════════════════════════════════════════════════════════');
  console.log('📊 OVERALL CONTENT STATISTICS');
  console.log('═══════════════════════════════════════════════════════════\n');

  const scores = allAnalyses.map(a => a.score);
  const avgScore = scores.reduce((a, b) => a + b, 0) / scores.length;
  const excellentCount = allAnalyses.filter(a => a.score >= 8).length;
  const goodCount = allAnalyses.filter(a => a.score >= 6 && a.score < 8).length;
  const fairCount = allAnalyses.filter(a => a.score >= 4 && a.score < 6).length;
  const poorCount = allAnalyses.filter(a => a.score < 4).length;

  console.log(`Total Slides Analyzed: ${totalSlides}`);
  console.log(`Average Content Score: ${avgScore.toFixed(2)}/10\n`);

  console.log('Quality Distribution:');
  console.log(`  🟢 EXCELLENT (8-10): ${excellentCount} (${(excellentCount / totalSlides * 100).toFixed(0)}%)`);
  console.log(`  🟡 GOOD (6-8):       ${goodCount} (${(goodCount / totalSlides * 100).toFixed(0)}%)`);
  console.log(`  🟠 FAIR (4-6):       ${fairCount} (${(fairCount / totalSlides * 100).toFixed(0)}%)`);
  console.log(`  🔴 POOR (0-4):       ${poorCount} (${(poorCount / totalSlides * 100).toFixed(0)}%)\n`);

  // Common issues
  const allIssues = allAnalyses.flatMap(a => a.issues);
  const allWarnings = allAnalyses.flatMap(a => a.warnings);

  if (allIssues.length > 0) {
    console.log('🔴 Critical Issues Found:');
    const issueCounts = {};
    allIssues.forEach(issue => {
      const key = issue.split(':')[0];
      issueCounts[key] = (issueCounts[key] || 0) + 1;
    });
    Object.entries(issueCounts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .forEach(([issue, count]) => {
        console.log(`  • ${issue}: ${count}x`);
      });
    console.log();
  }

  if (allWarnings.length > 0) {
    console.log('⚠️  Common Warnings:');
    const warningCounts = {};
    allWarnings.forEach(warning => {
      const key = warning.split('(')[0].trim();
      warningCounts[key] = (warningCounts[key] || 0) + 1;
    });
    Object.entries(warningCounts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .forEach(([warning, count]) => {
        console.log(`  • ${warning}: ${count}x`);
      });
    console.log();
  }

  // Best and worst
  const sorted = [...allAnalyses].sort((a, b) => b.score - a.score);
  
  console.log('🏆 TOP 3 BEST CONTENT:');
  sorted.slice(0, 3).forEach((a, idx) => {
    console.log(`${idx + 1}. Slide ${a.slideNumber} - ${a.title} (${a.layout}) - ${a.score.toFixed(1)}/10`);
  });

  console.log('\n⚠️  TOP 3 WORST CONTENT:');
  sorted.slice(-3).reverse().forEach((a, idx) => {
    console.log(`${idx + 1}. Slide ${a.slideNumber} - ${a.title} (${a.layout}) - ${a.score.toFixed(1)}/10`);
  });

  // Save report
  const report = {
    analysisDate: new Date().toISOString(),
    totalSlides,
    averageScore: avgScore,
    qualityDistribution: {
      excellent: excellentCount,
      good: goodCount,
      fair: fairCount,
      poor: poorCount
    },
    detailedAnalyses: allAnalyses
  };

  fs.writeFileSync(
    path.join(outputDir, '_CONTENT_QUALITY_ANALYSIS.json'),
    JSON.stringify(report, null, 2)
  );

  console.log('\n═══════════════════════════════════════════════════════════');
  console.log('✅ Analysis complete! Report saved to:');
  console.log(`   ${path.join(outputDir, '_CONTENT_QUALITY_ANALYSIS.json')}`);
  console.log('═══════════════════════════════════════════════════════════\n');
}

analyzeAll();
