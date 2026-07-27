import pptxgen from 'pptxgenjs';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Generate 3 PROFESSIONAL template previews
 * Designed like real human-made presentations (Apple, TED, Corporate style)
 */

// ═══════════════════════════════════════════════════════════
// TEMPLATE 1: CLEAN CORPORATE (Apple/Microsoft style)
// ═══════════════════════════════════════════════════════════
async function generateTemplate1() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = 'Template 1 - Clean Corporate';

  const bg = 'FFFFFF';
  const accent = '007AFF';
  const primary = '1D1D1F';
  const secondary = '86868B';
  const cardBg = 'F5F5F7';

  // SLIDE 1: Title Slide (Clean, centered, minimalist)
  const slide1 = pptx.addSlide();
  slide1.background = { color: bg };
  
  // Centered title
  slide1.addText('The Future of Innovation', { 
    x: 1.5, y: 2.0, w: 7, h: 1.0, 
    fontSize: 48, bold: true, color: primary, align: 'center',
    fontFace: 'Arial'
  });
  
  // Subtitle
  slide1.addText('A comprehensive overview of emerging trends', { 
    x: 2.0, y: 3.2, w: 6, h: 0.5, 
    fontSize: 18, color: secondary, align: 'center',
    fontFace: 'Arial'
  });
  
  // Small accent line under subtitle
  slide1.addShape('rect', {
    x: 4.5, y: 3.85, w: 1, h: 0.05,
    fill: { color: accent }
  });
  
  // Bottom right: Company/Date
  slide1.addText('Q1 2026', {
    x: 7.5, y: 5.0, w: 2, h: 0.3,
    fontSize: 12, color: secondary, align: 'right',
    fontFace: 'Arial'
  });

  // SLIDE 2: Big Statement (Single impactful message)
  const slide2 = pptx.addSlide();
  slide2.background = { color: bg };
  
  // Large centered statement
  slide2.addText('Simple.', { 
    x: 1, y: 1.8, w: 8, h: 1.5, 
    fontSize: 96, bold: true, color: accent, align: 'center',
    fontFace: 'Arial'
  });
  
  slide2.addText('Yet powerful enough to transform your business.', { 
    x: 1.5, y: 3.5, w: 7, h: 0.6, 
    fontSize: 20, color: secondary, align: 'center',
    fontFace: 'Arial'
  });

  // SLIDE 3: Three Column Features (Clean cards)
  const slide3 = pptx.addSlide();
  slide3.background = { color: bg };
  
  // Title
  slide3.addText('Key Benefits', {
    x: 0.5, y: 0.5, w: 9, h: 0.5,
    fontSize: 32, bold: true, color: primary,
    fontFace: 'Arial'
  });
  
  // Three clean feature cards
  const features = [
    { icon: '⚡', title: 'Fast', desc: 'Lightning-fast performance that scales with your needs' },
    { icon: '🔒', title: 'Secure', desc: 'Enterprise-grade security built into every layer' },
    { icon: '📊', title: 'Insightful', desc: 'Real-time analytics and actionable intelligence' }
  ];
  
  features.forEach((feat, i) => {
    const x = 0.5 + i * 3.2;
    
    // Card background
    slide3.addShape('roundRect', {
      x, y: 1.5, w: 2.9, h: 3.5,
      fill: { color: cardBg },
      line: { type: 'none' },
      rectRadius: 0.15
    });
    
    // Icon
    slide3.addText(feat.icon, {
      x: x + 0.3, y: 2.0, w: 2.3, h: 0.6,
      fontSize: 48, align: 'center'
    });
    
    // Title
    slide3.addText(feat.title, {
      x: x + 0.3, y: 2.8, w: 2.3, h: 0.4,
      fontSize: 20, bold: true, color: primary, align: 'center',
      fontFace: 'Arial'
    });
    
    // Description
    slide3.addText(feat.desc, {
      x: x + 0.3, y: 3.4, w: 2.3, h: 1.2,
      fontSize: 13, color: secondary, align: 'center', wrap: true,
      fontFace: 'Arial'
    });
  });

  // SLIDE 4: Image + Content Split (Professional 50/50)
  const slide4 = pptx.addSlide();
  slide4.background = { color: bg };
  
  // Left: Image placeholder
  slide4.addShape('rect', {
    x: 0.5, y: 0.5, w: 4.5, h: 4.625,
    fill: { color: cardBg }
  });
  
  slide4.addText('📸', {
    x: 0.5, y: 2.2, w: 4.5, h: 1,
    fontSize: 72, align: 'center', valign: 'middle'
  });
  
  // Right: Content
  slide4.addText('Real Results', {
    x: 5.3, y: 1.0, w: 4.2, h: 0.6,
    fontSize: 32, bold: true, color: primary,
    fontFace: 'Arial'
  });
  
  const bullets = [
    '98% customer satisfaction rating',
    '10x faster than traditional methods',
    'Trusted by Fortune 500 companies',
    'Available in 40+ countries worldwide'
  ];
  
  bullets.forEach((bullet, i) => {
    // Bullet point
    slide4.addShape('rect', {
      x: 5.3, y: 2.0 + i * 0.6, w: 0.08, h: 0.08,
      fill: { color: accent }
    });
    
    slide4.addText(bullet, {
      x: 5.5, y: 1.95 + i * 0.6, w: 4.0, h: 0.4,
      fontSize: 14, color: primary,
      fontFace: 'Arial'
    });
  });

  // SLIDE 5: Thank You / CTA (Simple, elegant)
  const slide5 = pptx.addSlide();
  slide5.background = { color: bg };
  
  // Large "Thank You"
  slide5.addText('Thank You', {
    x: 1.5, y: 2.0, w: 7, h: 1.0,
    fontSize: 54, bold: true, color: primary, align: 'center',
    fontFace: 'Arial'
  });
  
  // Accent line
  slide5.addShape('rect', {
    x: 3.5, y: 3.2, w: 3, h: 0.05,
    fill: { color: accent }
  });
  
  // Contact/CTA
  slide5.addText('hello@company.com  •  www.company.com', {
    x: 2, y: 3.6, w: 6, h: 0.4,
    fontSize: 16, color: secondary, align: 'center',
    fontFace: 'Arial'
  });

  const buffer = await pptx.write({ outputType: 'nodebuffer' });
  return buffer;
}

// ═════════════════════════════════════════════════════════
// TEMPLATE 2: MODERN DARK (Netflix/Spotify style)
// ═══════════════════════════════════════════════════════════
async function generateTemplate2() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = 'Template 2 - Modern Dark';

  const bg = '141414';
  const accent = 'E50914';
  const primary = 'FFFFFF';
  const secondary = 'B3B3B3';
  const cardBg = '1F1F1F';

  // SLIDE 1: Hero Title (Bold, simple)
  const slide1 = pptx.addSlide();
  slide1.background = { color: bg };
  
  // Small label/category
  slide1.addText('INNOVATION', {
    x: 0.5, y: 1.0, w: 9, h: 0.3,
    fontSize: 12, bold: true, color: accent, letterSpacing: 3,
    fontFace: 'Arial'
  });
  
  // Big title
  slide1.addText('The Next Chapter', {
    x: 0.5, y: 1.5, w: 9, h: 1.2,
    fontSize: 72, bold: true, color: primary,
    fontFace: 'Arial'
  });
  
  // Subtitle
  slide1.addText('Discover what\'s possible when technology meets creativity', {
    x: 0.5, y: 3.0, w: 6, h: 0.6,
    fontSize: 18, color: secondary,
    fontFace: 'Arial'
  });
  
  // Accent bar (bottom left corner detail)
  slide1.addShape('rect', {
    x: 0.5, y: 4.8, w: 0.15, h: 0.5,
    fill: { color: accent }
  });

  // SLIDE 2: Stats Grid (3 big numbers)
  const slide2 = pptx.addSlide();
  slide2.background = { color: bg };
  
  slide2.addText('By The Numbers', {
    x: 0.5, y: 0.6, w: 9, h: 0.5,
    fontSize: 28, bold: true, color: primary,
    fontFace: 'Arial'
  });
  
  const stats = [
    { num: '500M+', label: 'Global Users' },
    { num: '99.9%', label: 'Uptime' },
    { num: '#1', label: 'Industry Leader' }
  ];
  
  stats.forEach((stat, i) => {
    const x = 0.5 + i * 3.2;
    
    // Card
    slide2.addShape('rect', {
      x, y: 1.8, w: 2.9, h: 2.8,
      fill: { color: cardBg }
    });
    
    // Accent top border
    slide2.addShape('rect', {
      x, y: 1.8, w: 2.9, h: 0.08,
      fill: { color: accent }
    });
    
    // Number
    slide2.addText(stat.num, {
      x: x + 0.2, y: 2.5, w: 2.5, h: 0.8,
      fontSize: 48, bold: true, color: primary, align: 'center',
      fontFace: 'Arial'
    });
    
    // Label
    slide2.addText(stat.label, {
      x: x + 0.2, y: 3.5, w: 2.5, h: 0.5,
      fontSize: 16, color: secondary, align: 'center',
      fontFace: 'Arial'
    });
  });

  // SLIDE 3: Two Column (Problem/Solution)
  const slide3 = pptx.addSlide();
  slide3.background = { color: bg };
  
  slide3.addText('Why It Matters', {
    x: 0.5, y: 0.6, w: 9, h: 0.5,
    fontSize: 28, bold: true, color: primary,
    fontFace: 'Arial'
  });
  
  // Left column
  slide3.addText('THE CHALLENGE', {
    x: 0.5, y: 1.5, w: 4.2, h: 0.3,
    fontSize: 11, bold: true, color: accent, letterSpacing: 2,
    fontFace: 'Arial'
  });
  
  slide3.addText('Traditional approaches are slow, expensive, and difficult to scale in today\'s fast-paced environment.', {
    x: 0.5, y: 2.0, w: 4.2, h: 1.5,
    fontSize: 16, color: secondary, wrap: true,
    fontFace: 'Arial'
  });
  
  // Vertical divider
  slide3.addShape('rect', {
    x: 4.85, y: 1.5, w: 0.02, h: 3.0,
    fill: { color: '333333' }
  });
  
  // Right column
  slide3.addText('OUR SOLUTION', {
    x: 5.3, y: 1.5, w: 4.2, h: 0.3,
    fontSize: 11, bold: true, color: accent, letterSpacing: 2,
    fontFace: 'Arial'
  });
  
  slide3.addText('We deliver a modern, automated platform that saves time, reduces costs, and scales effortlessly.', {
    x: 5.3, y: 2.0, w: 4.2, h: 1.5,
    fontSize: 16, color: primary, wrap: true,
    fontFace: 'Arial'
  });

  // SLIDE 4: Bullet List (Clean, simple)
  const slide4 = pptx.addSlide();
  slide4.background = { color: bg };
  
  slide4.addText('What You Get', {
    x: 0.5, y: 0.6, w: 9, h: 0.5,
    fontSize: 28, bold: true, color: primary,
    fontFace: 'Arial'
  });
  
  const features = [
    'Enterprise-grade security and compliance',
    'Real-time collaboration across teams',
    'AI-powered automation and insights',
    'Seamless integration with existing tools',
    '24/7 dedicated support'
  ];
  
  features.forEach((feature, i) => {
    // Checkmark
    slide4.addText('✓', {
      x: 1.0, y: 1.8 + i * 0.6, w: 0.4, h: 0.4,
      fontSize: 20, bold: true, color: accent,
      fontFace: 'Arial'
    });
    
    // Feature text
    slide4.addText(feature, {
      x: 1.6, y: 1.8 + i * 0.6, w: 7.4, h: 0.4,
      fontSize: 18, color: primary,
      fontFace: 'Arial'
    });
  });

  // SLIDE 5: Call to Action (Strong CTA)
  const slide5 = pptx.addSlide();
  slide5.background = { color: bg };
  
  slide5.addText('Ready to Get Started?', {
    x: 1, y: 2.0, w: 8, h: 0.8,
    fontSize: 48, bold: true, color: primary, align: 'center',
    fontFace: 'Arial'
  });
  
  // CTA button visual
  slide5.addShape('roundRect', {
    x: 3.5, y: 3.2, w: 3, h: 0.7,
    fill: { color: accent },
    rectRadius: 0.1
  });
  
  slide5.addText('Start Free Trial', {
    x: 3.5, y: 3.2, w: 3, h: 0.7,
    fontSize: 20, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle',
    fontFace: 'Arial'
  });
  
  slide5.addText('No credit card required  •  14-day free trial', {
    x: 2, y: 4.2, w: 6, h: 0.4,
    fontSize: 14, color: secondary, align: 'center',
    fontFace: 'Arial'
  });

  const buffer = await pptx.write({ outputType: 'nodebuffer' });
  return buffer;
}

// ═══════════════════════════════════════════════════════════
// TEMPLATE 3: VIBRANT CREATIVE (Startup/Canva style)
// ═══════════════════════════════════════════════════════════
async function generateTemplate3() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = 'Template 3 - Vibrant Creative';

  const bg = 'F8FAFC';
  const accent = 'FF6B6B';
  const accent2 = '4ECDC4';
  const accent3 = 'FFE66D';
  const primary = '2D3748';
  const secondary = '718096';

  // SLIDE 1: Fun Title Slide
  const slide1 = pptx.addSlide();
  slide1.background = { color: bg };
  
  // Colorful accent shapes (background decoration)
  slide1.addShape('ellipse', {
    x: -0.5, y: -0.5, w: 2, h: 2,
    fill: { color: accent2, transparency: 80 }
  });
  
  slide1.addShape('ellipse', {
    x: 8.5, y: 4, w: 1.8, h: 1.8,
    fill: { color: accent, transparency: 80 }
  });
  
  // Title
  slide1.addText('Let\'s Build Something Amazing', {
    x: 1, y: 1.8, w: 8, h: 1.2,
    fontSize: 54, bold: true, color: primary, align: 'center',
    fontFace: 'Arial'
  });
  
  // Colorful underline
  slide1.addShape('rect', {
    x: 3, y: 3.15, w: 1.5, h: 0.12,
    fill: { color: accent }
  });
  slide1.addShape('rect', {
    x: 4.6, y: 3.15, w: 1.5, h: 0.12,
    fill: { color: accent2 }
  });
  slide1.addShape('rect', {
    x: 5.5, y: 3.15, w: 1.5, h: 0.12,
    fill: { color: accent3 }
  });
  
  // Subtitle
  slide1.addText('Your journey to success starts here', {
    x: 2, y: 3.8, w: 6, h: 0.5,
    fontSize: 20, color: secondary, align: 'center',
    fontFace: 'Arial'
  });

  // SLIDE 2: Icon + Stats (Colorful cards)
  const slide2 = pptx.addSlide();
  slide2.background = { color: bg };
  
  slide2.addText('Our Impact', {
    x: 0.5, y: 0.6, w: 9, h: 0.5,
    fontSize: 32, bold: true, color: primary,
    fontFace: 'Arial'
  });
  
  const impacts = [
    { icon: '🚀', num: '200+', label: 'Projects Launched', color: accent },
    { icon: '😊', num: '50K', label: 'Happy Customers', color: accent2 },
    { icon: '⭐', num: '4.9', label: 'Average Rating', color: accent3 }
  ];
  
  impacts.forEach((item, i) => {
    const x = 0.5 + i * 3.2;
    
    // Colorful card
    slide2.addShape('roundRect', {
      x, y: 1.8, w: 2.9, h: 3.0,
      fill: { color: 'FFFFFF' },
      line: { color: item.color, width: 3 },
      rectRadius: 0.2
    });
    
    // Icon
    slide2.addText(item.icon, {
      x: x + 0.3, y: 2.2, w: 2.3, h: 0.6,
      fontSize: 42, align: 'center'
    });
    
    // Number
    slide2.addText(item.num, {
      x: x + 0.3, y: 3.0, w: 2.3, h: 0.5,
      fontSize: 36, bold: true, color: item.color, align: 'center',
      fontFace: 'Arial'
    });
    
    // Label
    slide2.addText(item.label, {
      x: x + 0.3, y: 3.7, w: 2.3, h: 0.6,
      fontSize: 14, color: secondary, align: 'center',
      fontFace: 'Arial'
    });
  });

  // SLIDE 3: Timeline/Process (Colorful steps)
  const slide3 = pptx.addSlide();
  slide3.background = { color: bg };
  
  slide3.addText('How It Works', {
    x: 0.5, y: 0.6, w: 9, h: 0.5,
    fontSize: 32, bold: true, color: primary,
    fontFace: 'Arial'
  });
  
  const steps = [
    { num: '1', title: 'Sign Up', desc: 'Create your free account in seconds', color: accent },
    { num: '2', title: 'Customize', desc: 'Tailor everything to your needs', color: accent2 },
    { num: '3', title: 'Launch', desc: 'Go live and start growing', color: accent3 }
  ];
  
  steps.forEach((step, i) => {
    const x = 0.5 + i * 3.2;
    
    // Number circle
    slide3.addShape('ellipse', {
      x: x + 1.0, y: 1.8, w: 0.9, h: 0.9,
      fill: { color: step.color }
    });
    
    slide3.addText(step.num, {
      x: x + 1.0, y: 1.8, w: 0.9, h: 0.9,
      fontSize: 32, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle',
      fontFace: 'Arial'
    });
    
    // Title
    slide3.addText(step.title, {
      x, y: 3.0, w: 2.9, h: 0.4,
      fontSize: 20, bold: true, color: primary, align: 'center',
      fontFace: 'Arial'
    });
    
    // Description
    slide3.addText(step.desc, {
      x, y: 3.5, w: 2.9, h: 0.8,
      fontSize: 14, color: secondary, align: 'center', wrap: true,
      fontFace: 'Arial'
    });
    
    // Arrow (except last)
    if (i < 2) {
      slide3.addText('→', {
        x: x + 3.0, y: 2.1, w: 0.5, h: 0.5,
        fontSize: 28, color: step.color, align: 'center'
      });
    }
  });

  // SLIDE 4: Benefits Grid (2x2)
  const slide4 = pptx.addSlide();
  slide4.background = { color: bg };
  
  slide4.addText('Why Choose Us?', {
    x: 0.5, y: 0.6, w: 9, h: 0.5,
    fontSize: 32, bold: true, color: primary,
    fontFace: 'Arial'
  });
  
  const benefits = [
    { icon: '💡', title: 'Innovative', color: accent },
    { icon: '⚡', title: 'Fast', color: accent2 },
    { icon: '🎯', title: 'Accurate', color: accent3 },
    { icon: '🔒', title: 'Secure', color: accent }
  ];
  
  benefits.forEach((benefit, i) => {
    const x = 0.8 + (i % 2) * 4.5;
    const y = 1.8 + Math.floor(i / 2) * 2.0;
    
    // Card
    slide4.addShape('roundRect', {
      x, y, w: 4.0, h: 1.6,
      fill: { color: 'FFFFFF' },
      line: { type: 'none' },
      rectRadius: 0.15
    });
    
    // Colored left edge
    slide4.addShape('rect', {
      x, y, w: 0.15, h: 1.6,
      fill: { color: benefit.color }
    });
    
    // Icon
    slide4.addText(benefit.icon, {
      x: x + 0.4, y: y + 0.4, w: 0.8, h: 0.8,
      fontSize: 32
    });
    
    // Title
    slide4.addText(benefit.title, {
      x: x + 1.4, y: y + 0.5, w: 2.2, h: 0.6,
      fontSize: 22, bold: true, color: primary,
      fontFace: 'Arial'
    });
  });

  // SLIDE 5: Contact/CTA (Friendly & approachable)
  const slide5 = pptx.addSlide();
  slide5.background = { color: bg };
  
  // Decorative circles
  slide5.addShape('ellipse', {
    x: 1, y: 1, w: 1.5, h: 1.5,
    fill: { color: accent2, transparency: 85 }
  });
  
  slide5.addShape('ellipse', {
    x: 7.5, y: 3, w: 1.2, h: 1.2,
    fill: { color: accent, transparency: 85 }
  });
  
  // Main message
  slide5.addText('Let\'s Chat!', {
    x: 2, y: 2.0, w: 6, h: 0.8,
    fontSize: 48, bold: true, color: primary, align: 'center',
    fontFace: 'Arial'
  });
  
  slide5.addText('We\'d love to hear from you', {
    x: 2, y: 3.0, w: 6, h: 0.4,
    fontSize: 20, color: secondary, align: 'center',
    fontFace: 'Arial'
  });
  
  // Colorful CTA button
  slide5.addShape('roundRect', {
    x: 3.5, y: 3.8, w: 3, h: 0.6,
    fill: { color: accent },
    rectRadius: 0.3
  });
  
  slide5.addText('Get In Touch', {
    x: 3.5, y: 3.8, w: 3, h: 0.6,
    fontSize: 18, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle',
    fontFace: 'Arial'
  });

  const buffer = await pptx.write({ outputType: 'nodebuffer' });
  return buffer;
}

// Main execution
async function generatePreviews() {
  console.log('═══════════════════════════════════════════════════════════');
  console.log('🎨 GENERATING 3 PROFESSIONAL TEMPLATES');
  console.log('═══════════════════════════════════════════════════════════\n');

  const outputDir = path.join(__dirname, 'output');
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  console.log('[1/3] Template 1: Clean Corporate (Apple/Microsoft)...');
  const buffer1 = await generateTemplate1();
  fs.writeFileSync(path.join(outputDir, 'TEMPLATE_1_CLEAN_CORPORATE.pptx'), buffer1);
  console.log('     ✅ Saved: TEMPLATE_1_CLEAN_CORPORATE.pptx\n');

  console.log('[2/3] Template 2: Modern Dark (Netflix/Spotify)...');
  const buffer2 = await generateTemplate2();
  fs.writeFileSync(path.join(outputDir, 'TEMPLATE_2_MODERN_DARK.pptx'), buffer2);
  console.log('     ✅ Saved: TEMPLATE_2_MODERN_DARK.pptx\n');

  console.log('[3/3] Template 3: Vibrant Creative (Startup/Canva)...');
  const buffer3 = await generateTemplate3();
  fs.writeFileSync(path.join(outputDir, 'TEMPLATE_3_VIBRANT_CREATIVE.pptx'), buffer3);
  console.log('     ✅ Saved: TEMPLATE_3_VIBRANT_CREATIVE.pptx\n');

  console.log('═══════════════════════════════════════════════════════════');
  console.log('✅ Complete! 3 professional templates inspired by real brands');
  console.log(`📁 Location: ${outputDir}`);
  console.log('═══════════════════════════════════════════════════════════');
}

generatePreviews().catch(err => {
  console.error('❌ Error:', err);
  process.exit(1);
});
