import fs from 'fs';
import path from 'path';
import { execSync } from 'child_process';
import pptxgen from 'pptxgenjs';
import { logger } from '../../../utils/logger.js';

/**
 * Phase 4 — Template Fill Engine & Slide Assembly
 */
export const buildFillPptx = async (fillPlan, templateBlueprint, masterFilePath = null, autoImages = {}) => {
  const title = fillPlan.presentationTitle || 'Untitled Presentation';
  logger.info(`Phase 4: Assembling filled presentation — "${title}"`);

  let isPptxMaster = false;
  if (masterFilePath && fs.existsSync(masterFilePath)) {
    try {
      const ext = path.extname(masterFilePath).toLowerCase();
      if (ext === '.pptx') {
        isPptxMaster = true;
      } else {
        const buf = fs.readFileSync(masterFilePath);
        if (buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4B) {
          isPptxMaster = true;
        }
      }
    } catch (_) {}
  }

  // ─────────────────────────────────────────────────────────────────────────
  // PRIMARY PATH — Native PPTX Fill (Text + Image Placeholders)
  // ─────────────────────────────────────────────────────────────────────────
  if (isPptxMaster) {
    const tempDir = path.join(process.cwd(), 'temp');
    fs.mkdirSync(tempDir, { recursive: true });

    const timestamp = Date.now();
    const planJsonPath = path.join(tempDir, `fill_plan_${timestamp}.json`);
    const imgJsonPath = path.join(tempDir, `fill_images_${timestamp}.json`);
    const outputPptxPath = path.join(tempDir, `filled_${timestamp}.pptx`);
    const helperScript = path.join(
      process.cwd(), 'src', 'services', 'pptx', 'fill', 'directPptxReplacer.py'
    );

    const cleanup = () => {
      [planJsonPath, imgJsonPath, outputPptxPath].forEach(p => {
        try { if (fs.existsSync(p)) fs.unlinkSync(p); } catch (_) {}
      });
    };

    try {
      logger.info(`[fillPptxBuilder] Sanitized plan sent to Python: ${JSON.stringify(fillPlan.selectedSlides.map(s => ({ idx: s.slideIndex, cat: s.layoutCategory })), null, 2)}`);
      fs.writeFileSync(planJsonPath, JSON.stringify(fillPlan, null, 2), 'utf8');
      fs.writeFileSync(imgJsonPath, JSON.stringify(autoImages, null, 2), 'utf8');

      const pythonBin = getPythonBinary();
      if (!pythonBin) {
        throw new Error('Python interpreter not found in PATH');
      }

      const cmd = `"${pythonBin}" "${helperScript}" "${masterFilePath}" "${planJsonPath}" "${outputPptxPath}" "${imgJsonPath}"`;
      logger.info(`  Executing: ${pythonBin} directPptxReplacer.py ...`);

      const stdout = execSync(cmd, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'pipe'] });
      if (stdout) logger.info(`  Python output: ${stdout.trim()}`);

      if (!fs.existsSync(outputPptxPath)) {
        throw new Error('Python replacer ran but output file was not created');
      }

      const buffer = fs.readFileSync(outputPptxPath);
      cleanup();

      logger.info(`✅ Phase 4 Complete (native PPTX fill) — ${Math.round(buffer.length / 1024)} KB`);
      return buffer;

    } catch (err) {
      cleanup();
      logger.warn(`Primary PPTX fill path failed: ${err.message}`);
      logger.warn('Falling back to pptxgenjs builder...');
    }
  }

  logger.info('Phase 4 Fallback: Building presentation with pptxgenjs...');
  return buildFallbackPptx(fillPlan, templateBlueprint);
};

function getPythonBinary() {
  for (const bin of ['python3', 'python', 'py']) {
    try {
      execSync(`"${bin}" --version`, { stdio: 'ignore' });
      return bin;
    } catch (_) {}
  }
  return null;
}

async function buildFallbackPptx(fillPlan, templateBlueprint) {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_16x9';
  pptx.title = fillPlan.presentationTitle || 'Filled Presentation';
  pptx.author = 'Docs-Service AI';

  const brand = templateBlueprint?.brandTheme || {};
  const primary   = (brand.primaryColor   || '071E3D').replace('#', '');
  const secondary = (brand.secondaryColor || '1E293B').replace('#', '');
  const accent    = (brand.accentColor    || '38B6FF').replace('#', '');
  const bg        = (brand.backgroundColor || 'FFFFFF').replace('#', '');
  const font      = brand.fontFamily || 'Calibri';

  const selectedSlides = fillPlan.selectedSlides || [];

  for (let i = 0; i < selectedSlides.length; i++) {
    const slideData = selectedSlides[i];
    const slide = pptx.addSlide();
    const fill = slideData.fillContent || {};

    slide.background = { color: bg.toUpperCase() !== 'FFFFFF' ? bg : 'F8FAFC' };

    slide.addShape(pptx.shapes.RECTANGLE, {
      x: 0, y: 0, w: '100%', h: 0.12,
      fill: { color: accent }, line: { type: 'none' }
    });

    const fillValues = Object.values(fill).filter(v => typeof v === 'string' && v.trim());

    if (i === 0) {
      slide.addText(fillPlan.presentationTitle || 'Presentation Title', {
        x: 0.8, y: 1.8, w: 11.3, h: 2.2,
        fontSize: 38, bold: true, color: primary, fontFace: font
      });
    } else {
      slide.addText(fillValues[0] || `Section ${i + 1}`, {
        x: 0.8, y: 0.5, w: 11.3, h: 0.8,
        fontSize: 24, bold: true, color: primary, fontFace: font
      });
      if (fillValues.length > 1) {
        slide.addText(fillValues.slice(1).join('\n\n'), {
          x: 0.8, y: 1.5, w: 11.3, h: 4.8,
          fontSize: 14, color: secondary, fontFace: font
        });
      }
    }
  }

  return await pptx.write({ outputType: 'nodebuffer' });
}
