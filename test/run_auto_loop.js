import fs from 'fs';
import path from 'path';
import { execSync } from 'child_process';
import { buildPptxFile } from '../src/services/pptx/pptxService.js';

async function runAutoLoop() {
  const topics = [
    'Artificial Intelligence and Machine Learning',
    'Cancer Research and Clinical Diagnostics',
    'Global Economic Outlook and Financial Markets'
  ];

  const templatePath = path.join(process.cwd(), 'AGENCE DÉCLIC 3.pptx');
  const artifactsDir = `C:\\Users\\Rhenel Jhon Sajol\\.gemini\\antigravity-ide\\brain\\40d82de6-ac2c-4b88-9768-f8948dee2ffc`;

  for (let i = 0; i < topics.length; i++) {
    const topic = topics[i];
    console.log(`\n====================================================`);
    console.log(`🚀 ITERATION ${i + 1}/${topics.length}: Generating & Auditing "${topic}"...`);
    console.log(`====================================================`);

    const buffer = await buildPptxFile(topic, 'fill', templatePath);
    const pptxPath = path.join(process.cwd(), 'temp', `test_deck_${i + 1}.pptx`);
    fs.mkdirSync(path.join(process.cwd(), 'temp'), { recursive: true });
    fs.writeFileSync(pptxPath, buffer);

    console.log(`\n📸 Rendering slides to JPG and auditing structure...`);
    const cmd = `python test/auto_evaluate_presentation.py "${pptxPath}" "${artifactsDir}"`;
    const output = execSync(cmd, { encoding: 'utf-8' });
    console.log(output);
  }

  console.log(`\n✨ AUTOMATED EVALUATION LOOP COMPLETE FOR ALL ${topics.length} TOPICS!`);
}

runAutoLoop();
