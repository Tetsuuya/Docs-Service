# Gemini Architecture Variety Test

This script tests if Gemini generates **truly varied** presentation architectures across multiple prompts.

## What It Tests

- ✅ Slide count variety (not always 6 slides)
- ✅ Image count variety (not always 3 images)
- ✅ Layout sequence uniqueness (no repeated patterns like hero → split → cards → stat → timeline)
- ✅ First slide layout variation (not always hero)
- ✅ Image placement patterns (not always slides 1, 2, 5)
- ✅ Overall layout distribution

## How to Run

**Run from project root** (so .env is loaded):

```bash
node test/gemini-generation-json/test-variety.js
```

## Output

The script will:

1. **Generate 8 presentations** with different prompts:
   - Music
   - Space exploration
   - Artificial intelligence
   - Healthy cooking
   - Climate change
   - Entrepreneurship
   - Yoga and meditation
   - Cybersecurity

2. **Save JSON files** to `output/` directory:
   - Individual JSON for each generation
   - `_ANALYSIS_REPORT.json` with full analysis

3. **Print analysis** to console:
   - Slide count range and variety
   - Image count range and variety
   - Layout sequence uniqueness
   - First slide layout distribution
   - Image placement patterns
   - Overall layout usage statistics

## Success Criteria

✅ **GOOD** (High Variety):
- All layout sequences are unique
- Slide counts vary (4-11 range)
- Image counts vary (2-5 range)
- First slides use different layouts
- No "1,2,5" image placement pattern

❌ **BAD** (Low Variety):
- Duplicate layout sequences
- Always 6 slides
- Always 3 images
- Always hero first slide
- Always images on slides 1, 2, 5

## Example Output

```
📊 PATTERN ANALYSIS
═══════════════════════════════════════════════════════════

📏 SLIDE COUNTS:
   Range: 5 - 9
   Unique values: 5, 6, 7, 8, 9
   Variety score: 5/8 (62%)

🖼️  IMAGE COUNTS:
   Range: 2 - 5
   Unique values: 2, 3, 4, 5
   Variety score: 4/8 (50%)

🎨 LAYOUT SEQUENCES:
   Unique sequences: 8/8
   ✅ All sequences are unique!

🚀 FIRST SLIDE LAYOUTS:
   minimal: 2x (25%)
   stat: 2x (25%)
   quote: 2x (25%)
   split: 1x (12%)
   cards: 1x (12%)

📍 IMAGE PLACEMENT PATTERNS:
   Unique patterns: 8/8
   ✅ No "1,2,5" pattern detected!
```

## What to Look For

If you see:
- ⚠️  DUPLICATES FOUND - AI is repeating patterns
- ⚠️  Found "1,2,5" pattern - AI using old hardcoded pattern
- Low variety scores (<40%) - Need to increase randomization

Then the anti-pattern rules need to be stronger.
