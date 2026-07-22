# Implementation Roadmap: AI Document Generation Service (PoC)

## Executive Summary
This document breaks down the step-by-step execution plan for building the **AI Document Generation Service Proof-of-Concept (PoC)** using **Gemini 2.0 Flash Lite**, a **Node.js Express Backend**, the open-source **`docx`** library, and a **Mini HTML Frontend**.

---

## Phase 1: Environment & Project Setup
- **Goal:** Initialize project workspace, configure dependencies, and establish folder layout.
- **Deliverables:**
  1. Initialize Node.js environment (`package.json`).
  2. Install core packages: `express`, `@google/genai`, `docx`, `cors`, `dotenv`.
  3. Configure `.env` file for secure `GEMINI_API_KEY` management.
  4. Establish project folder structure tree:

```
Docs-Service/
├── public/
│   └── index.html               # Mini HTML/JS Web UI
├── src/
│   ├── config/
│   │   └── env.js               # Environment & Gemini API Key config
│   ├── controllers/
│   │   └── documentController.js # Handlers for all document format routes
│   ├── services/
│   │   ├── geminiService.js     # Gemini 2.0 Flash Lite AI Planner (JSON Schema)
│   │   ├── docxService.js       # Word (.docx) Generator Engine
│   │   ├── pptxService.js       # PowerPoint (.pptx) Generator Engine
│   │   ├── xlsxService.js       # Excel (.xlsx) Generator Engine
│   │   └── pdfService.js        # PDF (.pdf) Conversion Engine
│   ├── routes/
│   │   └── documentRoutes.js    # Express router for all document formats
│   ├── utils/
│   │   └── logger.js            # Enterprise structured console logger & middleware
│   └── app.js                   # Express application & middleware setup
├── temp/                        # Temporary working folder (for template uploads & PDF staging)
├── server.js                    # Server entry point (starts server on port 3000)
├── .env                         # Environment variables (GEMINI_API_KEY)
├── .env.example                 # Template for environment variables
├── .gitignore                   # Git ignore rule file (ignores node_modules & .env)
├── package.json                 # Dependencies & NPM scripts
└── project_phases.md            # Implementation Roadmap
```

---

## Phase 2: Backend API & Gemini Flash Lite Integration (`server.js`)
- **Goal:** Build backend endpoint and integrate Gemini 2.0 Flash Lite API with structured JSON output.
- **Deliverables:**
  1. Create Express HTTP server listening on port `3000`.
  2. Implement `POST /api/generate` endpoint accepting `{ prompt }`.
  3. Connect to Google's `@google/genai` SDK using `gemini-2.0-flash-lite`.
  4. Enforce strict JSON Schema output from Gemini containing:
     - Document Title & Subtitle
     - Array of Sections (Heading 1, Paragraphs)
     - Table Data Object (Headers array & Rows matrix)

---

## Phase 3A: Word Document Engine (`docxService.js`) - Current Focus
- **Goal:** Transform Gemini's JSON structure into a styled `.docx` Word file.
- **Deliverables:**
  1. Implement **Running Header** (e.g., *"DOCS SERVICE | CONFIDENTIAL"*).
  2. Implement **Running Footer** with dynamic Page Numbers (*"Page X of Y"*).
  3. Implement **Styled Typography** (Title in 28pt bold, Heading 1 in Navy Blue `#1E3A8A`).
  4. Implement **Styled Table** (Navy Blue `#1E3A8A` header row, white bold text, cell borders).
  5. Stream binary buffer to client.

---

## Phase 3B: PowerPoint Presentation Engine (`pptxService.js`)
- **Goal:** Build native PowerPoint slide decks from Gemini's JSON layout using `PptxGenJS`.
- **Deliverables:**
  1. Implement Master Slide Layouts (Title slide, 3-card layout, Split image/text).
  2. Implement Brand Color Schemes & Theme Tokens.
  3. Implement Native Editable Shapes, Text Boxes & Charts.

---

## Phase 3C: Excel Spreadsheet Engine (`xlsxService.js`)
- **Goal:** Build structured spreadsheets with working formulas using `ExcelJS`.
- **Deliverables:**
  1. Implement Worksheet styling, grid formatting & bold headers.
  2. Implement Excel Formulas (`=SUM()`, `=AVERAGE()`).
  3. Implement Dynamic Data Tables & Chart generation.

---

## Phase 3D: PDF Conversion Engine (`pdfService.js`)
- **Goal:** Convert generated `.pptx`, `.docx`, and `.xlsx` files into high-fidelity PDFs via `Gotenberg`.
- **Deliverables:**
  1. Connect to Dockerized `Gotenberg` LibreOffice API.
  2. Convert office buffers to PDF stream.
  3. Deliver `.pdf` files to client.

---

## Phase 4: Mini HTML Frontend UI (`public/index.html`)
- **Goal:** Create a clean, user-friendly browser interface for testing.
- **Deliverables:**
  1. Build responsive HTML/CSS interface with prompt input text area.
  2. Add "Generate Document" action button with loading state.
  3. Implement client-side `fetch()` logic to handle backend response.
  4. Implement automatic `.docx` file blob download in browser.

---

## Phase 5: Testing, Validation & Boss Presentation
- **Goal:** Ensure document fidelity and present working prototype to boss.
- **Deliverables:**
  1. Test simple prompts (*"Create a Q3 Project Status Report"*).
  2. Test complex data prompts (*"Create a Cloud Infrastructure Budget Proposal with Cost Tables"*).
  3. Verify `.docx` opens cleanly in Microsoft Word / Google Docs with intact headers, footers, and table colors.
  4. Share working prototype demonstration with management.
