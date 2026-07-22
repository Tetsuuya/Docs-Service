# Docs-Service Documentation

## Overview
Docs-Service is an enterprise AI-powered document generation microservice built with **Node.js, Express, Google Gemini API, and docx Engine**.

It transforms plain text prompts into styled, multi-page Microsoft Word (`.docx`) files with running headers, footers, dynamic page numbers, domain-aware color themes, formatted data tables, and callout boxes.

---

## System Architecture & Data Flow

```
[ User Web UI / Postman ]
           │ (POST /api/generate)
           ▼
[ Express Router: src/routes/documentRoutes.js ]
           │
           ▼
[ Controller: src/controllers/documentController.js ]
           │
           ▼
[ AI Planner: src/services/geminiService.js ]
           │ 1. Step 1: Generate Theme, Title, Table & Section Outline
           │ 2. Step 2: Concurrently generate multi-paragraph section content
           │ 3. Step 3: Return JSON AST (with unique doc_id & prompt)
           ▼
[ JSON History Backup: src/utils/historyStorage.js ]
           │ Saves payload to temp/history/doc_<timestamp>_<hash>.json
           ▼
[ Word Compiler Engine: src/services/docxService.js ]
           │ Applies dynamic colors, headers, footers & PageBreaks per section
           ▼
[ Binary Stream Download (.docx) ]
```

---

## Core Implemented Features

### 1. Gemini 2.0 Flash Lite AI Planner (`geminiService.js`)
* **Model:** Configured with `gemini-flash-lite-latest` (fast, cost-effective at ~$0.005 per 10-page document).
* **Automatic Fallback:** Includes candidate model fallbacks (`gemini-flash-lite-latest` ➔ `gemini-2.0-flash-lite-preview-02-05` ➔ `gemini-1.5-flash-latest`).
* **Typo Correction:** Automatically detects and fixes prompt typos (e.g. *"algorthim"* ➔ *"Algorithm"*).

### 2. Clean Page Breaks per Section (`docxService.js`)
* Inserts a native `PageBreak()` before every major section after Section 1.
* Ensures every major chapter starts at the top of a clean, dedicated new page.

### 3. Dynamic Domain-Aware Theme Colors
* Gemini infers the document domain and generates a matching Hex color palette (`primaryColor`, `secondaryColor`, `accentColor`, `lightBgColor`).
* *Examples:* Finance = Emerald (`#047857`), Tech = Slate (`#0F172A`), Executive = Navy (`#1E3A8A`), Literature = Purple (`#4A154B`).

### 4. JSON AST Storage & Unique Document IDs (`historyStorage.js`)
* Generates a unique Document ID for every request (e.g., `doc_1784733409539_2dfc4b`).
* Saves full JSON AST to `temp/history/*.json`.
* **Micro-Editing Ready:** Allows instant 1ms re-rendering of edited text or colors for **$0 AI cost** (no re-calling Gemini).

### 5. Enterprise Structured Logger (`logger.js`)
* Console logger with color-coded timestamps (`INFO`, `WARN`, `ERROR`).
* Express middleware tracking HTTP method, status codes, and execution speed in milliseconds (`ms`).

---

## API Reference

### 1. Unified Generation Endpoint
`POST /api/generate`
```json
{
  "prompt": "Create a 10 page reviewer about data structures and algorithms",
  "format": "docx",
  "mode": "scratch"
}
```
**Response:** Binary `.docx` file attachment with `X-Document-Id: doc_<timestamp>_<hash>` header.

### 2. Format Specific Endpoints
* `POST /api/generate/docx` - Word Document Generator (In Progress)
* `POST /api/generate/pptx` - PowerPoint Generator (To Be Implemented)
* `POST /api/generate/xlsx` - Excel Generator (To Be Implemented)
* `POST /api/generate/pdf` - PDF Generator (To Be Implemented)
