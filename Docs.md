# Docs-Service Documentation

## Overview
**Docs-Service** is an enterprise AI-powered document generation microservice built with **Node.js, Express, Google Gemini API, and docx Engine**.

It transforms plain text prompts and uploaded context files into styled, multi-page Microsoft Word (`.docx`) documents with executive cover pages, running headers, dynamic page numbers, domain-aware color themes, structured data tables, and callout boxes.

---

## System Architecture & Data Flow

```
[ User Web UI / Mobile / API Client ]
           │ (POST /api/generate with FormData)
           ▼
[ Express Router: src/routes/documentRoutes.js ] (Multer File Middleware)
           │
           ▼
[ Controller: src/controllers/documentController.js ]
           │
           ▼
[ AI Planner: src/services/geminiService.js ]
           │ 1. Multimodal context & badge OCR extraction (Images, PDFs, TXT, CSV)
           │ 2. Step 1: Generate Theme, Title, Executive Overview, TOC & Page Outline
           │ 3. Step 2: Concurrently generate calibrated page items
           │ 4. Step 3: Return Page-Structured JSON AST (with unique doc_id & prompt)
           ▼
[ JSON History Backup: src/utils/historyStorage.js ]
           │ Saves payload to temp/history/doc_<timestamp>_<hash>.json
           ▼
[ Word Compiler Engine: src/services/docxService.js ]
           │ Renders Cover Page (Banner, Image, Overview, TOC, Metadata)
           │ Renders Page-Structured Items with PageBreaks per page
           ▼
[ Binary Stream Download (.docx) ]
           │ Dynamic filename header: Content-Disposition: attachment; filename="title.docx"
```

---

## Core Implemented Features

### 1. Multimodal Input Engine
* Supports uploading context files alongside text prompts: images (`image/*`), `.pdf`, `.txt`, `.csv`, `.json`, `.md`.
* Performs OCR on uploaded image badges to preserve exact branding text.

### 2. Executive Cover Page (Page 1)
* **🏷️ Classification Tag Badge**: Dynamic domain tag (e.g. `🏷️ ENTERPRISE PLATFORM GUIDE`).
* **Primary Title Banner**: High-contrast theme box with title & subtitle.
* **🖼️ Embedded Cover Image**: Centered badge or logo screenshot (`ImageRun`).
* **📋 Executive Overview Box**: Styled summary card introducing document scope.
* **📚 Table of Contents Box**: Formatted outline mapping section titles to page numbers.
* **Publication Metadata Card**: Author and generated date details.
* **Cover PageBreak**: Forces major section content to start cleanly on Page 2.

### 3. Page-Structured JSON AST & Calibrated Compiler
* Gemini structures content into a calibrated `pages` JSON array.
* **Minimum Content Baseline**: 2-3 thorough paragraphs + 3-5 bulleted items + 1 callout box per page object.
* `docxService.js` renders each page item followed by a `PageBreak()`, guaranteeing **100% full pages with ZERO blank overflow pages**.

### 4. Dynamic Filename Engine
* Backend sanitizes `contentData.title` to compute clean HTTP headers:
  `Content-Disposition: attachment; filename="youtube_ecosystem_and_design_theme.docx"`
* CORS exposes `Content-Disposition` so frontend JS dynamically sets `a.download` to match document title.

### 5. Azure Container Apps & Scale-to-Zero ($0 Idle Cost)
* Containerized with `Dockerfile` (`node:20-alpine`).
* Deployed on **Azure Container Apps (ACA)** in `web-scrapper-rg` North Europe.
* Configured with **Scale-to-Zero** (`minReplicas: 0`) so server scales down to **0 instances ($0 cost)** when idle.

### 6. Automated CI/CD & Version Tracking
* Automated GitHub Actions pipeline ([`.github/workflows/deploy.yml`](.github/workflows/deploy.yml)) triggers build & deployment to Azure on `git push main`.
* **Health Endpoint**: `GET /health` or `GET /api/version` returns JSON status and live GitHub Commit SHA:
  ```json
  {
    "status": "ok",
    "service": "Docs-Service",
    "version": "481552d0a1b2...",
    "timestamp": "2026-07-23T14:05:00.000Z"
  }
  ```

---

## API Reference

### 1. Unified Generation Endpoint
`POST /api/generate`

**Form Data / Body Parameters:**
* `prompt` (string, required): e.g. *"create a 5 pages docxs about youtube, I want youtube design theme too"*
* `format` (string, optional): `"docx"` (default)
* `mode` (string, optional): `"scratch"` (default)
* `contextFile` (file, optional): Multimodal image or document file upload

**Headers Returned:**
* `Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document`
* `Content-Disposition: attachment; filename="youtube_ecosystem_and_design_theme.docx"`
* `X-Document-Id: doc_<timestamp>_<hash>`
* `X-App-Version: <github_commit_sha>`

### 2. Health & Version Endpoint
`GET /health` or `GET /api/version`

**Response:**
```json
{
  "status": "ok",
  "service": "Docs-Service",
  "version": "v1.0.1-release",
  "timestamp": "2026-07-23T14:05:00.000Z"
}
```
