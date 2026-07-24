# Implementation Roadmap: AI Document Generation Service (PoC)

## Executive Summary
This document breaks down the step-by-step execution plan for building the **AI Document Generation Service Proof-of-Concept (PoC)** using **Gemini 2.0 Flash Lite**, a **Node.js Express Backend**, the open-source **`docx`** library, and a **Mini HTML Frontend**.

---

## Phase 1: Environment & Project Setup - COMPLETED ✅
- **Goal:** Initialize project workspace, configure dependencies, and establish folder layout.

---

## Phase 2: Backend API & Gemini Flash Lite Integration - COMPLETED ✅
- **Goal:** Build backend endpoint and integrate Gemini 2.0 Flash Lite API with structured JSON output.

---

## Phase 3A: Word Document Engine (`docxService.js`) - COMPLETED ✅
- **Goal:** Transform Gemini's JSON structure into a styled `.docx` Word file with Cover Page, Table of Contents, Executive Overview, and Page-Structured Items.

---

## Phase 4: Multimodal HTML/JS Frontend UI (`public/index.html`) - COMPLETED ✅
- **Goal:** Create a clean, responsive interface supporting text prompts, multimodal file uploads, and dynamic filename downloads.

---

## Phase 5: Cloud Deployment & CI/CD Pipeline - COMPLETED ✅
- **Goal:** Deploy containerized microservice to Azure Container Apps with Scale-to-Zero ($0 idle cost) and GitHub Actions automation.
- **Deliverables:**
  1. Dockerized Node.js environment (`Dockerfile` with `node:20-alpine`).
  2. Pushed image to Azure Container Registry (`acrdocsservice.azurecr.io`).
  3. Deployed to **Azure Container Apps** in `web-scrapper-rg` North Europe with **Scale-to-Zero (`minReplicas: 0`)**.
  4. Configured automated GitHub Actions workflow (`.github/workflows/deploy.yml`) on `git push main`.
  5. Implemented `/health` & `/api/version` endpoints returning live commit SHA.

---

## Phase 6: PowerPoint, Excel & PDF Engines (Upcoming Roadmap)
- `POST /api/generate/pptx` - PowerPoint Generator Engine (`PptxGenJS`)
- `POST /api/generate/xlsx` - Excel Generator Engine (`ExcelJS`)
- `POST /api/generate/pdf` - High-Fidelity PDF Engine (`Gotenberg` / `LibreOffice`)

