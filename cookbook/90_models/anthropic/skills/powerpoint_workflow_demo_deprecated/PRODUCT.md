# PRODUCT.md — PowerPoint Workflow Suite

## Overview

The PowerPoint Workflow Suite is an AI-powered pipeline that transforms a simple text prompt into a polished, template-styled `.pptx` presentation. It combines multi-model AI content generation, intelligent image creation, deterministic template assembly, and optional vision-based quality assurance — all orchestrated through Agno's Workflow framework.

The system ships two complementary pipelines for different deck sizes, sharing a single codebase of helper functions, data models, and styling logic.

## Solution Architecture

```
                    ┌─────────────────┐
                    │   YOUR INPUTS   │
                    │  Text Prompt    │
                    │  + Template     │
                    └────────┬────────┘
                             │
         ┌───────────────────┴───────────────────┐
         │                                       │
    ≤ 7 slides                             8-15+ slides
         │                                       │
         ▼                                       ▼
  powerpoint_template_              powerpoint_chunked_
  workflow.py (single-call)         workflow.py (chunked)
         │                                       │
         └───────────────────┬───────────────────┘
                             │
                       ┌─────┴─────┐
                       │ Core Steps│
                       └───────────┘
```

### Modular Design

| Layer                | Responsibility                                                             |
| -------------------- | -------------------------------------------------------------------------- |
| **Core Engine**      | 5-step sequential pipeline: Content → Image Plan → Image Gen → Assembly → QA |
| **Data Layer**       | Pydantic models + dataclasses for structured data flow between steps       |
| **Interface Layer**  | CLI with 15+ flags; operates in Raw (no template) or Template-Assisted mode |
| **Integration Layer**| Anthropic API (Claude), Google Gemini API, NanoBanana image gen, LibreOffice headless |

---

## Key Features

### 1. Multi-Model AI Content Generation
Claude (Anthropic) with the native `pptx` skill writes all slide content — titles, bullets, tables, and charts — from a single text prompt. The chunked workflow adds a storyboard planning phase (via an Opus-level query optimizer) before generation.

### 2. Intelligent Brand-Aware Generation
An autonomous `brand_style_analyzer` agent (Sonnet + web search) detects brand directives in the prompt, optionally researches brand guidelines online, and injects structured `BrandStyleIntent` (colors, tone, fonts, style keywords) into every downstream LLM call.

### 3. Template-Faithful Assembly
A fully deterministic Step 4 builds a comprehensive **assembly knowledge file** — combining user intent, content plan, deep per-layout template analysis, and AI image assets — then maps all content onto the template's native layouts, fonts, colors, and placeholder regions. No LLM in the loop; testable and predictable.

### 4. AI Image Generation
Gemini-based image planning decides which slides benefit from visuals. NanoBanana generates 16:9 PNG images. Images are scaled/centered within the template's content area preserving aspect ratio.

### 5. Vision-Based Quality Assurance (Optional)
Gemini 2.5 Flash renders each slide to PNG (via LibreOffice headless), detects visual defects (text overflow, ghost text, low contrast, overlap), and auto-corrects critical issues using existing deterministic functions.

### 6. Production-Grade Reliability (3-Tier Fallback)
Each chunk goes through a fallback chain: **Tier 1** (Claude PPTX skill) → **Tier 2** (LLM code generation with python-pptx) → **Tier 3** (text-only deterministic). Once Tier 1 fails, it's bypassed for remaining chunks to avoid timeout cascading.

---

## User Journey

```
1. User provides a text prompt  ──────────► "Create a 10-slide deck about AI in healthcare"
                                             (optionally: --template corporate.pptx)

2. Brand/Style Analysis          ──────────► Detects "corporate" branding, searches guidelines

3. Storyboard Planning           ──────────► Opus optimizer creates per-slide storyboard

4. Content Generation            ──────────► Claude writes raw .pptx (chunked if >7 slides)

5. Image Planning + Generation   ──────────► Gemini decides image needs; NanoBanana generates

6. Template Assembly             ──────────► Knowledge file built; content mapped to template

7. Visual QA (optional)          ──────────► Vision model inspects + auto-corrects critical issues

8. Final PPTX delivered          ──────────► Professional presentation ready for use
```

---

## Technical Stack

| Component          | Technology                                                                              |
| ------------------ | --------------------------------------------------------------------------------------- |
| **Backend**        | Python 3.11+ (python-pptx, lxml, Pydantic, Pillow)                                     |
| **Orchestration**  | Agno Workflow framework (sequential Steps with shared `session_state`)                  |
| **Content LLM**    | Claude `claude-opus-4-6` (Anthropic) with `pptx` skill + streaming                        |
| **Planning LLM**   | Gemini `gemini-3-flash-preview` (Google) with structured output                       |
| **Brand Analysis** | Claude Sonnet + `web_search` tool (max 2 uses)                                          |
| **Image Gen**      | NanoBanana (Gemini Image Generation, 16:9 aspect ratio)                                 |
| **Vision QA**      | Gemini 2.5 Flash (vision) + LibreOffice headless PNG rendering                          |
| **File Download**  | Anthropic Files API (`beta.files.download`) with magic-byte file type detection          |
| **Infrastructure** | CLI-based; runs on local machine or server with Python + optional LibreOffice            |

---

## Current Status

| Metric              | Value                                                |
| -------------------- | ---------------------------------------------------- |
| **Development Phase** | Production (iterative improvement)                  |
| **Pipeline Modes**    | Single-call (≤7 slides) + Chunked (8-15+ slides)   |
| **Implemented Steps** | Steps 1-5 (Step 5 optional)                         |
| **Fallback Tiers**    | 3-tier (Skill → Code Gen → Text-only)               |
| **Test Coverage**     | Unit tests (brand parsing: 10 offline), visual cleanup tests |
| **Last Updated**      | 2026-03-01                                          |

---

## Success Metrics

| Metric               | Target / Current                                                   |
| --------------------- | ------------------------------------------------------------------ |
| **Reliability**       | 3-tier fallback ensures ~100% completion rate per chunk            |
| **Template Fidelity** | Knowledge-file-driven assembly matches template fonts, colors, layouts |
| **Visual Quality**    | WCAG contrast checks, ghost-text removal, fit-text auto-sizing     |
| **Performance**       | Tier 1: 2-5 min/chunk; Tier 3: <100ms (text-only, zero network)   |
| **Brand Accuracy**    | Autonomous web search for brand guidelines + structured intent     |
