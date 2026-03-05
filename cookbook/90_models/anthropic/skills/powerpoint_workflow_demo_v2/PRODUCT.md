# PRODUCT.md — Chunked PowerPoint Workflow v2

## Overview

AI-powered pipeline that transforms a text prompt into a polished, template-styled `.pptx` presentation. Combines multi-model AI content generation, intelligent brand detection, chunked orchestration for large decks, deterministic template assembly, and optional vision-based quality assurance — all orchestrated through Agno's Workflow framework.

**Entry point:** `powerpoint_chunked_workflow.py` (imports core logic from `powerpoint_template_workflow.py`)

## Solution Architecture

```
                    ┌─────────────────┐
                    │   YOUR INPUTS   │
                    │  Text Prompt    │
                    │  + Template     │
                    │  + Brand Intent │
                    └────────┬────────┘
                             │
                             ▼
                  powerpoint_chunked_
                  workflow.py (entry point)
                             │
              ┌──────────────┼──────────────┐
              │              │              │
              ▼              ▼              ▼
         Brand Parse    Storyboard    Chunk Generation
        (Sonnet+web)   (Opus+web)    (3-tier fallback)
              │              │              │
              └──────────────┼──────────────┘
                             │
                    ┌────────┴────────┐
                    │  Core Pipeline  │
                    │  (template_wf)  │
                    └────────┬────────┘
                             │
              ┌──────────────┼──────────────┐
              │              │              │
              ▼              ▼              ▼
        Image Pipeline  Template Assembly  Visual QA
        (Gemini+Nano)   (deterministic)   (Gemini vision)
              │              │              │
              └──────────────┼──────────────┘
                             │
                             ▼
                    ┌─────────────────┐
                    │  Final .pptx    │
                    └─────────────────┘
```

## Key Features

### 1. Chunked Orchestration
Splits large presentations (8-15+ slides) into configurable chunks (default: 3 slides/chunk), generates each independently, then merges with OPC-aware relationship management.

### 2. Brand-Aware Generation
Autonomous `brand_style_analyzer` agent uses a two-stage approach: a zero-cost Regex keyword pre-check to log explicit intent, followed by an OpenAI `gpt-4o-mini` extraction that runs *on every prompt* to catch any implicit styling directives the pre-check might miss (preserving Anthropic token budget). It discovers brand directives, researches guidelines online, and injects structured `BrandStyleIntent` into downstream calls. Template styling overrides query-level branding when a template is provided.

### 3. Multi-Provider Architecture
Supports swapping auxiliary agents via `--llm-provider {claude,openai,gemini}`. The workflow dynamically routes to different model variants depending on the selected provider mode:

| Agent Role | Claude (Default) | OpenAI | Gemini |
|------------|------------------|--------|--------|
| **Brand Analysis** | `claude-sonnet-4-6` | `gpt-5-mini` | `gemini-3-flash-preview` |
| **Storyboard / Plan** | `claude-opus-4-6` | `gpt-5.2` | `gemini-3-pro-preview` |
| **Fallback (Tier 2)** | `claude-haiku-4-5` | `gpt-5-mini` | `gemini-3-flash-preview` |
| **Image Plan** | `gemini-3-flash-preview`* | `gpt-5-mini` | `gemini-3-flash-preview` |
| **Visual Review** | `gemini-2.5-flash`* | `gpt-5-mini` | `gemini-2.5-flash` |
| **Search Tool** | `web_search_20250305` | `web_search_preview`| `search=True` |

*(Note: The core Content Generator (Tier 1) is hard-locked to **Claude Opus** to utilize its native PPTX skill capabilities. Additionally, Image Planning and Visual Review use Gemini models even under the Claude provider setting due to multimodal feature requirements).*

### 4. 3-Tier Fallback System
| Tier | Generator | Quality | Speed |
|------|-----------|---------|-------|
| 1 | Claude PPTX skill (`opus`) | 100% | 30s–5min/chunk |
| 2 | LLM code gen (`haiku` fallback) | 80–92% | 10–30s/chunk |
| 3 | Text-only (deterministic) | Structural | <1s/chunk |

### 5. Template-Faithful Assembly
Fully deterministic Step 3 builds a comprehensive **assembly knowledge file** — combining user intent, content plan, deep per-layout template analysis, and AI image assets — then maps content onto template layouts, fonts, colors, and placeholder regions. No LLM in the loop.

### 6. AI Image Generation
Gemini-based planning decides which slides need visuals. NanoBanana generates 16:9 PNG images, scaled/centered within the template's content area preserving aspect ratio.

### 7. Vision-Based Quality Assurance
Optional Gemini 2.5 Flash renders each slide to PNG (via LibreOffice headless), detects visual defects, and auto-corrects critical issues. Includes upfront missing key validation to prevent crash blocks. Fully non-blocking.

### 8. Global API Rate Limit Tracker
An internal `_RateLimitTracker` aggregates estimated token counts dynamically across all Claude API calls throughout the entire pipeline. It detects incoming transient `429` rate limit hits without breaking the machine state, and handles execution pacing using parameterized (random `60–120s`) inter-chunk sleeps with live countdowns.

### 9. Template Quality Safeguards
When using `--template`, five automatic safeguards protect presentation quality:
- **Per-slide rendering** — PPTX→PDF→PNG pipeline renders every slide individually (not just slide 1) so the visual review inspects all slides
- **Background detection** — 6-layer cascade (shape→slide→layout→master→theme→large shapes) correctly identifies dark template backgrounds for proper text contrast
- **Minimum font size** — Enforces 10pt body / 14pt title floor to prevent `fit_text()` from shrinking text to unreadable sizes
- **Overlap reflow** — Post-transfer detection resolves shape collisions by vertical stacking with minimum dimension enforcement
- **Template-aware LLM prompts** — Tier 2 code generation includes template background color, text color guidance, and layout constraints

Requires `poppler-utils` (`sudo apt-get install -y poppler-utils`). See [DESIGN_visual_quality.md](DESIGN_visual_quality.md) for technical details.

---

## Workflow Steps

| Step | Name | Agent/Function | Output |
|------|------|----------------|--------|
| 0 | Brand/Style Parse | `brand_style_analyzer` (Sonnet) | `BrandStyleIntent` |
| 1 | Optimize & Plan | `query_optimizer` (Opus) | `StoryboardPlan` + markdown files |
| 2 | Generate Chunks | Claude PPTX skill / fallback agents | N chunk `.pptx` files |
| 3 | Process Chunks | Deterministic template assembly | N assembled `.pptx` files |
| 4 | Visual Review | `slide_quality_reviewer` (Gemini) | Quality reports + fixes |
| 5 | Merge Chunks | `_merge_pptx_zip_level()` | Final `.pptx` |

---

## Technical Stack

| Component | Technology |
|-----------|-----------|
| **Runtime** | Python 3.8+ |
| **Orchestration** | Agno Workflow (sequential Steps + shared `session_state`) |
| **Provider Factory** | `agents/` package dynamically loads Swappable Agent Modules |
| **Content LLM** | **Claude** `claude-opus-4-6` + `pptx` skill + `context-1m` beta (Locked) |
| **Brand Analysis** | Two-Stage (Regex Fast-Check + OpenAI `gpt-4o-mini`) |
| **Storyboard** | Swappable (Claude Sonnet / GPT-5.2 / Gemini 3 Pro) |
| **Code Fallback** | Swappable (Claude Haiku / GPT-5 Mini) |
| **Image Planning** | Swappable (Gemini 3 Flash / GPT-5 Mini) |
| **Image Gen** | NanoBanana (Gemini, 16:9 aspect ratio) |
| **Vision QA** | Swappable (Gemini 2.5 Flash / GPT-5 Mini) + LibreOffice headless |
| **PPTX Engine** | python-pptx + lxml |
| **Data Validation** | Pydantic v2 |
| **File Download** | Anthropic Files API + magic-byte detection |

---

## File Inventory

| File | Lines | Purpose |
|------|-------|---------|
| `powerpoint_chunked_workflow.py` | ~3,380 | **Entry point** — chunked orchestration |
| `powerpoint_template_workflow.py` | ~7,218 | Core pipeline (imported via wildcard) |
| `agents/` | Package | Multi-provider agent implementations (Claude, OpenAI, Gemini) |
| `file_download_helper.py` | ~162 | Claude skill file download utility |
| `test_brand_style_parsing.py` | ~400 | Unit tests for brand/style parsing |

---

## Current Status

| Metric | Value |
|--------|-------|
| **Phase** | Production (iterative improvement) |
| **Capacity** | Scalable (chunked architecture via `powerpoint_chunked_workflow.py`) |
| **Steps** | 6 (Step 4 optional, Step 5 visual review optional) |
| **Fallback** | 3-tier (Skill → Code Gen → Text-only) |
| **Template Safeguards** | 5 (rendering, background detection, font guard, overlap reflow, prompt constraints) |
| **Tests** | Brand parsing: 10 unit tests, template fixes: 14 integration tests |
| **Last Updated** | 2026-03-05 |

---

## Success Metrics

| Metric | Target |
|--------|--------|
| **Reliability** | 3-tier fallback → ~100% completion rate per chunk |
| **Template Fidelity** | Knowledge-file-driven assembly matches template fonts, colors, layouts |
| **Visual Quality** | WCAG contrast ≥3.0, ghost-text removal, fit-text auto-sizing, 10pt/14pt font guard, overlap reflow |
| **Performance** | Tier 1: 2-5 min/chunk; Tier 2: 10-30s; Tier 3: <100ms |
| **Brand Accuracy** | Autonomous web search + structured `BrandStyleIntent` |
