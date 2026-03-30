# PRODUCT.md — Chunked PowerPoint Workflow v2

## Overview

AI-powered pipeline that transforms a text prompt into a polished, template-styled `.pptx` presentation. Imagine an **AI Creative Director** that understands your brand, researches your mission, and builds a professional deck from scratch—all while maintaining the highest visual standards.

**Entry point:** `powerpoint_chunked_workflow.py` (orchestrates branding, planning, and generation).

---

## 📘 Product Manager's Operational Guide

This section explains how to leverage the workflow's key visual and branding features in your daily operations.

### 1. How to Use "Themes" (Theme Factory)
The workflow includes an autonomous **Theme Factory** that stores high-fidelity design definitions.
*   **Usage**: Simply add `Stictly use the '{theme-name}' theme` to your prompt.
*   **Predefined Themes**:
    *   `midnight-galaxy`: Dark, premium, cosmic purple/navy aesthetic.
    *   `arctic-frost`: Clean, icy blue, high-modern professional look.
    *   `botanical-garden`: Organic, green-toned, sustainable-focused design.
*   **Benefit**: Guaranteed color harmony and typography without needing a template file.

### 2. Autonomous Branding (The "Acting as..." Prompt)
The engine doesn't just write text; it researches your brand identity live.
*   **Prompting**: Start your request with `"Acting as a marketing manager at Tissot..."`.
*   **What Happens**: The system triggers an autonomous web search to find the latest Tissot logos, hex codes, and "Swiss Heritage" brand tone.
*   **Benefit**: Zero manual configuration for brand-aligned presentations.

### 3. Template vs. No-Template Modes
*   **No-Template**: Best for rapid brainstorming or pitch ideas. The system builds a "Design System" on the fly based on your brand or chosen theme.
*   **With Template (`-t`)**: Best for board decks or corporate reports. The system "inherits" the template's exact colors, fonts, and layouts, then cleans up placeholder text automatically.

### 4. Cost & Performance Tracking
Use the `--verbose` flag to see a **Usage Summary** at the end of every run. This helps you track the budget per presentation and understand which AI models were utilized.

---

## Solution Architecture

```mermaid
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
          Brand Parse    Visual Profile    Storyboard    Chunk Generation
         (Haiku+web)     (Analysis)       (Haiku+web)    (3-tier fallback)
               │              │                │              │
               └──────────────┴────────────────┼──────────────┘
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

---

## Key Features

### 1. Zero-Failure Delivery (Chunking & Fallbacks)
Large decks (10-20 slides) are split into small groups (chunks). Each group has a **3-Tier Safety Net**:
*   **Primary**: High-fidelity native PowerPoint shapes (charts, grids).
*   **Secondary**: Visual-heavy code generation.
*   **Tertiary**: Text-only structural fallback.
*   *Outcome: You always get a slide, even if the primary AI engine is under high load.*

### 2. Intelligent Brand Discovery
The `brand_style_analyzer` researches your brand guidelines online. It extracts hex codes and typography, then propagates them through the entire generation process to ensure consistent identity.

### 3. High-Fidelity Theme Factory
A dedicated design layer that stores expert-curated palettes. When you select a theme like `midnight-galaxy`, the engine overrides generic defaults with high-contrast, premium styling rules specific to that aesthetic.

### 4. Template-Agnostic Sanitization
When using your corporate `.pptx` template, the engine automatically:
*   Deletes "PLACEHOLDER TEXT" without touching your logo or branded borders.
*   Enforces a **Font Floor** (min 10pt body / 14pt title) to ensure readability.
*   Corrects shape overlaps and text overflows in real-time.

---

## Multi-Provider Architecture

The workflow dynamically routes to the best model for each task:

| Agent Role | Claude (Default) | OpenAI | Gemini |
|------------|------------------|--------|--------|
| **Brand Analysis** | `claude-haiku-4-5` | `gpt-5-mini` | `gemini-3-flash-preview` |
| **Storyboard / Plan** | `claude-haiku-4-5` | `gpt-5.2` | `gemini-3.1-pro-preview` |
| **Code Fallback** | `claude-haiku-4-5` | `gpt-5.2` / `mini` | `gemini-3.1-pro-preview` / `flash` |
| **Image Planning** | `gemini-2.5-flash`* | `gpt-5-mini` | `gemini-2.5-flash` |
| **Visual Review** | `gemini-2.5-flash`* | `gpt-5-mini` | `gemini-2.5-flash` |

---

## Success Metrics

| Metric | Business Target |
|--------|----------------|
| **Reliability** | ~100% completion rate via 3-tier fallback. |
| **Brand Fidelity** | Automatic match of brand colors/fonts via web-research. |
| **Visual Accessibility** | WCAG contrast ≥3.0 and readable font sizes (min 10pt). |
| **Performance** | Rapid turnaround (approx. 30s - 1min per slide group). |
| **Efficiency** | Significant reduction in manual Slide-Master editing. |

---

## Technical Inventory

| File | Purpose |
|------|---------|
| `powerpoint_chunked_workflow.py` | **Main entry point** — chunked orchestration & state. |
| `powerpoint_template_workflow.py` | Technical core — layout logic, sanitization, and assembly. |
| `theme-factory/` | High-fidelity design definitions and theme references. |
| `file_download_helper.py` | Asset management for AI-generated artifacts. |

> [!NOTE]
> Requires `poppler-utils` and `libreoffice` for visual review features. See [DESIGN_visual_quality.md](DESIGN_visual_quality.md) for deeper technical layout rules.

---

## 🚀 Supplemental PM Guide: Setup & Testing

### 1. Step-by-Step Environment Setup

#### A. Install Python 3.10+
- **Windows**: [Download and install](https://www.python.org/downloads/windows/) or use `winget install Python.Python.3.11`.
- **macOS**: `brew install python` (Requires [Homebrew](https://brew.sh/)).
- **WSL (Ubuntu)**: `sudo apt update && sudo apt install -y python3 python3-pip`.

#### B. Install UV (The Agno Recommended Package Manager)
- **Windows (PowerShell)**:
  ```powershell
  powershell -c "irm https://astral.sh/uv/install.ps1 | iex"
  ```
- **macOS / Linux / WSL**:
  ```bash
  curl -LsSf https://astral.sh/uv/install.sh | sh
  source $HOME/.cargo/env  # Refresh terminal environment
  ```

#### C. Install Project Dependencies
Once `uv` is installed, run the following to install the required AI frameworks, PowerPoint libraries, and Langfuse observability packages:
```bash
uv pip install agno anthropic openai google-genai python-pptx pillow lxml python-dotenv \
    openinference-instrumentation-agno opentelemetry-exporter-otlp-proto-http opentelemetry-sdk
```

#### D. Install System Libraries (Critical for Visual Review)
The vision-based defect fixing (Visual QA) requires these tools:

- **Windows**:
  1. Install [LibreOffice](https://www.libreoffice.org/download/download/): `winget install LibreOffice.LibreOffice`
  2. Install [Poppler](https://github.com/oschwartz10612/poppler-windows/releases/): Download and add to your PATH.
- **macOS**:
  ```bash
  brew install --cask libreoffice
  brew install poppler
  ```
- **WSL (Ubuntu)**:
  ```bash
  sudo apt update && sudo apt install -y libreoffice poppler-utils
  ```

#### E. Configure `.env` File
Create a `.env` file in the root directory.

```env
# AI Provider Keys
ANTHROPIC_API_KEY="sk-ant-..."      # Required for Claude Content Generator
GOOGLE_API_KEY="AIzaSy..."          # Required for Image Gen & Visual QA
OPENAI_API_KEY="sk-proj-..."        # Required if using --llm-provider openai

# Observability (Langfuse)
LANGFUSE_PUBLIC_KEY="pk-lf-..."     # For tracing and quality monitoring
LANGFUSE_SECRET_KEY="sk-lf-..."
OTEL_EXPORTER_OTLP_ENDPOINT="https://cloud.langfuse.com/api/public/otel"
```

### 📋 CLI Reference: Command Line Arguments

| Flag | Meaning | Default |
| :--- | :--- | :--- |
| `--prompt, -p` | The topic or brand directive for the presentation. | (Required) |
| `--template, -t` | Path to a `.pptx` template file to inherit styling. | `None` |
| `--output, -o` | Desired filename for the final presentation. | `presentation_chunked.pptx` |
| `--llm-provider` | Provider for all swappable agents (`claude`, `openai`, `gemini`). | `claude` |
| `--chunk-size` | Number of slides generated per Claude API call. | `1` |
| `--start-tier` | Starting quality tier (`1`=Skill, `2`=Code-Gen, `3`=Text-Only). | `1` |
| `--visual-review` | Enable vision-agent QA to fix design defects. | `False` |
| `--visual-passes` | Maximum correction attempts per slide. | `3` |
| `--no-images` | Disable AI image generation entirely. | `False` |
| `--no-web-search` | Disable the agent's ability to research the prompt online. | `False` |
| `--verbose, -v` | Enable detailed debug logs and Cost/Token summary. | `False` |

### 🧪 Example Gallery: Testing the Pipeline

#### 1. Basic Brainstorming (Auto-decide slides)
```bash
python powerpoint_chunked_workflow.py -p "Future of AI in Healthcare" \
--chunk-size 1 \
--start-tier 2 \
--no-images \
--verbose \
--visual-review \
--visual-passes 3 \
--llm-provider openai
```

#### 2. Branded Executive Pitch (Live Research)
```bash
python powerpoint_chunked_workflow.py \
    -p "Create a 7-slide presentation about AI trends using Nike branding"
    --chunk-size 1 \
    --start-tier 2 \
    --no-images \
    --verbose \
    --visual-review \
    --visual-passes 3 \
    --llm-provider openai
```
*Detects "Nike branding", performs web search for hex codes, and applies them.*

#### 3. Branded Corporate Template (Large Deck)
```bash
python powerpoint_chunked_workflow.py \
    -t templates/my_template.pptx
    --chunk-size 1 \
    --start-tier 2 \
    --no-images \
    --verbose \
    --visual-review \
    --visual-passes 3 \
    --llm-provider openai
    -p "15-slide enterprise AI strategy for the board"
```

#### 5. High-Stakes Review (Vision-Agent QA)
```bash
python powerpoint_chunked_workflow.py \
    -t templates/my_template.pptx \
    --chunk-size 1 \
    --start-tier 2 \
    --no-images \
    --verbose \
    --visual-review \
    --visual-passes 3 \
    --llm-provider openai
    -p "Complex multi-chart financial report"
```
*LibreOffice renders slides; Gemini Vision inspects and fixes overlaps.*