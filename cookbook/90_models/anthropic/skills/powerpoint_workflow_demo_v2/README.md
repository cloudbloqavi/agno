# PowerPoint Chunked Workflow (v2)

AI-powered PowerPoint generation using Claude Agent Skills with chunked orchestration,
brand-aware styling, and a 3-tier fallback for production reliability.

## Quick Start

```bash
# Install dependencies
uv pip install agno anthropic python-pptx google-genai pillow lxml python-dotenv

# Set API keys (or use .env file)
export ANTHROPIC_API_KEY="your_anthropic_key" # Required for Claude PPTX generator
export OPENAI_API_KEY="your_openai_key"       # Required if using --llm-provider openai
export GOOGLE_API_KEY="your_google_key"       # Required if using --llm-provider gemini

# Run
python powerpoint_chunked_workflow.py \
    -p "Create a 10-slide AI transformation strategy deck"
```

## Entry Point

**`powerpoint_chunked_workflow.py`** is the single runnable script. It imports
all core pipeline logic from `powerpoint_template_workflow.py` via wildcard import.
*(Note: `powerpoint_template_workflow.py` acts strictly as a core library and cannot run independently).*

```
powerpoint_chunked_workflow.py   ← Run this (orchestration + chunking, ~3,380 lines)
    ↓ wildcard import
powerpoint_template_workflow.py  ← Core pipeline (content gen, images, template assembly, ~7,218 lines)
    ↓ import
agents/                          ← Provider-specific swappable agents (Claude, OpenAI, Gemini)
file_download_helper.py          ← Claude skill file download utility (~160 lines)
```

## Usage Examples

```bash
# Basic (auto-decides slide count, 3 slides per chunk):
python powerpoint_chunked_workflow.py \
    -p "Create a presentation about AI in healthcare"

# Brand-aware generation:
python powerpoint_chunked_workflow.py \
    -p "Create a 7-slide presentation about AI trends using Nike branding"

# With template (template styling overrides query branding):
python powerpoint_chunked_workflow.py \
    -t templates/my_template.pptx --chunk-size 4 \
    -p "12-slide enterprise AI strategy deck"

# Full options: visual review, custom output:
python powerpoint_chunked_workflow.py \
    -t templates/my_template.pptx \
    -p "12-slide enterprise AI strategy deck" \
    --chunk-size 3 --visual-review --visual-passes 5 \
    -o final_deck.pptx

# Quick: no images, no template:
python powerpoint_chunked_workflow.py \
    -p "Startup pitch deck for SaaS product" --no-images

# Enable template visual references (high-fidelity, higher token cost):
python powerpoint_chunked_workflow.py \
    -t templates/my_template.pptx --template-visuals \
    -p "Premium brand presentation"

# Switch LLM Provider for swappable agents (OpenAI gpt-5.2 or Gemini 3 Pro):
# Note: Content Generator always uses Claude to retain native PPTX skills
python powerpoint_chunked_workflow.py \
    -p "Quarterly review deck" --llm-provider openai

# Start at Tier 2 (skip Claude PPTX skill, use LLM code gen directly):
python powerpoint_chunked_workflow.py \
    -p "Quarterly review deck" --start-tier 2
```

## CLI Flags

| Flag | Description | Default |
|------|-------------|---------|
| `--prompt, -p` | Presentation topic prompt | (required) |
| `--template, -t` | Path to .pptx template file | None |
| `--output, -o` | Output filename | `presentation_chunked.pptx` |
| `--chunk-size` | Slides per API call | 1 |
| `--llm-provider` | LLM provider for swappable agents (claude, openai, gemini) | claude |
| `--max-retries` | Retries per chunk on failure | 2 |
| `--start-tier` | Starting tier (1=PPTX skill, 2=LLM code, 3=text-only) | 1 |
| `--no-images` | Skip AI image generation | disabled |
| `--min-images` | Min slides with AI images | 1 |
| `--visual-review` | Enable vision QA per chunk | disabled |
| `--visual-passes` | Max visual review passes per chunk | 3 |
| `--template-visuals, -tv` | Inject base64 template slide images into prompts | disabled |
| `--footer-text` | Footer text for all slides | None |
| `--date-text` | Date text for footer date placeholder | None |
| `--show-slide-numbers` | Show slide number placeholders | disabled |
| `--inter-chunk-delay-min` | Minimum random delay between chunks (ms) | Provider Specific |
| `--inter-chunk-delay-max` | Maximum random delay between chunks (ms) | Provider Specific |
| `--no-stream` | Disable streaming for agents | disabled |
| `--verbose, -v` | Enable verbose/debug logging | disabled |

## Swappable LLM Providers

The workflow uses a dynamic swappable agent architecture controlled by the `--llm-provider` flag. 

| Agent Role | Claude Mode (`claude`) | OpenAI Mode (`openai`) | Gemini Mode (`gemini`) |
|------------|------------------------|------------------------|------------------------|
| **Brand Parse** | `gpt-4o-mini` | `gpt-5-mini` | `gemini-3-flash-preview` |
| **Brand Parse Fallback**| `gpt-5-mini` | `gemini-3-flash-preview` | `gpt-4o-mini` |
| **Storyboard** | `claude-sonnet-4-6` | `gpt-5.2` | `gemini-3.1-pro-preview` |
| **Storyboard Fallback**| `gpt-5.2` | `gemini-3.1-pro-preview` | `gpt-5.2` |
| **Code Fallback** | `claude-sonnet-4-6` / `haiku` | `gpt-5.2` / `mini` | `gemini-3.1-pro-preview` / `flash` |
| **Image Plan** | `gemini-2.5-flash`* | `gpt-5-mini` | `gemini-2.5-flash` |
| **Image Plan Fallback**| `gpt-5-mini` | `gemini-2.5-flash` | `gpt-5-mini` |
| **Visual QA** | `gemini-2.5-flash`* | `gpt-5-mini` (vision) | `gemini-2.5-flash` |
| **Visual QA Fallback**| `gpt-5-mini` | `gemini-2.5-flash` | `gpt-5-mini` |
| **Content Gen** | `claude-opus-4-6` (Locked) | `claude-opus-4-6` (Locked) | `claude-opus-4-6` (Locked) |

*\* Note: Even in `claude` mode, Image Planning and Visual QA explicitly default to Gemini models because Anthropic models lack the requisite API multimodal features.*

## Workflow Steps

| Step | Name | Description |
|------|------|-------------|
| 0 | Brand/Style Parse | (within Step 1) Detects brand intent via `brand_style_analyzer` agent |
| 1 | Optimize & Plan | LLM creates storyboard with brand-aware search, visual profile analysis, and per-slide content |
| 2 | Generate Chunks | Claude PPTX skill called N times; 3-tier fallback per chunk |
| 3 | Process Chunks | Template assembly + image pipeline per chunk |
| 4 | Visual Review | *(optional)* Gemini vision QA per chunk |
| 5 | Merge Chunks | OPC-aware merge of all chunk PPTX files into final output |

## Token Usage & Cost Summary

When running with the `--verbose` or `-v` flag, the workflow provides a detailed breakdown of token consumption and estimated cost across all used LLM models at the end of the execution.

It tracks:
- **Total Calls**: Number of API requests per model.
- **Tokens (In/Out)**: Input and output token counts.
- **Estimated Cost**: Calculated based on approximate 2026 provider pricing.

This summary is printed to both the console (log) and the `OUTPUT.md` file.

## 3-Tier Fallback System with Universal HA

| Tier | Generator | Quality | Trigger |
|------|-----------|---------|---------|
| 1 | Claude PPTX skill | 100% — charts, tables, rich visuals | Primary |
| 2 | LLM code generation (OpenAI 4-step fallback on fail) | 80–92% — python-pptx native charts | Timeout >300s or all retries fail |
| 3 | python-pptx text-only | Structural only | Complete LLM Failure |

## Brand/Style-Aware Parsing

The workflow detects brand directives in user prompts (e.g. "using Nike branding"):
- **`brand_style_analyzer`** agent (Claude Sonnet + web_search) extracts brand colors, tone, typography
- **Template override**: When a template is provided, its styling takes precedence
- **Visual Profile Analysis**: Programmatic analysis of template density and style is injected into the optimizer prompt
- Brand context is injected into optimizer, Tier 1, and Tier 2 prompts
- **Audience-to-Style mapping**: Optimizer determines the target audience and selects one of 4 visual archetypes (bold_modern, clean_minimal, etc.)

## Adaptive No-Template Design System

When no template is provided, the system builds a visual identity on the fly:
- **Presets**: Maps audience to specific color/font pairings (e.g. `corporate_professional` for board meetings).
- **Brand Overrides**: Brand palette and fonts override preset defaults for consistent ID.
- **Verbose Logging**: Clear `[VERBOSE] [DESIGN SYSTEM]` logs for color/font overrides.

## Output

All output goes to `output_chunked/chunked_workflow_work/session_<id>/` including:
- `storyboard/` — slide-by-slide markdown storyboard
- `chunk_NNN.pptx` — individual chunk files
- `chunk_NNN_assembled.pptx` — template-assembled chunks
- `presentation_chunked.pptx` — final merged output
- `prompt_*.txt` — saved prompts for debugging

Console output is redirected to `OUTPUT.md` in the script directory.

## Documentation Index

### 📐 Architecture & Design

| Document | Description |
|----------|-------------|
| [Architecture](ARCHITECTURE_powerpoint_chunked_workflow.md) | Full system architecture — pipeline modes, data models, step-by-step flow, session state schema, CLI reference, template style extraction |
| [Visual Quality Design](DESIGN_visual_quality.md) | Visual quality improvement design — Phase 1 (deterministic fixes), Phase 2 (visual review agent), Phase 3 (template quality safeguards) |
| [Product Spec](PRODUCT.md) | Product specification and requirements |

### 🤖 Agents & Skills

| Document | Description |
|----------|-------------|
| [Agent Catalog](AGENTS.md) | All agents (brand parser, storyboard, code fallback, image planner, visual QA), their roles, models, and swappable provider configs |
| [Skills](./skills/) | Skill definition files (core, domain, meta, tools) used by Claude agents |

### 🔧 Operations & Debugging

| Document | Description |
|----------|-------------|
| [Output Log](OUTPUT.md) | Latest console output from the most recent workflow run (auto-generated) |
| [Test Log](TEST_LOG.md) | Test execution history and results |
| [Scratchpad](SCRATCHPAD.md) | Notes, work-in-progress tracker, and development journal |
| [Library Patches](lib_patches/PATCHES.md) | Required patches to `agno` library: `claude.py` (structured output fix) and `python.py` (PythonTools error logging) |

### 📁 Source Files

| File | Purpose |
|------|---------|
| `powerpoint_chunked_workflow.py` | **Main entry point** — chunked orchestration, storyboard planning, 3-tier fallback, chunk merging (~3,800 lines) |
| `powerpoint_template_workflow.py` | Core pipeline — content gen, images, template assembly, visual review, all helpers (~8,000 lines) |
| `agents/` | Swappable provider agents package (Claude, OpenAI, Gemini) and shared Pydantic schemas |
| `file_download_helper.py` | Claude skill file download utility via Anthropic Files API (~160 lines) |
| `test_brand_style_parsing.py` | Unit tests for brand/style parsing |
| `.env` | API keys (`ANTHROPIC_API_KEY`, `OPENAI_API_KEY`, `GOOGLE_API_KEY`) |
| `templates/` | 9 `.pptx` template files for template-assisted generation |

---

## Prerequisites

### Python Dependencies

```bash
uv pip install agno anthropic openai google-genai python-pptx pillow lxml python-dotenv
```

### API Keys

```bash
export ANTHROPIC_API_KEY="your_anthropic_key"   # ALWAYS required (Content Generator)
export OPENAI_API_KEY="your_openai_key"          # Required if using --llm-provider openai
export GOOGLE_API_KEY="your_google_key"          # Required if using --llm-provider gemini
```

Or create a `.env` file in this directory with the same key-value pairs.

### System Dependencies

```bash
# Required for --visual-review (slide rendering)
sudo apt-get install -y libreoffice

# Required for per-slide PNG rendering (used by visual review)
sudo apt-get install -y poppler-utils
```

> **Note:** Both system dependencies are only required when using `--visual-review` with `--template`. The workflow skips visual review gracefully if either is missing.


## Logging Conventions

```
[TIMING] step_XXX completed in X.Xs   — Always printed
[ERROR] ...                            — Always printed
[WARNING] ...                          — Always printed
[BRAND] ...                            — Always printed (brand detection)
[BRAND OVERRIDE] ...                   — Always printed (template overrides query)
[RENDER] ...                           — Always printed (slide PNG rendering pipeline)
[BG DETECT] ...                        — Always printed (background color detection)
[FONT GUARD] ...                       — Always printed (minimum font size enforcement)
[OVERLAP FIX] ...                      — Always printed (shape overlap detection & reflow)
[TEMPLATE CTX] ...                     — Always printed (template context injected into LLM prompt)
[VERBOSE] ...                          — Only with --verbose / -v
```

## Template Quality Safeguards

When using `--template`, several automatic safeguards protect presentation quality:

| Safeguard | What It Does |
|-----------|-------------|
| **Per-slide rendering** | Renders every slide to PNG via PPTX→PDF→PNG pipeline (`pdftoppm`) for visual review |
| **Background detection** | 6-layer detection (shape → slide → layout → master → theme → large shapes) prevents wrong contrast |
| **Minimum font size** | Enforces 10pt body / 14pt title minimum — prevents unreadable text from `fit_text()` shrinkage |
| **Layout Sanitization** | 8-pass correction engine: boundary clamping, min size enforcement, overlap orphan removal, icon purging, column alignment snapping, and iterative reflow (preserves structural template elements) |
| **Smart Template Purge** | Intelligent classification (Fix 13) preserves branded headers, footers, and decorative motifs while clearing placeholder text |
| **Template-aware prompts** | Tier 2 LLM code generation includes spatial grid rules, decoration bans, and layout constraints |
| **Visual Profile Analysis** | Programmatic analysis of layout density, dark/light dominance, and content constraints injected into the storyboard optimizer |
| **Accent Line Purge** | Automatic detection and removal of slanting/diagonal LINE geometry from templates |
| **Footer Numbering** | Incremental page numbering (1, 2, 3...) for cloned template footer bands |
| **Content Sanitization** | Heuristic removal of hardcoded template text (e.g. "AGILE PROJECT PLAN") while preserving branding motifs |
