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
| `--chunk-size` | Slides per Claude API call | 3 |
| `--llm-provider` | LLM provider for swappable agents (claude, openai, gemini) | claude |
| `--max-retries` | Retries per chunk on failure | 2 |
| `--start-tier` | Starting tier (1=PPTX skill, 2=LLM code, 3=text-only) | 1 |
| `--no-images` | Skip AI image generation | disabled |
| `--min-images` | Min slides with AI images | 1 |
| `--visual-review` | Enable vision QA per chunk | disabled |
| `--visual-passes` | Max visual review passes per chunk | 3 |
| `--footer-text` | Footer text for all slides | None |
| `--date-text` | Date text for footer date placeholder | None |
| `--show-slide-numbers` | Show slide number placeholders | disabled |
| `--no-stream` | Disable streaming for agents | disabled |
| `--verbose, -v` | Enable verbose/debug logging | disabled |

## Workflow Steps

| Step | Name | Description |
|------|------|-------------|
| 0 | Brand/Style Parse | (within Step 1) Detects brand intent via `brand_style_analyzer` agent |
| 1 | Optimize & Plan | LLM creates storyboard with brand-aware search, tone, per-slide content |
| 2 | Generate Chunks | Claude PPTX skill called N times; 3-tier fallback per chunk |
| 3 | Process Chunks | Template assembly + image pipeline per chunk |
| 4 | Visual Review | *(optional)* Gemini vision QA per chunk |
| 5 | Merge Chunks | OPC-aware merge of all chunk PPTX files into final output |

## 3-Tier Fallback System

| Tier | Generator | Quality | Trigger |
|------|-----------|---------|---------|
| 1 | Claude PPTX skill | 100% — charts, tables, rich visuals | Primary |
| 2 | LLM code generation | 80–92% — python-pptx native charts | Timeout >300s or all retries fail |
| 3 | python-pptx text-only | Structural only | Tier 2 failure |

## Brand/Style-Aware Parsing

The workflow detects brand directives in user prompts (e.g. "using Nike branding"):
- **`brand_style_analyzer`** agent (Claude Sonnet + web_search) extracts brand colors, tone, typography
- **Template override**: When a template is provided, its styling takes precedence
- Brand context is injected into optimizer, Tier 1, and Tier 2 prompts

## Output

All output goes to `output_chunked/chunked_workflow_work/session_<id>/` including:
- `storyboard/` — slide-by-slide markdown storyboard
- `chunk_NNN.pptx` — individual chunk files
- `chunk_NNN_assembled.pptx` — template-assembled chunks
- `presentation_chunked.pptx` — final merged output
- `prompt_*.txt` — saved prompts for debugging

Console output is redirected to `OUTPUT.md` in the script directory.

## File Inventory

| File | Purpose |
|------|---------|
| `powerpoint_chunked_workflow.py` | **Main entry point** — chunked orchestration |
| `powerpoint_template_workflow.py` | Core pipeline (imported by chunked workflow) |
| `agents/` | Package containing provider agents (Claude, OpenAI, Gemini) and shared schemas |
| `file_download_helper.py` | Claude skill file download utility |
| `test_brand_style_parsing.py` | Unit tests for brand/style parsing |
| `.env` | API keys (ANTHROPIC_API_KEY, GOOGLE_API_KEY) |
| `templates/` | 9 .pptx template files |
| `lib_patches/` | External library patches documentation |
| `.skills/` | Skill definition files (core, domain, meta, tools) |
| `AGENTS.md` | Agent catalog and architecture |
| `PRODUCT.md` | Product specification |
| `SCRATCHPAD.md` | Notes and work-in-progress tracker |
| `ARCHITECTURE_powerpoint_chunked_workflow.md` | Detailed architecture docs |
| `DESIGN_visual_quality.md` | Visual quality review design docs |
| `TEST_LOG.md` | Test execution log |

## Prerequisites

- Python 3.8+
- `uv pip install agno anthropic openai google-genai python-pptx pillow lxml python-dotenv`
- `ANTHROPIC_API_KEY` (Always required for Claude's PPTX generator)
- `OPENAI_API_KEY` (Required if using `--llm-provider openai`)
- `GOOGLE_API_KEY` (Required if using `--llm-provider gemini`)
- LibreOffice (optional, for `--visual-review` slide rendering)

## External Library Patches

This project required modifications to two files in the `agno` library.
See [`lib_patches/PATCHES.md`](lib_patches/PATCHES.md) for details:
- `claude.py` — structured output detection fix for `claude-sonnet-4-6`
- `python.py` — enhanced PythonTools error logging with tracebacks

## Additional Documentation

- [Architecture](ARCHITECTURE_powerpoint_chunked_workflow.md) — detailed system design
- [Visual Quality Design](DESIGN_visual_quality.md) — visual review step design

## Logging Conventions

```
[TIMING] step_XXX completed in X.Xs   — Always printed
[ERROR] ...                            — Always printed
[WARNING] ...                          — Always printed
[BRAND] ...                            — Always printed (brand detection)
[BRAND OVERRIDE] ...                   — Always printed (template overrides query)
[VERBOSE] ...                          — Only with --verbose / -v
```
