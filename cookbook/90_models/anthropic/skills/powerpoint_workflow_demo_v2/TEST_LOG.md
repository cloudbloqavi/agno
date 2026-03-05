# TEST_LOG — powerpoint_workflow_demo

---

### powerpoint_template_workflow.py — Phase 1 deterministic visual quality improvements

**Status:** NOT RUN

**Description:** Five targeted deterministic improvements applied to address visual quality issues: shape rescaling from Claude's default slide dimensions to template content region (`_rescale_shape_xml`, `_transfer_shapes`); footer standardization with new `--footer-text`, `--date-text`, `--show-slide-numbers` CLI flags and idx=10/11/12 placeholder population before cleanup runs; line-length wrap factor in `_compute_text_ratio` so slides with long bullets get proportionally more text region height; source dimensions stored in session_state immediately after Claude PPTX is opened (`step_generate_content`); `fit_text()` fallback hardening in `_populate_placeholder_with_format` to cap `rPr.sz` OOXML attributes for viewers without `MSO_AUTO_SIZE` support.

**Result:** Syntax and lint verified only. `./scripts/format.sh` and `./scripts/validate.sh` pass with exit code 0. Pre-existing F841 lint warnings (`ns_a`, `bodyPr`, `lstStyle`, `layout_names`) also fixed. Runtime end-to-end test requires `ANTHROPIC_API_KEY` and `GOOGLE_API_KEY`.

---

### powerpoint_template_workflow.py — Phase 2 optional visual review agent (Step 5)

**Status:** NOT RUN

**Description:** Optional `--visual-review` Step 5 added that renders each slide to PNG via LibreOffice headless subprocess, inspects each with Gemini 2.5 Flash vision using `output_schema=SlideQualityReport`, and applies safe corrections for `critical`-severity issues. Components added: `ShapeIssue` and `SlideQualityReport` Pydantic models, `PresentationQualityReport` stored in `session_state["quality_report"]`, `_render_pptx_to_images()`, `slide_quality_reviewer` agent (Gemini 2.5 Flash), `_apply_visual_corrections()`, and `step_visual_quality_review()` executor. Correction scope v1: contrast (`increase_contrast`), ghost text / empty placeholders (`clear_placeholder`), text overflow (`reduce_font_size`). Visual blandness detected and warned only. Shape repositioning deferred to v2.

**Result:** Syntax and lint verified only. `./scripts/format.sh` and `./scripts/validate.sh` pass with exit code 0. Runtime test requires LibreOffice and `GOOGLE_API_KEY`.

---

### powerpoint_template_workflow.py — Live test: dynamicv1.pptx (7-slide SpaceTech deck)

**Status:** PASS

**Description:** End-to-end live test using `dynamicv1.pptx` template, 7-slide SpaceTech startup prompt, with `--visual-review` enabled. Command: `python powerpoint_template_workflow.py -t dynamicv1.pptx -o report18.pptx -p "Create a 7-slide presentation on a SpaceTech Startup based out of India" --visual-review`. Four bugs were discovered during this test and fixed: (1) visual review reported 0 slides inspected because `_render_pptx_to_images()` glob patterns did not match LibreOffice's zero-padded output naming (e.g. `report18001.png`); (2) footer missing on title slide because `dynamicv1.pptx` title layout has no idx=11 placeholder — fallback text box at slide's bottom 7% zone added; (3) text overflow on slides 2 and 3 due to `LINE_SPACING_FACTOR=1.5` and `hard_max=18` being too generous; (4) images placed in tiny footer-left slot because `_best_visual_placeholder()` score tuple `(overlap==0, -overlap, area)` allowed zero-overlap tiny slots to outrank large content placeholders. A deterministic reassembly test using the existing `skill_output_9Nsopu7e.pptx` intermediate verified all 7 slides assemble correctly with footer injection, no text overflow, correct chart/table placement, and empty placeholder cleanup.

**Result:** All 7 slides assembled correctly after bug fixes. Output file `reassembly_candidate.pptx` (153,600 bytes, 7 slides) confirmed. All four bugs fixed and validated through deterministic reassembly.

---

### powerpoint_chunked_workflow.py — Initial implementation

**Status:** NOT RUN

**Description:** New chunked PPTX generation workflow implementing: query optimization with storyboard planning (Step 1: `step_optimize_and_plan`), chunked Claude PPTX skill calls with configurable `--chunk-size` (Step 2: `step_generate_chunks`), per-chunk template assembly and image pipeline when `--template` is provided (Step 3: `step_process_chunks`), per-chunk visual QA with up to `--visual-passes` passes when `--visual-review` and `--template` are provided (Step 4: `step_visual_review_chunks`), and OPC-aware PPTX merge of all chunks (Step 5: `step_merge_chunks`). New CLI args: `--chunk-size` (default 3), `--max-retries` (default 2).

**Result:** Syntax check PASS. Runtime test requires `ANTHROPIC_API_KEY` and `GOOGLE_API_KEY`.

---

### powerpoint_chunked_workflow.py — Three targeted improvements

**Status:** NOT RUN

**Description:** Three improvements applied: (1) `--visual-passes` CLI argument replacing hardcoded `range(3)` in `step_visual_review_chunks()`, with session state key `visual_passes` and updated "pass X/Y" messages; (2) consistent `if VERBOSE:` guards throughout, gating debug prints in `step_optimize_and_plan`, `generate_chunk_pptx`, `step_generate_chunks`, `step_process_chunks`, `step_visual_review_chunks`, `merge_pptx_files`, and `step_merge_chunks`; (3) step-level and sub-operation timing via `[TIMING]` print tags for all major functions.

**Result:** Syntax check PASS. Runtime test requires `ANTHROPIC_API_KEY` and `GOOGLE_API_KEY`.

---

### powerpoint_chunked_workflow.py — 3-tier fallback system

**Status:** NOT RUN

**Description:** 3-tier chunk generation fallback system added to `step_generate_chunks()`. Tier 1: Claude PPTX skill (`generate_chunk_pptx()`) run via `ThreadPoolExecutor` with `CHUNK_TIMEOUT_SECONDS=300` per attempt. On timeout or all retries exhausted, sets `session_state["use_fallback_generator"]=True` and all remaining chunks bypass Tier 1. Tier 2: LLM code generation (`generate_chunk_pptx_v2()`) using `fallback_code_agent` (`claude-opus-4-6` without `context-1m` beta + `PythonTools`) that writes and executes a `python-pptx` + `matplotlib` script; produces real Office charts and tables; no internal retry. Tier 3: direct `python-pptx` text-only fallback (`generate_chunk_pptx_fallback()`); uses `FALLBACK_SLIDE_LAYOUT_MAP` for layout selection; zero network I/O; always succeeds. Also added `--start-tier` CLI argument (default 1) allowing users to start directly at Tier 2 or Tier 3.

**Result:** Syntax check PASS. Runtime test requires `ANTHROPIC_API_KEY`. Tier 2 and Tier 3 output files are standard `.pptx` compatible with downstream template assembly and merge steps.

---

## Live Test Instructions

To perform a live end-to-end test once API keys are available:

```bash
# Prerequisites
export ANTHROPIC_API_KEY="..."
export GOOGLE_API_KEY="..."

# Install LibreOffice (for --visual-review)
# apt-get install libreoffice

# Basic test (no visual review)
.venvs/demo/bin/python cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/powerpoint_chunked_workflow.py \
    --template cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/my_template.pptx \
    --output /tmp/test_output.pptx \
    -v

# With visual review
.venvs/demo/bin/python cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/powerpoint_chunked_workflow.py \
    --template cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/my_template.pptx \
    --output /tmp/test_output_reviewed.pptx \
    --visual-review \
    --footer-text "Confidential" --show-slide-numbers \
    -v

# Text-only (no images, faster)
.venvs/demo/bin/python cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/powerpoint_chunked_workflow.py \
    --template cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/my_template.pptx \
    --no-images \
    --output /tmp/test_no_images.pptx \
    -v

# Multi-Provider Test (OpenAI for auxillary agents, Claude for Tier 1)
export OPENAI_API_KEY="..."
.venvs/demo/bin/python cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/powerpoint_chunked_workflow.py \
    --template cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/my_template.pptx \
    --llm-provider openai \
    --output /tmp/test_openai_aux.pptx \
    -v

# Fallback Tier Testing (Force Tier 2 LLM Code Generation)
.venvs/demo/bin/python cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/powerpoint_chunked_workflow.py \
    --template cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/my_template.pptx \
    --start-tier 2 \
    --output /tmp/test_tier2_fallback.pptx \
    -v

# Explicit Brand/Style Intent
.venvs/demo/bin/python cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/powerpoint_chunked_workflow.py \
    --prompt "Create a 5-slide deck about cloud computing in the style of Apple, using minimalist design and SF Pro typography." \
    --output /tmp/test_apple_branding.pptx \
    -v
```

What to verify in the output:
- Title slide has correct title and subtitle
- Content slides use template fonts, colors, and layout
- Tables use template header/cell styling (not plain Calibri defaults)
- Charts use template series colors (not default blue)
- Shapes from Claude's slide appear at correct positions (not off-screen or at default Claude coordinates)
- Footer text appears on all slides when `--footer-text` is used
- Slide numbers appear when `--show-slide-numbers` is used
- With `--visual-review`: `quality_report` in session_state, critical issues corrected

---

### powerpoint_chunked_workflow.py — Brand/style-aware query parsing

**Status:** PASS (offline tests)

**Description:** Brand/style-aware query parsing subsystem added to the chunked workflow. New components: `BrandStyleIntent` Pydantic model (brand_name, color_palette, tone_override, typography_hints, style_keywords, source, content_query); `brand_style_analyzer` Agno agent (Claude Sonnet with `web_search`, max 2 uses, `output_schema=BrandStyleIntent`) that autonomously detects branding directives and decides whether to search for brand guidelines; `parse_brand_style_intent()` function that calls the agent; `extract_style_from_template()` that reads .pptx theme XML for colors, fonts, and company name heuristics; `_build_brand_override_log()` for structured override logging when template styling overrides query-level branding; `_format_brand_context_for_prompt()` for markdown prompt injection. Integration points: Step 1 (`step_optimize_and_plan`) calls parsing before optimizer, injects brand context into optimizer prompt and search queries, handles template override; Tier 1 (`generate_chunk_pptx`) injects brand context as `## Brand/Style Guidance` section in chunk prompts; Tier 2 (`generate_chunk_pptx_v2`) appends brand context to GLOBAL CONTEXT in code-gen prompt; Tier 3 unchanged (no LLM call); `main()` initializes `brand_style_intent: None` in session_state. Downstream steps (process, visual review, merge) unchanged per SCRATCHPAD requirement.

**Result:** 10/10 offline tests PASS (`test_brand_style_parsing.py`): model defaults, field values, JSON roundtrip, format_brand_context (empty/populated), build_brand_override_log, extract_style_from_template (basic, company name, nonexistent), template override flow. Existing `test_pptx.py` passes (no regression). Python `ast.parse()` syntax check PASS. Runtime end-to-end test requires `ANTHROPIC_API_KEY`.

**Test file:** `test_brand_style_parsing.py` (self-contained, no agno/anthropic dependency needed)

**Test commands:**
```bash
python test_brand_style_parsing.py     # 10 offline brand/style tests
python test_pptx.py                    # existing visual cleanup tests
```

