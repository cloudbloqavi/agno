# AGENTS.md — Chunked PowerPoint Workflow v2

## Agent Architecture

Sequential specialist pattern — agents never communicate directly. Each handles one phase,
sharing context through `session_state` and Pydantic-validated structured outputs.

```
step_optimize_and_plan()
  ├── brand_style_analyzer  (Sonnet + web_search)  → BrandStyleIntent
  └── query_optimizer       (Opus + web_search)    → StoryboardPlan

step_generate_chunks()
  ├── Claude PPTX Skill Agent  (Tier 1)  → raw .pptx
  ├── fallback_code_agent      (Tier 2)  → code-gen .pptx
  └── (deterministic fallback) (Tier 3)  → text-only .pptx

step_plan_images()
  └── image_planner            (Gemini Flash)  → ImagePlan

step_visual_quality_review()
  └── slide_quality_reviewer   (Gemini 2.5 Flash)  → SlideQualityReport
```

---

## Agent Catalog

### 1. Brand Style Analyzer

| Property | Value |
|----------|-------|
| **File** | `agents/` package (e.g. `claude_agents.py`, `openai_agents.py`, `gemini_agents.py`) |
| **Variable** | `brand_style_analyzer` |
| **Model** | Swappable: Claude Sonnet (`claude-sonnet-4-6`), OpenAI (`gpt-5-mini`), Gemini (`gemini-3-flash-preview`) |
| **Tools** | Provider-native web search (`web_search_20250305`, `web_search_preview`, or `search=True`) |
| **Output Schema** | `BrandStyleIntent` (Pydantic) |
| **Purpose** | Detects brand/style directives in the user prompt (e.g. "using Nike branding", "in the style of Apple"). Autonomously decides whether to search for brand guidelines online. Returns structured intent: brand name, color palette, tone, typography hints, style keywords. |
| **Trigger** | Start of `step_optimize_and_plan()`, before the query optimizer |
| **Design Choice** | Claude Sonnet (fast, cheap) — lightweight analysis, not content generation |

**Output fields:**
- `has_branding` — True when a brand directive is detected
- `brand_name` — Extracted brand name (e.g. "Nike")
- `color_palette` — Hex codes (e.g. `["#FF6600", "#000000"]`)
- `tone_override` — Inferred tone (e.g. "empowering", "innovative")
- `typography_hints` — Font families (e.g. `["Futura", "Helvetica Neue"]`)
- `style_keywords` — 3-5 descriptors (e.g. `["bold", "sporty", "dynamic"]`)
- `content_query` — User query with branding clause removed
- `source` / `source_detail` — "query" or "template" with provenance

### 2. Query Optimizer

| Property | Value |
|----------|-------|
| **File** | `agents/` package |
| **Variable** | `query_optimizer` |
| **Model** | Swappable: Claude Opus (`claude-opus-4-6`), OpenAI (`gpt-5.2`), Gemini (`gemini-3-pro-preview`) |
| **Tools** | Provider-native web search |
| **Output** | `StoryboardPlan` (via prompt instructions + manual JSON parse) |
| **Purpose** | Takes the user prompt + brand context and produces a researched, structured storyboard that guides all downstream chunk generation. Plans slide count, narrative flow, tone, brand voice, and per-slide content outline. |
| **Note** | `output_schema` is intentionally omitted — `claude-opus-4-6` does not support structured outputs, which causes Agno to make an internal non-streaming extraction call that the `context-1m` beta rejects. Storyboard JSON is requested via prompt instructions and parsed manually. |

**Output fields:**
- `total_slides` — Planned count (respects user-specified count, else 8-15)
- `presentation_title` — Main title
- `search_topic` — Primary research phrase for web search
- `target_audience` — Intended audience
- `tone` / `brand_voice` — Presentation tone and brand voice style
- `global_context` — 2-3 sentence shared context
- `slides` — List of `SlideStoryboard` (title, type, key points, visual suggestion, transition note)

### 3. Content Generation Agent (Tier 1)

| Property | Value |
|----------|-------|
| **File** | `powerpoint_template_workflow.py` (single-call) / `powerpoint_chunked_workflow.py` (per-chunk) |
| **Model** | Claude Opus `claude-opus-4-6` + `context-1m-2025-08-07` beta |
| **Skills** | `pptx` (Anthropic Agent Skill — native PowerPoint creation in sandbox) |
| **Output** | Raw `.pptx` file (downloaded via Anthropic Files API + `file_download_helper.py`) |
| **Purpose** | Primary content generator. Writes slide content within a sandboxed environment. Produces titles, bullets, tables, and charts. In chunked mode, generates N-slide chunks guided by storyboard context + brand context. |
| **Timeout** | 300s per attempt (`CHUNK_TIMEOUT_SECONDS`); retries up to `max_retries` |
| **Brand injection** | Brand context appended as `## Brand/Style Guidance` section in chunk prompt |

### 4. Fallback Code Agent (Tier 2)

| Property | Value |
|----------|-------|
| **File** | `agents/` package |
| **Variable** | `fallback_code_agent` |
| **Model** | Swappable: Claude Opus, OpenAI `gpt-5.2`, Gemini `gemini-3-pro-preview` |
| **Tools** | `PythonTools` (code execution — `save_and_run_python_code`) |
| **Output** | `.pptx` file generated via python-pptx + matplotlib code |
| **Purpose** | Fallback when Tier 1 fails. Generates and immediately executes a Python script that builds slides with real Office charts (`ChartData`), matplotlib PNG embeds, and tables. No internal retry — escalates to Tier 3 on failure. |
| **Quality** | 80–92% of Tier 1 quality |
| **Brand injection** | Brand context appended to `GLOBAL CONTEXT` in code-gen prompt |

### 5. Image Planner

| Property | Value |
|----------|-------|
| **File** | `agents/` package |
| **Variable** | `image_planner` |
| **Model** | Swappable: Gemini `gemini-3-flash-preview`, OpenAI `gpt-5-mini` |
| **Output Schema** | `ImagePlan` → list of `SlideImageDecision` (Pydantic structured output) |
| **Purpose** | Reviews each slide's content summary and decides which slides benefit from AI-generated images. Outputs per-slide yes/no + image prompt + reasoning. |
| **Decision rules** | Title slides: usually YES · Data slides (tables/charts): usually NO · Slides with existing images: ALWAYS NO · Closing slides: usually NO |

### 6. Slide Quality Reviewer

| Property | Value |
|----------|-------|
| **File** | `agents/` package |
| **Variable** | `slide_quality_reviewer` |
| **Model** | Swappable: Gemini `gemini-2.5-flash`, OpenAI `gpt-5-mini` (vision) |
| **Output Schema** | `SlideQualityReport` → per-slide assessment with `ShapeIssue` list |
| **Purpose** | Inspects rendered slide PNGs for visual defects. Reports issues with severity (critical/moderate/minor) and suggests programmatic fixes. |
| **Correction scope** | Auto-fixes: `low_contrast`, `ghost_text`, `empty_placeholder`, `text_overflow` · Detect-only: visual blandness · Deferred: `overlap` repositioning |
| **Prerequisite** | LibreOffice headless (for PNG rendering). Fully non-blocking — skips gracefully if unavailable. |

---

## Agent Communication Protocol

Agents communicate **indirectly** via `session_state`, a shared dictionary:

| Producer | Consumer | Channel |
|----------|----------|---------|
| `brand_style_analyzer` | `query_optimizer` | `session_state["brand_style_intent"]` → `BrandStyleIntent` |
| `query_optimizer` | `step_generate_chunks` | `session_state["storyboard"]` → `StoryboardPlan` + storyboard `.md` files |
| Content Agent (Tier 1-3) | `step_process_chunks` | `session_state["chunk_files"]` → list of `.pptx` paths |
| `image_planner` | `step_generate_images` | `session_state["slides_data"]` enriched with image plan |
| `slide_quality_reviewer` | `step_merge_chunks` | `session_state["reviewed_chunks"]` → dict of `.pptx` paths |

---

## AI Skills & Tools

| Agent | Skill / Tool | Provider | Purpose |
|-------|-------------|----------|---------|
| Content Agent (Tier 1) | `pptx` skill | Anthropic | Native PowerPoint creation in sandbox |
| Fallback Agent (Tier 2) | `PythonTools` | Swappable | Execute generated python-pptx + matplotlib |
| Brand Analyzer | `web_search` or `search=True` | Swappable | Research brand guidelines (max 2 uses) |
| Query Optimizer | `web_search` or `search=True` | Swappable | Research topic for storyboard (max 5 uses) |
| Image Planner | (structured output only) | Swappable | No tools — uses Pydantic output schema |
| Quality Reviewer | (vision input only) | Swappable | No tools — processes PNG images |

---

## Deterministic Functions (No AI)

These functions execute without any LLM call:

| Function | Capability |
|----------|-----------|
| `_extract_slide_content()` | Shape classification and content extraction |
| `_find_best_layout()` | Template layout scoring and selection |
| `_compute_region_map()` | Text/visual region splitting |
| `_populate_slide()` | Full slide assembly orchestration |
| `_transfer_tables/images/charts/shapes` | Content positioning within ContentArea |
| `_ensure_text_contrast()` | WCAG-based contrast correction |
| `_fit_to_area()` | Aspect-ratio-preserving image scaling |
| `_analyze_template_in_depth()` | Deep per-layout design language extraction |
| `_build_assembly_knowledge_file()` | Comprehensive knowledge file consolidation |
| `extract_style_from_template()` | Theme XML parsing for brand override |
| `_merge_pptx_zip_level()` | Multi-chunk PPTX merging with OPC relationship management |
| `generate_chunk_pptx_fallback()` | Tier 3 text-only slide generation |

---

## Error Handling & Resilience

| Strategy | Implementation |
|----------|---------------|
| **3-Tier Fallback** | Tier 1 → Tier 2 → Tier 3 per chunk; session-level flag bypasses Tier 1 after first failure |
| **Per-attempt Timeout** | 300s via `ThreadPoolExecutor.result(timeout=...)` |
| **Configurable Retries** | `--max-retries` per chunk (default: 2) with exponential backoff |
| **Start Tier Override** | `--start-tier 2` skips Tier 1 entirely (useful when skill is known to be down) |
| **Graceful Degradation** | Visual review is fully non-blocking; missing LibreOffice or API errors return success with warning |
| **Missing API Keys** | Image generation gracefully skipped; pipeline continues |
| **Interim File Preservation** | All chunk files preserved in `output_chunked/` for debugging (no auto-cleanup) |
| **Output Redirect** | All stdout/stderr + Python logging captured to `OUTPUT.md` for post-mortem analysis |
