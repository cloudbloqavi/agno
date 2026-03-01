# AGENTS.md — PowerPoint Workflow Suite

## Agent Architecture

The pipeline employs a **sequential specialist** pattern rather than a hierarchical manager/worker model. Each agent handles one distinct phase of the pipeline. Agents never communicate directly; they share context through `session_state` and structured Pydantic outputs.

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

| Property         | Value                                                                  |
| ---------------- | ---------------------------------------------------------------------- |
| **File**         | `powerpoint_chunked_workflow.py`                                       |
| **Model**        | Claude Sonnet (`claude-sonnet-4-20250514`)                                 |
| **Tools**        | `web_search` (max 2 invocations)                                       |
| **Output**       | `BrandStyleIntent` (Pydantic structured output)                       |
| **Purpose**      | Detects brand/style directives in the user prompt. Autonomously decides whether to search for brand guidelines online. Returns structured intent: brand name, color palette, tone, typography hints, style keywords. |
| **Trigger**      | Start of `step_optimize_and_plan()`, before the query optimizer        |
| **Design choice**| Claude Sonnet (fast, cheap) — this is a lightweight analysis task, not content generation |

### 2. Query Optimizer

| Property         | Value                                                                  |
| ---------------- | ---------------------------------------------------------------------- |
| **File**         | `powerpoint_chunked_workflow.py`                                       |
| **Model**        | Claude Opus (`claude-opus-4-6`) + `context-1m-2025-08-07` beta           |
| **Tools**        | `web_search` (max 5 invocations)                                       |
| **Output**       | `StoryboardPlan` (Pydantic: total_slides, target_audience, narrative_arc, global_context, per-slide storyboards) |
| **Purpose**      | Takes the user prompt + brand context and produces a researched, structured storyboard that guides all downstream chunk generation. Plans slide count, narrative flow, and per-slide content outline. |
| **Injection**    | Brand context from `brand_style_analyzer` is prepended to the optimizer prompt |

### 3. Content Generation Agent (Tier 1)

| Property         | Value                                                                  |
| ---------------- | ---------------------------------------------------------------------- |
| **File**         | `powerpoint_template_workflow.py` / `powerpoint_chunked_workflow.py`   |
| **Model**        | Claude Opus (`claude-opus-4-6`) + `context-1m-2025-08-07` beta           |
| **Skills**       | `pptx` (Anthropic Agent Skill — native PowerPoint creation)            |
| **Output**       | Raw `.pptx` file (downloaded via Anthropic Files API)                  |
| **Purpose**      | Primary content generator. Writes slide content within a sandboxed environment. Produces titles, bullets, tables, and charts. In chunked mode, generates N-slide chunks guided by storyboard context. |
| **Timeout**      | 300s per attempt (`CHUNK_TIMEOUT_SECONDS`); retries up to `max_retries` |

### 4. Fallback Code Agent (Tier 2)

| Property         | Value                                                                  |
| ---------------- | ---------------------------------------------------------------------- |
| **File**         | `powerpoint_chunked_workflow.py`                                       |
| **Model**        | Claude Opus (`claude-opus-4-6`) (without `context-1m` beta)              |
| **Tools**        | `PythonTools` (code execution)                                         |
| **Output**       | `.pptx` file generated via python-pptx + matplotlib code              |
| **Purpose**      | Fallback when Tier 1 fails. Generates and immediately executes a Python script that builds slides with real Office charts, matplotlib PNG embeds, and tables. No internal retry — escalates to Tier 3 on failure. |
| **Quality**      | 80–92% of Tier 1 quality                                              |

### 5. Image Planner

| Property         | Value                                                                  |
| ---------------- | ---------------------------------------------------------------------- |
| **File**         | `powerpoint_template_workflow.py`                                      |
| **Model**        | Gemini `gemini-3-flash-preview` (Google)                              |
| **Output**       | `ImagePlan` → list of `SlideImageDecision` (Pydantic structured output)|
| **Purpose**      | Reviews each slide's content summary and decides which slides benefit from AI-generated images. Outputs per-slide yes/no + image prompt + reasoning. |
| **Decision rules**| Title slides: usually YES · Data slides with tables/charts: usually NO · Slides with existing images: ALWAYS NO · Closing slides: usually NO |

### 6. Slide Quality Reviewer

| Property         | Value                                                                  |
| ---------------- | ---------------------------------------------------------------------- |
| **File**         | `powerpoint_template_workflow.py`                                      |
| **Model**        | Gemini 2.5 Flash (vision)                                             |
| **Output**       | `SlideQualityReport` → per-slide quality assessment with `ShapeIssue` list |
| **Purpose**      | Inspects rendered slide PNGs for visual defects. Reports issues with severity (critical/major/moderate/minor) and suggests programmatic fixes. |
| **Correction scope** | Auto-fixes: `low_contrast`, `ghost_text`, `empty_placeholder`, `text_overflow` · Detect-only: visual blandness · Deferred: `overlap` repositioning |

---

## Agent Communication

Agents communicate **indirectly** via `session_state`, a shared dictionary passed through all workflow steps.

### Data Flow Protocol

| From Agent               | To Step                   | Channel                                         |
| ------------------------ | ------------------------- | ------------------------------------------------ |
| `brand_style_analyzer`   | `query_optimizer`         | `session_state["brand_style_intent"]` (BrandStyleIntent) |
| `query_optimizer`        | `step_generate_chunks`    | `session_state["storyboard"]` (StoryboardPlan) + storyboard markdown files |
| Content Agent (Tier 1-3) | `step_process_chunks`     | `session_state["chunk_files"]` (list of .pptx paths) |
| `image_planner`          | `step_generate_images`    | `session_state["slides_data"]` enriched with image plan |
| `slide_quality_reviewer` | `step_merge_chunks`       | `session_state["reviewed_chunks"]` (dict of .pptx paths) |

### Structured Output Schemas

All AI agents produce **Pydantic-validated structured output** — no free-form text parsing:

- `BrandStyleIntent` → brand name, colors, tone, fonts, style keywords, source
- `StoryboardPlan` → total slides, audience, narrative arc, global context, per-slide storyboards
- `ImagePlan` → per-slide image decision, prompt, reasoning
- `SlideQualityReport` → overall quality score, blandness flag, list of `ShapeIssue` items

---

## Agent Skills & Tools

### AI Skills

| Agent                  | Skill / Tool          | Purpose                                    |
| ---------------------- | --------------------- | ------------------------------------------ |
| Content Agent (Tier 1) | `pptx` (Anthropic)    | Native PowerPoint creation in sandbox      |
| Fallback Agent (Tier 2)| `PythonTools` (Agno)  | Execute generated python-pptx + matplotlib |
| Brand Analyzer         | `web_search`          | Research brand guidelines (max 2 uses)     |
| Query Optimizer        | `web_search`          | Research topic for storyboard (max 5 uses) |

### Deterministic Skills (No AI)

| Function                                | Capability                                        |
| --------------------------------------- | ------------------------------------------------- |
| `_extract_slide_content()`              | Shape classification and content extraction        |
| `_find_best_layout()`                   | Template layout scoring and selection              |
| `_compute_region_map()`                 | Text/visual region splitting                       |
| `_populate_slide()`                     | Full slide assembly orchestration                  |
| `_transfer_tables/images/charts/shapes` | Content positioning within ContentArea             |
| `_ensure_text_contrast()`               | WCAG-based contrast correction                     |
| `_fit_to_area()`                        | Aspect-ratio-preserving image scaling              |
| `_analyze_template_in_depth()`          | Deep per-layout design language extraction          |
| `_build_assembly_knowledge_file()`      | Comprehensive knowledge file consolidation          |
| `extract_style_from_template()`         | Theme XML parsing for brand override               |
| `merge_pptx_files()` / `_clone_slide()` | Multi-chunk PPTX merging                          |

---

## Error Handling & Resilience

| Strategy                    | Implementation                                              |
| --------------------------- | ----------------------------------------------------------- |
| **3-Tier Fallback**         | Tier 1 → Tier 2 → Tier 3 per chunk; session-level flag bypasses Tier 1 after first failure |
| **Per-attempt Timeout**     | 300s via `ThreadPoolExecutor.result(timeout=...)`           |
| **Exponential Backoff**     | Retries with increasing delay on Tier 1 failures            |
| **Graceful Degradation**    | Step 5 is fully non-blocking; missing LibreOffice/API errors return success with warning |
| **Missing API Keys**        | Image generation gracefully skipped; pipeline continues     |
| **Interim File Preservation** | All chunk files preserved for debugging (no auto-cleanup) |
