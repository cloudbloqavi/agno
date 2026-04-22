# Graph Report - .  (2026-04-22)

## Corpus Check
- 46 files · ~126,648 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 648 nodes · 1839 edges · 25 communities detected
- Extraction: 69% EXTRACTED · 31% INFERRED · 0% AMBIGUOUS · INFERRED: 567 edges (avg confidence: 0.55)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- [[_COMMUNITY_Claude Model Layer|Claude Model Layer]]
- [[_COMMUNITY_Agent Shared Infrastructure|Agent Shared Infrastructure]]
- [[_COMMUNITY_Chunked Workflow Orchestrator|Chunked Workflow Orchestrator]]
- [[_COMMUNITY_Slide Layout Engine|Slide Layout Engine]]
- [[_COMMUNITY_Multi-Provider Agent Modules|Multi-Provider Agent Modules]]
- [[_COMMUNITY_Template Visual Profile Tests|Template Visual Profile Tests]]
- [[_COMMUNITY_Contrast and Diagnostics|Contrast and Diagnostics]]
- [[_COMMUNITY_No-Template Design System|No-Template Design System]]
- [[_COMMUNITY_Content Classification Engine|Content Classification Engine]]
- [[_COMMUNITY_Brand Style Parsing Tests|Brand Style Parsing Tests]]
- [[_COMMUNITY_Template Style Application|Template Style Application]]
- [[_COMMUNITY_Architecture Documentation|Architecture Documentation]]
- [[_COMMUNITY_Assembly Knowledge Pipeline|Assembly Knowledge Pipeline]]
- [[_COMMUNITY_Shape Overlap and Backdrop Safety|Shape Overlap and Backdrop Safety]]
- [[_COMMUNITY_File Download Helper|File Download Helper]]
- [[_COMMUNITY_Theme Factory|Theme Factory]]
- [[_COMMUNITY_XML Sanitization|XML Sanitization]]
- [[_COMMUNITY_Storyboard Planning|Storyboard Planning]]
- [[_COMMUNITY_Token Rate Tracking|Token Rate Tracking]]
- [[_COMMUNITY_Chart and Table Transfer|Chart and Table Transfer]]
- [[_COMMUNITY_Image Planning|Image Planning]]
- [[_COMMUNITY_Test Utilities|Test Utilities]]
- [[_COMMUNITY_PPTX Slide Cleaner|PPTX Slide Cleaner]]
- [[_COMMUNITY_Session State Management|Session State Management]]
- [[_COMMUNITY_Update Docs Utility|Update Docs Utility]]

## God Nodes (most connected - your core abstractions)
1. `Claude` - 264 edges
2. `PythonTools` - 72 edges
3. `StoryboardPlan` - 63 edges
4. `LayoutConstraints` - 61 edges
5. `SlideStoryboard` - 60 edges
6. `_populate_slide()` - 36 edges
7. `step_assemble_template()` - 28 edges
8. `step_optimize_and_plan()` - 21 edges
9. `_apply_visual_corrections()` - 21 edges
10. `generate_chunk_pptx()` - 20 edges

## Surprising Connections (you probably didn't know these)
- `Walkthrough 1` --references--> `Architecture Doc`  [INFERRED]
  Walkthrough1.md → ARCHITECTURE_powerpoint_chunked_workflow.md
- `Walkthrough 2` --references--> `Architecture Doc`  [INFERRED]
  Walkthrough2.md → ARCHITECTURE_powerpoint_chunked_workflow.md
- `Brand/Style-Aware Query Parsing` --conceptually_related_to--> `Theme Factory Skill`  [INFERRED]
  ARCHITECTURE_powerpoint_chunked_workflow.md → theme-factory/SKILL.md
- `Determines if a shape is a structural template backdrop.     Uses persistent nam` --uses--> `Claude`  [INFERRED]
  /mnt/c/Users/aviji/repo/agno/cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/powerpoint_template_workflow.py → /mnt/c/Users/aviji/repo/agno/cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/lib_patches/anthropic/claude.py
- `Mark a shape as a protected backdrop that should never be reflowed.      Adds a` --uses--> `Claude`  [INFERRED]
  /mnt/c/Users/aviji/repo/agno/cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/powerpoint_template_workflow.py → /mnt/c/Users/aviji/repo/agno/cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/lib_patches/anthropic/claude.py

## Hyperedges (group relationships)
- **Template Assembly Pipeline** — architecture_template_assembly, architecture_layout_sanitization, architecture_knowledge_file, design_visual_quality [INFERRED 0.85]
- **Theme Factory Collection** — theme_factory_skill, theme_arctic_frost, theme_midnight_galaxy, theme_ocean_depths, theme_golden_hour, theme_modern_minimalist [EXTRACTED 1.00]
- **Production Reliability System** — architecture_3tier_fallback, architecture_multi_provider, patches_python_tools_traceback [INFERRED 0.80]

## Communities

### Community 0 - "Claude Model Layer"
Cohesion: 0.04
Nodes (69): Claude, _extract_container_id_from_messages(), Parse the Claude streaming response into ModelProviderResponse objects., Parse the given Anthropic-specific usage into an Agno MessageMetrics object., Validate model configuration after initialization, Check if the current model supports native structured outputs.          Returns:, Check if structured outputs are being used in this request.          Args:, Validate that the current model supports extended thinking.          Raises: (+61 more)

### Community 1 - "Agent Shared Infrastructure"
Cohesion: 0.07
Nodes (71): _RateLimitTracker, Agno Workflow: Chunked PowerPoint Generation Pipeline.  A chunked workflow that, Uses the Agno Skills Architecture to either select an existing theme or dynamica, Extract branding/styling intent from the user query via a two-stage approach., Extract branding/styling information from a .pptx template file.      Reads the, Analyze a template's visual layout characteristics programmatically.      Opens, Build a structured log message when template styling overrides query branding., Format a BrandStyleIntent as a markdown section for injection into LLM prompts. (+63 more)

### Community 2 - "Chunked Workflow Orchestrator"
Cohesion: 0.09
Nodes (48): list_themes(), _analyze_template_visual_profile(), BrandStyleIntent, _build_brand_override_log(), build_chunked_workflow(), _build_no_template_design_system(), _build_visual_reference_section(), _countdown_sleep() (+40 more)

### Community 3 - "Slide Layout Engine"
Cohesion: 0.07
Nodes (40): _apply_standard_line_spacing(), _best_text_placeholder(), _best_visual_placeholder(), _compute_max_font_size(), _compute_region_map(), _compute_text_ratio(), ContentArea, _ensure_chart_fills_area() (+32 more)

### Community 4 - "Multi-Provider Agent Modules"
Cohesion: 0.08
Nodes (30): BaseModel, create_agents(), Claude-specific agent definitions for PowerPoint workflow.  Provides the 5+1 swa, Create and return all 5 swappable agents using Claude models.      Brand style a, get_openai_fallback_agents(), Universal OpenAI fallback agent definitions for PowerPoint workflow.  Provides h, Create and return the universal fallback agents using OpenAI models.      These, create_agents() (+22 more)

### Community 5 - "Template Visual Profile Tests"
Cohesion: 0.11
Nodes (35): _analyze_template_visual_profile(), _format_visual_profile_for_prompt(), Offline unit tests for Template Visual Profile analysis logic.  Tests the SlideL, Format a TemplateVisualProfile as a markdown section for the optimizer prompt., Visual characteristics of a single template slide., SlideLayoutProfile should have sensible defaults., TemplateVisualProfile should have sensible defaults., Single-slide minimal template should produce a valid profile. (+27 more)

### Community 6 - "Contrast and Diagnostics"
Cohesion: 0.12
Nodes (30): deep_diagnose(), clean_presentation_visual_noise_and_contrast(), _contrast_ratio(), enforce_final_contrast(), _ensure_text_contrast(), _extract_color_from_solid_fill(), _fix_chart_text_contrast(), _get_shape_background_color() (+22 more)

### Community 7 - "No-Template Design System"
Cohesion: 0.11
Nodes (31): _add_filled_rect(), _add_textbox_styled(), _apply_accent_pattern_to_slide(), _apply_default_design_system(), _apply_density_reduction(), _build_card_grid(), _build_chevron_process(), _build_hero_layout() (+23 more)

### Community 8 - "Content Classification Engine"
Cohesion: 0.08
Nodes (29): Enum, _apply_chart_style(), _classify_slide_semantics(), ContentMix, _download_brand_logo(), _extract_kpi_metrics(), _find_matching_template_layout(), _is_text_placeholder_type() (+21 more)

### Community 9 - "Brand Style Parsing Tests"
Cohesion: 0.14
Nodes (28): BrandStyleIntent, _build_brand_override_log(), extract_style_from_template(), _format_brand_context_for_prompt(), Offline unit tests for Brand/Style parsing logic.  Tests the BrandStyleIntent mo, BrandStyleIntent should have sensible defaults when no branding is present., BrandStyleIntent should accept and store all fields correctly., BrandStyleIntent should survive JSON serialization/deserialization. (+20 more)

### Community 10 - "Template Style Application"
Cohesion: 0.09
Nodes (28): _apply_accent_color_to_body(), _apply_accent_color_to_title(), _apply_body_accent_border(), _apply_visual_corrections(), _enforce_typography_hierarchy(), _ensure_chart_data_labels(), _fix_alignment(), _fix_paragraph_alignment_body() (+20 more)

### Community 11 - "Architecture Documentation"
Cohesion: 0.08
Nodes (29): 3-Tier Fallback Architecture, Brand/Style-Aware Query Parsing, Assembly Knowledge File, Layout Sanitization Pipeline, Multi-Provider Agent Routing, Langfuse Observability Layer, Architecture Doc, Template Assembly Pipeline (+21 more)

### Community 12 - "Assembly Knowledge Pipeline"
Cohesion: 0.08
Nodes (26): _build_assembly_knowledge_file(), _classify_content_mix(), _clear_shape_text_only(), _detect_logo_shapes(), _extract_header_style_from_preserved_shapes(), _find_best_layout(), _get_visual_style_preset(), _inject_content_into_carriers() (+18 more)

### Community 13 - "Shape Overlap and Backdrop Safety"
Cohesion: 0.17
Nodes (12): _clear_unused_placeholders(), _fix_overlapping_shapes(), _is_backdrop(), Sanitize a single slide's shape layout to fix spatial defects.      Performs eig, Determines if a shape is a structural template backdrop.     Uses persistent nam, Return True if shape (or any child in a group) contains visible text.      Uses, Fixes overlapping non-placeholder shapes and enforces slide boundaries.      Alg, Remove unused placeholder XML elements from slide to prevent ghost text.      Th (+4 more)

### Community 14 - "File Download Helper"
Cohesion: 0.25
Nodes (7): detect_file_extension(), download_skill_files(), Download files created by Claude Agent Skills from the API response.      Args, Detect file type from magic bytes (file header).      Args:         file_cont, main(), Run an isolated test of the Claude Content Generator (Tier 1)., test_content_agent()

### Community 15 - "Theme Factory"
Cohesion: 0.2
Nodes (10): ChartExtract, _extract_slide_content(), ImageData, Extracted image data with position., Extracted chart data with position., All extracted content from a single slide., Extract all content from a slide including text, tables, images, charts, and sha, Extracted table data with position. (+2 more)

### Community 16 - "XML Sanitization"
Cohesion: 0.25
Nodes (8): _analyze_template_in_depth(), _extract_shape_design_info(), _is_picture_placeholder(), _layout_richness_score(), Check if a shape is a picture placeholder using multiple strategies., Heuristic for visually rich layouts (cards, shapes, picture slots)., Extract design properties from a non-placeholder shape (decorative element)., Perform a thorough per-layout analysis of the template's complete design languag

### Community 17 - "Storyboard Planning"
Cohesion: 0.33
Nodes (6): _add_picture_within_bounds(), _fit_to_area(), Scale dimensions to fit within an area while preserving aspect ratio.      Args:, Insert image scaled to fit bounds while preserving aspect ratio., Transfer extracted images to a slide, scaled to fit the content area., _transfer_images()

### Community 18 - "Token Rate Tracking"
Cohesion: 0.33
Nodes (6): _apply_template_font_to_shape_xml(), Rescale a shape element's position/size from source slide dimensions     to targ, Apply a template font family to all text runs inside a raw shape XML element., Transfer simple shapes by deep-copying their XML to the target slide.      When, _rescale_shape_xml(), _transfer_shapes()

### Community 19 - "Chart and Table Transfer"
Cohesion: 0.33
Nodes (6): _classify_template_shape(), _extract_shape_fill_hex(), _get_shape_text_content(), Extract all visible text from a shape (including nested groups).      Returns th, Extract the solid fill color hex from a shape, or empty string.      Inspects ``, Classify a template shape as 'structural', 'content_carrier', or 'disposable'.

### Community 20 - "Image Planning"
Cohesion: 0.5
Nodes (4): _apply_table_style(), Apply template-derived styling to a newly created table., Transfer extracted table data to a slide, repositioned to the content area., _transfer_tables()

### Community 22 - "Test Utilities"
Cohesion: 1.0
Nodes (1): Extract the most recent container ID from message provider_data.          Reads

### Community 30 - "PPTX Slide Cleaner"
Cohesion: 1.0
Nodes (1): Product Requirements

### Community 31 - "Session State Management"
Cohesion: 1.0
Nodes (1): README

### Community 32 - "Update Docs Utility"
Cohesion: 1.0
Nodes (1): Scratchpad Notes

## Knowledge Gaps
- **89 isolated node(s):** `Detect file type from magic bytes (file header).      Args:         file_cont`, `Download files created by Claude Agent Skills from the API response.      Args`, `Accept design_score as alias for overall_quality and vice-versa.          Vision`, `Parsed branding/styling intent extracted from a user query or template file.`, `Decision about whether a slide needs an AI-generated image.` (+84 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **Thin community `Test Utilities`** (1 nodes): `Extract the most recent container ID from message provider_data.          Reads`
  Too small to be a meaningful cluster - may be noise or needs more connections extracted.
- **Thin community `PPTX Slide Cleaner`** (1 nodes): `Product Requirements`
  Too small to be a meaningful cluster - may be noise or needs more connections extracted.
- **Thin community `Session State Management`** (1 nodes): `README`
  Too small to be a meaningful cluster - may be noise or needs more connections extracted.
- **Thin community `Update Docs Utility`** (1 nodes): `Scratchpad Notes`
  Too small to be a meaningful cluster - may be noise or needs more connections extracted.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `Claude` connect `Claude Model Layer` to `Agent Shared Infrastructure`, `Chunked Workflow Orchestrator`, `Slide Layout Engine`, `Multi-Provider Agent Modules`, `Contrast and Diagnostics`, `No-Template Design System`, `Content Classification Engine`, `Template Style Application`, `Assembly Knowledge Pipeline`, `Shape Overlap and Backdrop Safety`, `File Download Helper`, `Theme Factory`, `XML Sanitization`, `Storyboard Planning`, `Token Rate Tracking`, `Chart and Table Transfer`, `Image Planning`?**
  _High betweenness centrality (0.483) - this node is a cross-community bridge._
- **Why does `_analyze_template_visual_profile()` connect `Template Visual Profile Tests` to `Chunked Workflow Orchestrator`?**
  _High betweenness centrality (0.099) - this node is a cross-community bridge._
- **Why does `PythonTools` connect `Agent Shared Infrastructure` to `Chunked Workflow Orchestrator`, `Multi-Provider Agent Modules`?**
  _High betweenness centrality (0.070) - this node is a cross-community bridge._
- **Are the 233 inferred relationships involving `Claude` (e.g. with `_RateLimitTracker` and `SlideLayoutProfile`) actually correct?**
  _`Claude` has 233 INFERRED edges - model-reasoned connections that need verification._
- **Are the 25 inferred relationships involving `str` (e.g. with `_render_template_slides_to_png()` and `_build_visual_reference_section()`) actually correct?**
  _`str` has 25 INFERRED edges - model-reasoned connections that need verification._
- **Are the 61 inferred relationships involving `PythonTools` (e.g. with `_RateLimitTracker` and `SlideLayoutProfile`) actually correct?**
  _`PythonTools` has 61 INFERRED edges - model-reasoned connections that need verification._
- **Are the 60 inferred relationships involving `StoryboardPlan` (e.g. with `_RateLimitTracker` and `SlideLayoutProfile`) actually correct?**
  _`StoryboardPlan` has 60 INFERRED edges - model-reasoned connections that need verification._