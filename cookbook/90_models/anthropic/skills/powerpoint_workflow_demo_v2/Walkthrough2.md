# SCRATCHPAD Concerns — Walkthrough

## Changes Made

### Concern 3: Agent/Provider Traceability ✅
- Added `_log_agent_banner()` function that prints a formatted banner with agent name, model ID, provider, and step context
- Inserted banner calls at **all 7** agent invocation sites (brand parser, query optimizer, Tier 1 chunk generator, Tier 2 primary + lite, image planner, visual reviewer)
- Always-on — visible without `--verbose`

render_diffs(file:///mnt/c/Users/aviji/repo/agno/cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/powerpoint_chunked_workflow.py)

---

### Concern 5: Rate Limit Fix ✅
- **Retry delay**: Changed from `1s/2s` exponential to `60-90s` with jitter + countdown logging
- **Default chunk size**: Reduced from `3` → `2` (CLI flag + docstring + help text)
- **New test**: [test_content_agent_isolated.py](file:///mnt/c/Users/aviji/repo/agno/cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/tests/test_content_agent_isolated.py) — standalone Tier 1 agent test (requires API key)

---

### Concern 4: Tier 2 Model Upgrade ✅
- Upgraded `fallback_code_agent` from `claude-haiku-4-5` → `claude-sonnet-4-6` in all 3 provider modules
- Added `fallback_code_agent_lite` (haiku/mini/flash) as cheaper retry across all providers
- Implemented **sonnet → haiku → Tier 3** chain in `generate_chunk_pptx_v2()` with full banner + rate-tracking at each stage

render_diffs(file:///mnt/c/Users/aviji/repo/agno/cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/agents/claude_agents.py)

---

### Concern 2: Layout Sanitization ✅
- Added `sanitize_slide_layout()` with 3-pass algorithm:
  1. **Boundary clamping** (5% safe margin)
  2. **Minimum size enforcement** (1" × 0.5")
  3. **Overlap detection & vertical reflow** (>30% overlap threshold)
- Added `sanitize_presentation()` wrapper
- Integrated into all 3 tier generators (after `clean_presentation_visual_noise_and_contrast()`)

render_diffs(file:///mnt/c/Users/aviji/repo/agno/cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2/powerpoint_template_workflow.py)

---

### Concern 1+6: Template Visual References ✅
- `_render_template_slides_to_png()`: Uses LibreOffice→PDF→PNG pipeline at 100 DPI
- `_match_storyboard_to_template_slide()`: Type-based mapping (title→slide 0, content→slide 1, data→slide 3, closing→last)
- `_build_visual_reference_section()`: Base64-encodes matched PNGs into markdown prompt section
- Integrated into **Step 1** (auto-render on template load) and both **Tier 1 + Tier 2** prompts
- Graceful degradation: if LibreOffice/pdftoppm unavailable, continues without visual references

---

## Files Modified

| File | Changes |
|------|---------|
| `powerpoint_chunked_workflow.py` | Agent banners, retry fix, chunk size, visual reference infra, sanitize integration |
| `powerpoint_template_workflow.py` | `sanitize_slide_layout()`, `sanitize_presentation()` |
| `agents/claude_agents.py` | sonnet upgrade, `fallback_code_agent_lite` |
| `agents/openai_agents.py` | `fallback_code_agent_lite` (gpt-5-mini) |
| `agents/gemini_agents.py` | `fallback_code_agent_lite` (flash) |
| `agents/__init__.py` | New role in `AGENT_ROLES` + validation |
| `tests/test_content_agent_isolated.py` | **[NEW]** Standalone Tier 1 test |

## Verification

| Check | Result |
|-------|--------|
| Syntax compilation (7 files) | ✅ All pass |
| Existing tests (10 brand/style tests) | ✅ All pass, 0 regressions |

## Next Steps for Manual Testing

```bash
# Full E2E run with template + verbose (validates all concerns)
python powerpoint_chunked_workflow.py \
  -t templates/your_template.pptx \
  -p "Create a 6-slide AI healthcare deck with data charts" \
  --verbose --chunk-size 2
```

Check `OUTPUT.md` for: agent banners, retry delays ≥60s, sonnet in Tier 2, layout sanitize logs, and visual reference injection logs.
