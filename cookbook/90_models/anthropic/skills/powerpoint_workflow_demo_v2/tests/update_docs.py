import re

design_doc = "DESIGN_visual_quality.md"
with open(design_doc, "r") as f:
    design_content = f.read()

# Update Table
table_old = """| Template-aware LLM prompts | `generate_chunk_pptx_v2()` | LLM unaware of template constraints | ✅ Implemented |"""
table_new = """| Template-aware LLM prompts | `generate_chunk_pptx_v2()` | LLM unaware of template constraints | ✅ Implemented |
| OOXML Sequence Strictness | `_make_high_contrast_fill()` | PowerPoint rejected contrast changes | ✅ Implemented |
| Duck-Typing Review Payload | `step_visual_quality_review()` | Pydantic namespace mismatches | ✅ Implemented |
| Chart Scaling Bounds | `enforce_final_contrast()` | Microscopic 0.4" charts | ✅ Implemented |
| Chart Overlap Repositioning | `enforce_final_contrast()` | Adjusted charts stacked over each other | ✅ Implemented |
"""

design_content = design_content.replace(table_old, table_new)

# Add new sections to DESIGN
new_sections = """
---

#### Fix 6: Strict OOXML Schema Application for Text Colors — Technical Deep-Dive

**Root cause:** While the Python logic correctly computed the need for high-contrast white text (`#FFFFFF`) on dark backgrounds, PowerPoint silently rejected the modification. This is because PowerPoint strictly adheres to open XML schemas which demand that the color declaration `<a:solidFill>` mathematically precedes structural font modifiers like `<a:latin>` or `<a:ea>` inside the `<a:rPr>` text run properties element. Appending the color element to the end of the property tree effectively corrupts the container from PowerPoint's renderer perspective, causing it to fall back to the default `#000000`.

**Solution — Positional insertion algorithm:**

```python
insert_idx = 0
tags_after_fill = {"latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"}
for i, child in enumerate(rPr):
    tag = child.tag.split("}")[-1]
    if tag in tags_after_fill:
        insert_idx = i
        break
    insert_idx = i + 1

rPr.insert(insert_idx, solid_fill)
```
**Test result:** Previously black placeholders on Slide 2/3/5 now accurately render white.

---

#### Fix 7: Multi-Layer Background Detection (Layer 3.5)

**Root cause:** Custom templates like `Career-Path-Template` sometimes construct a full-bleed dark background not using the literal `<p:bg>` XML property, but instead using an enormous wrapper shape set to 120% slide bounds located exclusively inside `SlideMaster.shapes`. The background heuristic analyzer did not natively inspect shapes located at the SlideMaster level, thus assuming the page was white.

**Solution — Layer 3.5 Master Shape Traversal:**
Created a sub-routine in `_get_shape_background_color` which calculates the area of all Master slide shapes. If an arbitrary element encompasses >80% of the `src_slide_width` × `src_slide_height` bounds, it intercepts the background scan and utilizes its fill instead.

---

#### Fix 8: Agent Response Duck-Typing

**Root cause:** When reviewing visual elements via `--visual-review`, `SlideQualityReport` is generated at runtime dynamically. The Python `isinstance()` logic was failing the schema check by comparing the `SlideQualityReport` located inside the LLM memory namespace versus the one declared in `__main__`.

**Solution:**
Abandoned stringent memory references in favor of duck-typing schema field checks (e.g. verifying `hasattr("issues")` and `hasattr("overall_quality")`) allowing the QA loop to ingest the payloads flawlessly.
"""

if "Fix 6: Strict OOXML Schema" not in design_content:
    design_content = design_content.replace("---

#### New Constants", new_sections + "\n---\n\n#### New Constants")

with open(design_doc, "w") as f:
    f.write(design_content)

arch_doc = "ARCHITECTURE_powerpoint_chunked_workflow.md"
with open(arch_doc, "r") as f:
    arch_content = f.read()

arch_table_old = """| Per-slide rendering | QA (detective) | `_render_pptx_to_images()` | During `--visual-review` step |"""
arch_table_new = """| Per-slide rendering | QA (detective) | `_render_pptx_to_images()` | During `--visual-review` step |
| OOXML Indexing strictness | Assembly (corrective) | `_make_high_contrast_fill()` | Resolves invisible structural color edits |
| Duck-Typing Pydantic | System (preventative) | `step_visual_quality_review()` | Resolves framework-level module loading failures |
| Spatial Overflow Offset | Assembly (corrective) | `enforce_final_contrast()` | Staggers sequentially upscaled bounding boxes |"""

arch_content = arch_content.replace(arch_table_old, arch_table_new)

with open(arch_doc, "w") as f:
    f.write(arch_content)

print("Updated MD documents.")
