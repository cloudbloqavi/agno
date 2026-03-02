# Library Patches for PowerPoint Workflow Demo v2

This directory contains copies of `agno` library files that were **modified by cloudbloqavi** as part of the PowerPoint workflow demo development. These modifications are tracked separately for portability and traceability.

## Patched Files

### 1. `anthropic/claude.py`

**Original location:** `libs/agno/agno/models/anthropic/claude.py`
**Commit:** `ca573b1e` (2026-03-02)
**Commit message:** `fix: Update Claude structured output detection to correctly identify newer sub-versions using a date-detection heuristic and add comprehensive unit tests.`

**What changed:** The `_supports_structured_outputs` property method.

**Why:** The previous pattern-based exclusion logic (`startswith("claude-sonnet-4-")` and `not startswith("claude-sonnet-4-5")`) incorrectly blocked newer models like `claude-sonnet-4-6` that **do** support structured outputs. The fix uses a date-detection heuristic: only 8-digit dated suffixes starting with `"20"` (e.g., `claude-sonnet-4-20250514`) are excluded, while short numeric sub-versions (e.g., `claude-sonnet-4-6`) pass through as supported.

**Impact on workflow:** Without this fix, the `brand_style_analyzer` agent (which uses `claude-sonnet-4-6` with `output_schema=BrandStyleIntent`) fails to use structured outputs.

```diff
-        if self.id.startswith("claude-sonnet-4-") and not self.id.startswith("claude-sonnet-4-5"):
-            return False
-        if self.id.startswith("claude-opus-4-") and not (
-            self.id.startswith("claude-opus-4-1") or self.id.startswith("claude-opus-4-5")
-        ):
-            return False
+        for prefix in ("claude-sonnet-4-", "claude-opus-4-"):
+            if self.id.startswith(prefix):
+                suffix = self.id[len(prefix):]
+                if suffix and len(suffix) >= 8 and suffix[:8].isdigit() and suffix[:2] == "20":
+                    return False
```

---

### 2. `tests/test_structured_output_support.py`

**Original location:** `libs/agno/tests/unit/models/anthropic/test_structured_output_support.py`
**Commit:** `ca573b1e` (2026-03-02)
**Status:** New file (not a modification)

**What it does:** Comprehensive unit tests for the `_supports_structured_outputs` method, covering:
- Blacklisted models
- Legacy `claude-3-*` family
- Dated versions of `claude-sonnet-4` and `claude-opus-4`
- Newer sub-versions like `claude-sonnet-4-6`

---

### 3. `tools_python.py`

**Original location:** `libs/agno/agno/tools/python.py`
**Commit:** `5b2994aa9` (2026-02-28)
**Commit message:** `refactor: enhance error logging in PythonTools for better traceability and debugging`

**What changed:** The `save_and_run_python_code` method's exception handler.

**Why:** The original error handler only logged the exception message, making it difficult to trace the root cause of failures during Tier 2 (LLM code generation) fallback. The fix adds full traceback output to both the logger and the return value.

**Impact on workflow:** Tier 2 fallback in `powerpoint_chunked_workflow.py` uses `PythonTools` to execute LLM-generated code. Without this fix, debugging Tier 2 failures is extremely difficult because only the exception message (not the traceback) is returned to the LLM agent.

```diff
 import functools
 import runpy
+import traceback
 from pathlib import Path

         except Exception as e:
-            logger.error(f"Error saving and running code: {e}")
-            return f"Error saving and running code: {e}"
+            tb_str = traceback.format_exc()
+            logger.error(f"Error saving and running code: {e}\n{tb_str}")
+            return f"Error saving and running code: {e}\nTraceback:\n{tb_str}"
```

---

## Predecessor Files

The workflow scripts evolved through several early versions at `cookbook/90_models/anthropic/skills/` (one level up, outside the demo folder) before being consolidated into `powerpoint_workflow_demo/`. These predecessor files **no longer exist on disk** — they were moved into the demo folder. Relevant commits:

| Commit | Files | Description |
|--------|-------|-------------|
| `03bb0b14c` | `skills/file_download_helper.py` | Initial file download helper |
| `ef2040899` | `skills/agent_with_powerpoint_template.py`, `skills/powerpoint_template_workflow.py` | Early workflow + agent |
| `3f56717c6` | `skills/powerpoint_template_workflow.py` | Image placeholder detection |
| `0a4e2c731` | `skills/agent_with_powerpoint_template.py`, `skills/file_download_helper.py`, `skills/powerpoint_template_workflow.py` | Error handling + early claude.py change |
| `e33f5a094` | `skills/powerpoint_template_workflow.py` | Collision resolution |

These are fully superseded by the current files in `powerpoint_workflow_demo/` and `powerpoint_workflow_demo_v2/`.

---

## How to Apply

These patches are already applied in the installed `agno` library within this workspace. The files here serve as **documentation** of what was changed for portability if the workspace is shared or the library is updated.

To verify the installed library matches:
```bash
diff lib_patches/anthropic/claude.py ../../../../libs/agno/agno/models/anthropic/claude.py
diff lib_patches/tools_python.py ../../../../libs/agno/agno/tools/python.py
```

## How to Verify

```bash
# Run the unit tests for claude.py fix
cd /mnt/c/Users/aviji/repo/agno
python -m pytest libs/agno/tests/unit/models/anthropic/test_structured_output_support.py -v
```
