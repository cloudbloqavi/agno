"""
Unit tests for Claude._supports_structured_outputs() model detection.

Ensures that the structured output support check correctly distinguishes between:
- Old models that don't support structured outputs (claude-3-*, claude-sonnet-4, etc.)
- Dated pinned versions of old models (claude-sonnet-4-20250514)
- Newer sub-versions that DO support structured outputs (claude-sonnet-4-5, claude-sonnet-4-6)

Regression: claude-sonnet-4-6 was incorrectly excluded by an overly broad
startswith("claude-sonnet-4-") pattern check. Fixed by using a date-detection
heuristic instead of version-prefix exclusions.

Modified by: cloudbloqavi
"""

import pytest

from agno.models.anthropic.claude import Claude


# --- Models that SHOULD support structured outputs ---

@pytest.mark.parametrize(
    "model_id",
    [
        "claude-sonnet-4-5-20250929",  # Default model
        "claude-sonnet-4-5",           # Alias
        "claude-sonnet-4-6",           # Newer sub-version (regression case)
        "claude-sonnet-4-7",           # Future sub-version
        "claude-opus-4-1",             # Opus 4.1
        "claude-opus-4-5",             # Opus 4.5
        "claude-opus-4-6",             # Newer Opus sub-version
        "claude-opus-4-1-20250929",    # Dated Opus 4.1
    ],
)
def test_supports_structured_outputs_true(model_id: str):
    """Models that support structured outputs should return True."""
    claude = Claude(id=model_id)
    assert claude._supports_structured_outputs() is True, (
        f"Model '{model_id}' should support structured outputs"
    )


# --- Models that should NOT support structured outputs ---

@pytest.mark.parametrize(
    "model_id",
    [
        # Claude 3.x family
        "claude-3-opus-20240229",
        "claude-3-sonnet-20240229",
        "claude-3-haiku-20240307",
        "claude-3-opus",
        "claude-3-sonnet",
        "claude-3-haiku",
        # Claude 3.5 family
        "claude-3-5-sonnet-20240620",
        "claude-3-5-sonnet-20241022",
        "claude-3-5-sonnet",
        "claude-3-5-haiku-20241022",
        "claude-3-5-haiku-latest",
        "claude-3-5-haiku",
        # Claude Sonnet 4 base (no structured output support)
        "claude-sonnet-4-20250514",
        "claude-sonnet-4",
    ],
)
def test_does_not_support_structured_outputs(model_id: str):
    """Models without structured output support should return False."""
    claude = Claude(id=model_id)
    assert claude._supports_structured_outputs() is False, (
        f"Model '{model_id}' should NOT support structured outputs"
    )


# --- Dated version detection ---

def test_dated_sonnet_4_excluded():
    """Dated versions of claude-sonnet-4 (YYYYMMDD suffix) should be excluded."""
    claude = Claude(id="claude-sonnet-4-20250514")
    assert claude._supports_structured_outputs() is False


def test_future_dated_sonnet_4_excluded():
    """Future dated versions of claude-sonnet-4 should also be excluded."""
    claude = Claude(id="claude-sonnet-4-20260101")
    assert claude._supports_structured_outputs() is False


def test_dated_opus_4_excluded():
    """Dated versions of claude-opus-4 (YYYYMMDD suffix) should be excluded."""
    claude = Claude(id="claude-opus-4-20250601")
    assert claude._supports_structured_outputs() is False


# --- Regression: claude-sonnet-4-6 ---

def test_sonnet_4_6_supports_structured_outputs():
    """Regression test: claude-sonnet-4-6 MUST support structured outputs.

    Previously excluded by: startswith("claude-sonnet-4-") and not startswith("claude-sonnet-4-5")
    """
    claude = Claude(id="claude-sonnet-4-6")
    assert claude._supports_structured_outputs() is True


def test_sonnet_4_6_sets_capability_flag():
    """claude-sonnet-4-6 should have supports_native_structured_outputs=True after init."""
    claude = Claude(id="claude-sonnet-4-6")
    assert claude.supports_native_structured_outputs is True
