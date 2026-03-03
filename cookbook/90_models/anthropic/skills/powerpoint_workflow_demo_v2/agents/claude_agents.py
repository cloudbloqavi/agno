"""
Claude-specific agent definitions for PowerPoint workflow.

Provides the 5 swappable agents using Anthropic Claude models and native tools.
The Content Generator agent (with PPTX skill) is NOT included — it stays
in the main workflow files because it is always Claude regardless of provider.

Models used:
    brand_style_analyzer    -> claude-sonnet-4-6  (fast, cheap, structured output)
    query_optimizer         -> claude-opus-4-6    (powerful, long context, research)
    fallback_code_agent     -> claude-opus-4-6    (code gen + execution)
    image_planner           -> gemini-3-flash-preview (unchanged — already uses Gemini)
    slide_quality_reviewer  -> gemini-2.5-flash   (unchanged — already uses Gemini)
"""

import os
import sys
from pathlib import Path
from typing import Dict

from agno.agent import Agent
from agno.models.google import Gemini
from agno.tools.python import PythonTools

# Import Claude from local patch (same as main workflow files)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from lib_patches.anthropic.claude import Claude

# Import shared models and instructions from the parent package
from agents._shared import (
    BRAND_STYLE_ANALYZER_INSTRUCTIONS,
    IMAGE_PLANNER_INSTRUCTIONS,
    PPTX_CODE_GEN_INSTRUCTIONS,
    SLIDE_QUALITY_REVIEWER_INSTRUCTIONS,
    BrandStyleIntent,
    ImagePlan,
    SlideQualityReport,
)


def create_agents() -> Dict[str, Agent]:
    """Create and return all 5 swappable agents using Claude models."""

    brand_style_analyzer = Agent(
        name="Brand Style Analyzer",
        model=Claude(
            id="claude-sonnet-4-6",
            max_tokens=8192,
            betas=["structured-outputs-2025-11-13"],
        ),
        description=(
            "You analyze user presentation requests to detect and extract branding "
            "or styling intent.  When a brand is mentioned, you decide whether you "
            "already know enough about the brand's visual identity (colors, tone, "
            "typography) or whether you need to search for brand guidelines."
        ),
        instructions=BRAND_STYLE_ANALYZER_INSTRUCTIONS,
        tools=[
            {
                "type": "web_search_20250305",
                "name": "web_search",
                "max_uses": 2,
            }
        ],
        output_schema=BrandStyleIntent,
        markdown=False,
    )

    # output_schema is intentionally omitted: claude-opus-4-6 does not support structured
    # outputs, which causes Agno to make an internal non-streaming extraction call that the
    # context-1m beta rejects ("Streaming is required for operations that may take longer
    # than 10 minutes").  The storyboard JSON is instead requested via prompt instructions
    # and parsed manually.
    query_optimizer = Agent(
        name="Presentation Strategist",
        model=Claude(
            id="claude-opus-4-6",
            betas=["context-1m-2025-08-07"],
            max_tokens=128000,
        ),
        description=(
            "You are a presentation strategist who first searches the web for current, "
            "relevant facts and data about the topic, then creates an optimized presentation "
            "plan with a per-slide storyboard grounded in that research."
        ),
        tools=[
            {
                "type": "web_search_20250305",
                "name": "web_search",
                "max_uses": 5,
            }
        ],
        markdown=False,
    )

    # betas=[\"context-1m-2025-08-07\"] intentionally omitted because this agent
    # is invoked with stream=False, and context-1m beta mandates streaming.
    fallback_code_agent = Agent(
        name="PPTX Code Generator",
        model=Claude(id="claude-opus-4-6", max_tokens=128000),
        instructions=PPTX_CODE_GEN_INSTRUCTIONS,
        tools=[
            PythonTools(
                base_dir=Path("."),
            )
        ],
        markdown=False,
    )

    image_planner = Agent(
        name="Image Planner",
        model=Gemini(id="gemini-3-flash-preview"),
        instructions=IMAGE_PLANNER_INSTRUCTIONS,
        output_schema=ImagePlan,
        markdown=False,
    )

    slide_quality_reviewer = Agent(
        name="Senior UI/UX Presentation Designer",
        model=Gemini(id="gemini-2.5-flash"),
        instructions=SLIDE_QUALITY_REVIEWER_INSTRUCTIONS,
        output_schema=SlideQualityReport,
        markdown=False,
    )

    return {
        "brand_style_analyzer": brand_style_analyzer,
        "query_optimizer": query_optimizer,
        "fallback_code_agent": fallback_code_agent,
        "image_planner": image_planner,
        "slide_quality_reviewer": slide_quality_reviewer,
    }
