"""
Claude-specific agent definitions for PowerPoint workflow.

Provides the 5+1 swappable agents using Anthropic Claude models and native tools.
The Content Generator agent (with PPTX skill) is NOT included — it stays
in the main workflow files because it is always Claude regardless of provider.

Models used (optimised to preserve the claude-opus-4-6 30K input-token/min budget):
    brand_style_analyzer      -> gpt-4o-mini  (OpenAI; completely separate rate-limit pool)
    query_optimizer           -> claude-sonnet-4-6   (was Opus; Sonnet is sufficient for
                                                      storyboarding and shares the same pool
                                                      but is much more token-efficient)
    fallback_code_agent       -> claude-sonnet-4-6   (primary Tier 2; better code gen quality
                                                      than haiku for native charts/infographics)
    fallback_code_agent_lite  -> claude-haiku-4-5    (lite Tier 2 fallback; 50K token/min pool;
                                                      used when sonnet fails or is rate-limited)
    image_planner             -> gemini-3-flash-preview (unchanged — already uses Gemini)
    slide_quality_reviewer    -> gemini-2.5-flash       (unchanged — already uses Gemini)

Rate limit reference (Tier 1 as of 2026-03):
    claude-sonnet-4-6  : 50 RPM | 30K input tokens/min | 8K output tokens/min
    claude-opus-4-6    : 50 RPM | 30K input tokens/min | 8K output tokens/min
    claude-haiku-4-5   : 50 RPM | 50K input tokens/min | 10K output tokens/min
    gpt-4o-mini        : OpenAI limits (separate company — no impact on Anthropic quota)
"""

import os
import sys
from pathlib import Path
from typing import Dict

from agno.agent import Agent
from agno.models.google import Gemini
from agno.models.openai import OpenAIChat
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
    """Create and return all 5 swappable agents using Claude models.

    Brand style analyzer uses OpenAI gpt-4o-mini to keep brand analysis completely
    off the Anthropic token-per-minute quota, preserving it for chunk generation.
    """

    # Brand analysis runs on gpt-4o-mini (OpenAI) so it doesn't consume Anthropic
    # input tokens.  The rate-limit pool for gpt-4o-mini is entirely separate.
    brand_style_analyzer = Agent(
        name="Brand Style Analyzer",
        model=OpenAIChat(
            id="gpt-4o-mini",
            max_tokens=2048,
        ),
        description=(
            "You analyze user presentation requests to detect and extract branding "
            "or styling intent.  When a brand is mentioned, you decide whether you "
            "already know enough about the brand's visual identity (colors, tone, "
            "typography) or whether you need to search for brand guidelines."
        ),
        instructions=BRAND_STYLE_ANALYZER_INSTRUCTIONS,
        output_schema=BrandStyleIntent,
        markdown=False,
    )

    # Downgraded from claude-opus-4-6 → claude-sonnet-4-6.
    # Sonnet is fully capable of building a structured storyboard JSON; Opus-level
    # reasoning is not needed here.  Both share the 30K input-token/min pool but
    # Sonnet uses significantly fewer tokens per call.
    # max_tokens capped at 4096: a 15-slide storyboard JSON is ~2,000-3,000 tokens.
    # output_schema omitted intentionally (same reason as before — Opus note still
    # applies in reverse: Sonnet with context-1m also requires streaming for large
    # outputs, so we parse JSON from prompt-instructed plain-text response).
    query_optimizer = Agent(
        name="Presentation Strategist",
        model=Claude(
            id="claude-sonnet-4-6",
            betas=["context-1m-2025-08-07"],
            max_tokens=4096,
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
                # Reduced from 5 → 2: limits web-search overhead in Step 1.
                # Each web search call adds latency and input tokens; 2 focused
                # searches provide sufficient grounding for a storyboard.
                "max_uses": 2,
            }
        ],
        markdown=False,
    )

    # Upgraded from claude-haiku-4-5 → claude-sonnet-4-6 for better python-pptx
    # code generation quality (native charts, infographics, shaped layouts).
    # Sonnet shares the 30K input-token/min pool with Opus and the query optimizer,
    # so the rate tracker must account for Tier 2 calls consuming that budget.
    # max_tokens capped at 16384: python-pptx scripts for 2 slides are ~500-1000 lines.
    fallback_code_agent = Agent(
        name="PPTX Code Generator",
        model=Claude(id="claude-sonnet-4-6", max_tokens=16384, betas=["context-1m-2025-08-07"]),
        instructions=PPTX_CODE_GEN_INSTRUCTIONS,
        tools=[
            PythonTools(
                base_dir=Path("."),
            )
        ],
        markdown=False,
    )

    # Lite fallback: used when sonnet fails (rate limit, error) before escalating
    # to Tier 3. Haiku has a SEPARATE 50K input-token/min budget.
    fallback_code_agent_lite = Agent(
        name="PPTX Code Generator (Lite)",
        model=Claude(id="claude-haiku-4-5", max_tokens=16384),
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

    brand_style_analyzer_fallback = Agent(
        name="Brand Style Analyzer (Fallback)",
        model=Gemini(id="gemini-3-flash-preview", search=True),
        description="Fallback agent for Brand Style Analyzer using Gemini in case of rate limits or errors.",
        instructions=BRAND_STYLE_ANALYZER_INSTRUCTIONS,
        output_schema=BrandStyleIntent,
        markdown=False,
    )

    query_optimizer_fallback = Agent(
        name="Presentation Strategist (Fallback)",
        model=Gemini(id="gemini-3-pro-preview", search=True),
        description="Fallback agent for Presentation Strategist using Gemini in case of rate limits or errors.",
        markdown=False,
    )

    fallback_code_agent_fallback = Agent(
        name="PPTX Code Generator (Fallback)",
        model=Gemini(id="gemini-3-pro-preview"),
        description="Fallback agent for PPTX Code Generator using Gemini in case of rate limits or errors.",
        instructions=PPTX_CODE_GEN_INSTRUCTIONS,
        tools=[
            PythonTools(
                base_dir=Path("."),
            )
        ],
        markdown=False,
    )

    fallback_code_agent_lite_fallback = Agent(
        name="PPTX Code Generator (Lite Fallback)",
        model=Gemini(id="gemini-2.5-flash"),
        description="Fallback agent for PPTX Code Generator (Lite) using Gemini in case of rate limits or errors.",
        instructions=PPTX_CODE_GEN_INSTRUCTIONS,
        tools=[
            PythonTools(
                base_dir=Path("."),
            )
        ],
        markdown=False,
    )

    image_planner_fallback = Agent(
        name="Image Planner (Fallback)",
        model=OpenAIChat(id="gpt-4o-mini"),
        description="Fallback agent for Image Planner using OpenAI in case of rate limits or errors.",
        instructions=IMAGE_PLANNER_INSTRUCTIONS,
        output_schema=ImagePlan,
        markdown=False,
    )

    slide_quality_reviewer_fallback = Agent(
        name="Senior UI/UX Presentation Designer (Fallback)",
        model=OpenAIChat(id="gpt-4o-mini"),
        description="Fallback agent for UI/UX Presentation Designer using OpenAI in case of rate limits or errors.",
        instructions=SLIDE_QUALITY_REVIEWER_INSTRUCTIONS,
        output_schema=SlideQualityReport,
        markdown=False,
    )

    return {
        "brand_style_analyzer": brand_style_analyzer,
        "brand_style_analyzer_fallback": brand_style_analyzer_fallback,
        "query_optimizer": query_optimizer,
        "query_optimizer_fallback": query_optimizer_fallback,
        "fallback_code_agent": fallback_code_agent,
        "fallback_code_agent_fallback": fallback_code_agent_fallback,
        "fallback_code_agent_lite": fallback_code_agent_lite,
        "fallback_code_agent_lite_fallback": fallback_code_agent_lite_fallback,
        "image_planner": image_planner,
        "image_planner_fallback": image_planner_fallback,
        "slide_quality_reviewer": slide_quality_reviewer,
        "slide_quality_reviewer_fallback": slide_quality_reviewer_fallback,
    }
