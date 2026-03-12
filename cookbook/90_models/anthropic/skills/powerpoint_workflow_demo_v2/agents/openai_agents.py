"""
OpenAI-specific agent definitions for PowerPoint workflow.

Provides the 5 swappable agents using OpenAI models via the Responses API.

Models used:
    brand_style_analyzer    -> gpt-5-mini   (fast, cost-effective, structured output)
    query_optimizer         -> gpt-5.2      (latest flagship, powerful research)
    fallback_code_agent     -> gpt-5.2      (powerful code gen + execution)
    image_planner           -> gpt-5-mini   (fast, structured output)
    slide_quality_reviewer  -> gpt-5-mini   (vision analysis)

Web search: Uses OpenAI's native web_search_preview tool via Responses API.
"""

from pathlib import Path
from typing import Dict

from agno.agent import Agent
from agno.models.google import Gemini
from agno.models.openai import OpenAIResponses
from agno.tools.python import PythonTools

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
    """Create and return all 5 swappable agents using OpenAI models."""

    brand_style_analyzer = Agent(
        name="Brand Style Analyzer",
        model=OpenAIResponses(id="gpt-5-mini"),
        description=(
            "You analyze user presentation requests to detect and extract branding "
            "or styling intent.  When a brand is mentioned, you decide whether you "
            "already know enough about the brand's visual identity (colors, tone, "
            "typography) or whether you need to search for brand guidelines."
        ),
        instructions=BRAND_STYLE_ANALYZER_INSTRUCTIONS,
        tools=[
            {"type": "web_search_preview"},
        ],
        output_schema=BrandStyleIntent,
        markdown=False,
    )

    query_optimizer = Agent(
        name="Presentation Strategist",
        model=OpenAIResponses(id="gpt-5.2"),
        description=(
            "You are a presentation strategist who first searches the web for current, "
            "relevant facts and data about the topic, then creates an optimized presentation "
            "plan with a per-slide storyboard grounded in that research."
        ),
        tools=[
            {"type": "web_search_preview"},
        ],
        markdown=False,
    )

    fallback_code_agent = Agent(
        name="PPTX Code Generator",
        model=OpenAIResponses(id="gpt-5.2"),
        instructions=PPTX_CODE_GEN_INSTRUCTIONS,
        tools=[
            PythonTools(
                base_dir=Path("."),
            )
        ],
        markdown=False,
    )

    fallback_code_agent_lite = Agent(
        name="PPTX Code Generator (Lite)",
        model=OpenAIResponses(id="gpt-5-mini"),
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
        model=OpenAIResponses(id="gpt-5-mini"),
        instructions=IMAGE_PLANNER_INSTRUCTIONS,
        output_schema=ImagePlan,
        markdown=False,
    )

    slide_quality_reviewer = Agent(
        name="Senior UI/UX Presentation Designer",
        model=OpenAIResponses(id="gpt-5-mini"),
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
        model=Gemini(id="gemini-3-flash-preview"),
        description="Fallback agent for Image Planner using Gemini in case of rate limits or errors.",
        instructions=IMAGE_PLANNER_INSTRUCTIONS,
        output_schema=ImagePlan,
        markdown=False,
    )

    slide_quality_reviewer_fallback = Agent(
        name="Senior UI/UX Presentation Designer (Fallback)",
        model=Gemini(id="gemini-2.5-flash"),
        description="Fallback agent for UI/UX Presentation Designer using Gemini in case of rate limits or errors.",
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
