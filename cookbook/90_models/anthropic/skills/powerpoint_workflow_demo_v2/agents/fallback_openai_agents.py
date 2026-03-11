"""
Universal OpenAI fallback agent definitions for PowerPoint workflow.

Provides highly available fallback agents using OpenAI's latest Pro and Lite models.
These agents sit beneath the primary provider (Claude/Gemini/OpenAI) and are
triggered dynamically when upstream capacity errors (HTTP 429/529) occur during chunk generation or core tasks.

Models used (March 2026 specifications):
    fallback_content_generator  -> gpt-5.4      (Flagship reasoning, 1M context, robust for PPTX)
    fallback_code_agent         -> gpt-5.4      (Complex code execution & structural fixes)
    fallback_code_agent_lite    -> o3-mini      (Cost-efficient, fast code execution/corrections)

Note: The Content Generator is typically bound to Claude's PPTX skill. The fallback
version here leverages GPT-5.4's advanced coding/tool capabilities combined with python-pptx
to bridge the gap when Claude is entirely unavailable.
"""

from pathlib import Path
from typing import Dict

from agno.agent import Agent
from agno.models.openai import OpenAIResponses
from agno.tools.python import PythonTools

from agents._shared import (
    PPTX_CODE_GEN_INSTRUCTIONS,
    SlideQualityReport,
)


def get_openai_fallback_agents() -> Dict[str, Agent]:
    """Create and return the universal fallback agents using OpenAI models.
    
    These agents mirror the required capabilities of the Tier 1 / Tier 2 fallback chains in
    the main orchestrator, offering high reliability during primary provider outages.
    """

    # GPT-5.4 Pro fallback for complex code generation, layout handling, and tier 2 escalation
    fallback_code_agent = Agent(
        name="Universal PPTX Code Generator (Fallback)",
        model=OpenAIResponses(id="gpt-5.4"),
        instructions=PPTX_CODE_GEN_INSTRUCTIONS,
        tools=[
            PythonTools(
                base_dir=Path("."),
            )
        ],
        markdown=False,
    )

    # o3-mini Lite fallback for faster, minor code adjustments/corrections
    fallback_code_agent_lite = Agent(
        name="Universal PPTX Code Generator Lite (Fallback)",
        model=OpenAIResponses(id="o3-mini"),
        instructions=PPTX_CODE_GEN_INSTRUCTIONS,
        tools=[
            PythonTools(
                base_dir=Path("."),
            )
        ],
        markdown=False,
    )
    
    # Pro model for robust web search and storyboard planning during Tier 1 outages
    fallback_query_optimizer = Agent(
        name="Universal Presentation Strategist (Fallback)",
        model=OpenAIResponses(id="gpt-5.4"),
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
    
    # Advanced logic: If Claude native PPTX generation (Tier 1) fails entirely due to limits, 
    # the fallback content generator uses GPT-5.4 to output python-pptx manipulation scripts 
    # that achieve near-Tier 1 quality, prioritizing native charts and visual richness.
    fallback_content_generator = Agent(
        name="Universal PPTX Content Generator (Fallback)",
        model=OpenAIResponses(id="gpt-5.4"),
        description=(
            "You are a master presentation structurer and Python PPTX expert. "
            "You generate rich PowerPoint slides using python-pptx code to create "
            "visually stunning charts, tables, and layouts when the primary presentation "
            "engine is unavailable."
        ),
        instructions=PPTX_CODE_GEN_INSTRUCTIONS, # Share the same robust code-gen instructions
        tools=[
            PythonTools(
                base_dir=Path("."),
            )
        ],
        markdown=False,
    )

    return {
        "fallback_content_generator": fallback_content_generator,
        "fallback_code_agent": fallback_code_agent,
        "fallback_code_agent_lite": fallback_code_agent_lite,
        "fallback_query_optimizer": fallback_query_optimizer,
    }
