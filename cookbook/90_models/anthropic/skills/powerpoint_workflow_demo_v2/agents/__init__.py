"""
Provider-based agent factory for PowerPoint workflow.

Usage:
    from agents import get_agents

    agents = get_agents("claude")  # or "openai" or "gemini"
    brand_analyzer = agents["brand_style_analyzer"]
    optimizer = agents["query_optimizer"]
    fallback = agents["fallback_code_agent"]
    planner = agents["image_planner"]
    reviewer = agents["slide_quality_reviewer"]

The Content Generator (chunk_agent / content_agent) is NOT included here
because it has a hard dependency on Claude's PPTX skill and cannot be swapped.
"""

from typing import Dict

from agno.agent import Agent

# Lazy imports to avoid pulling in all providers at module load time.
_PROVIDER_MODULES = {
    "claude": "agents.claude_agents",
    "openai": "agents.openai_agents",
    "gemini": "agents.gemini_agents",
}

SUPPORTED_PROVIDERS = list(_PROVIDER_MODULES.keys())

AGENT_ROLES = [
    "brand_style_analyzer",
    "query_optimizer",
    "fallback_code_agent",
    "fallback_code_agent_lite",
    "image_planner",
    "slide_quality_reviewer",
]


def get_agents(provider: str) -> Dict[str, Agent]:
    """Return a dict of pre-configured Agent instances for the given provider.

    Args:
        provider: One of "claude", "openai", or "gemini".

    Returns:
        Dict mapping role name -> Agent instance. Keys are:
            brand_style_analyzer, query_optimizer, fallback_code_agent,
            fallback_code_agent_lite, image_planner, slide_quality_reviewer.

    Raises:
        ValueError: If provider is not recognized.
    """
    provider = provider.lower().strip()
    if provider not in _PROVIDER_MODULES:
        raise ValueError(
            "Unknown LLM provider '%s'. Supported: %s"
            % (provider, ", ".join(SUPPORTED_PROVIDERS))
        )

    import importlib

    module = importlib.import_module(_PROVIDER_MODULES[provider])
    agents = module.create_agents()

    # Validate that all expected roles are present
    missing = [r for r in AGENT_ROLES if r not in agents]
    if missing:
        raise RuntimeError(
            "Provider module '%s' is missing agent roles: %s"
            % (provider, ", ".join(missing))
        )

    return agents
