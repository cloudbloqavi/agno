import os
from agno.agent import Agent
from agno.models.google import Gemini

agent = Agent(
    model=Gemini(id="gemini-3-pro-preview", search=True, max_output_tokens=8192),
    markdown=False
)
response = agent.run("Count from 1 to 500.", stream=True, yield_run_output=True)
res = None
for event in response:
    try:
        from agno.agent import RunOutput
        if isinstance(event, RunOutput):
            res = event
    except ImportError:
        res = event

if res and hasattr(res, "content"):
    print("Content length:", len(res.content))
    print("Last 100 chars:", res.content[-100:])
else:
    print("No valid content found.")
