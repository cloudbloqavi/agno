import os
import json
from agno.agent import Agent
from agno.models.openai import OpenAIChat
from dotenv import load_dotenv

load_dotenv()

agent = Agent(
    model=OpenAIChat(id="gpt-4o-mini"),
    markdown=True,
    add_datetime_to_instructions=True,
)
resp = agent.run("Hi, say hello quickly!")
print(resp.metrics)
if hasattr(agent, "run_response") and agent.run_response:
    print(agent.run_response.metrics)

