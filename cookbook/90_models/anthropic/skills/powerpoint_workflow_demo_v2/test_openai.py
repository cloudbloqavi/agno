import urllib3
urllib3.disable_warnings()

from agno.agent import Agent
from agno.models.openai import OpenAIResponses
try:
    model = OpenAIResponses(id="gpt-5.2", max_completion_tokens=8192)
    print("OpenAIResponses with max_completion_tokens success!")
except Exception as e:
    print("Error:", e)
