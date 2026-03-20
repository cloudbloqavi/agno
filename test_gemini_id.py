from agno.models.google import Gemini
m = Gemini(id="gemini-3-pro-preview")
print("Has id?", hasattr(m, "id"))
print("Has id value:", getattr(m, "id", None))
