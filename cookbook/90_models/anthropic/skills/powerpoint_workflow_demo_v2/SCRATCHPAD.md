I have copied "theme-factory" folder locally from anthropic skills folder from github repo - https://github.com/anthropics/skills. 

I have a thought. For my  current codebase inside "powerpoint_workflow_demo_v2" where I typically run the script without any template file like this (some sample examples), 

  ```
  /mnt/c/Users/aviji/repo/agno/.venvs/demo/bin/python powerpoint_chunked_workflow.py \
  -p "Research on latest (2026) Coca-Cola market cap in India, including Stock Market performance in last 3 months and create a 3-slide presentation with proper visuals. Use Coca-Cola branding style like colors, fonts etc. with modern, dark theme." \
  --chunk-size 1 \
  --start-tier 2 \
  --no-images \
  --verbose \
  --visual-review \
  --visual-passes 3 \
  --llm-provider openai \
  -o cocacola_india.pptx
  ```

I want to improve my current code logic by using "theme-factory" folder from anthropic skills folder as copied here. But I don't want any interactive behavior with end user as mentioned in the SKILL.md file. So I want to modify it accordingly based on my current codebase structure and logic. Then I want to leverage this skill in my Agno workflow with it's inbuilt Skills support (https://docs.agno.com/skills/overview). You can use Agno docs mcp for this.

