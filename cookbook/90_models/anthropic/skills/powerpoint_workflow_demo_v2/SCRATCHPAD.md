I have few concerns related to current powerpoint_chunked_workflow.py (main entry script) workflow logic.

###
- the logic doesn't work good, i.e., fails to follow and incorporate into the final presentation pptx file for the visually enriched template files i.e. template files with lot of charts, graphs, infographics, smart arts etc.

- repeated visual issues like text placeholders, text boxes, shapes, images etc. are not placed correctly, overlapping, not visible, not editable, etc.

- For each execution step it's not clear (even in verbose mode) currently which Agent is active or which LLM provider is currently active.

- I think dynamic python tool agent - "PPTX Code Generator" is not working as good as expected and we may need to make this portion strong with high end models like claude opus or claude sonnet (including a fallback logic to use haiku model)

- If start the workflow with "Tier 1" mode I can see even for the first Claude with PPTX skill agent call it's throwing rate limit of 30k and then retry mechnism with random values of 200ms, 1000ms etc. As it's a per minute token limit, we may need to improve it with min. 1 min delay between each call with some randomness. We may need to investigate and check if our approach is correct of using this - "Content Generator" agent properly or not. May be adding a proper test case file only for this agent just to check if the agent or claude api with skills really working or we have some other logic or implementation issue.

- For tools like Manus prsentation or Claude Powerpoint Addon presentation, I have seen if attach a screenshot of an existing ppt or templat file (single slide) ask the AI tool to create a slide based on that, it nearly approx 80-90% can mimic the the output slide in powerpoint native format. Can we use this approach?
###

I need your help if it's possible to address the above pointers and incorporate efficient, robust logic for our current workflow.