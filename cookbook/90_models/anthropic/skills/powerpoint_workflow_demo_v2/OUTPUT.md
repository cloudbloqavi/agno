usage: powerpoint_chunked_workflow.py [-h] [--template TEMPLATE]
                                      [--output OUTPUT] [--prompt PROMPT]
                                      [--no-images] [--no-stream]
                                      [--min-images MIN_IMAGES]
                                      [--visual-review]
                                      [--footer-text FOOTER_TEXT]
                                      [--date-text DATE_TEXT]
                                      [--show-slide-numbers] [--verbose]
                                      [--llm-provider {claude,openai,gemini}]
                                      [--chunk-size CHUNK_SIZE]
                                      [--max-retries MAX_RETRIES]
                                      [--visual-passes VISUAL_PASSES]
                                      [--start-tier {1,2,3}]
                                      [--inter-chunk-delay-min MS]
                                      [--inter-chunk-delay-max MS]

Chunked PPTX generation workflow — overcomes Claude API limits for large
presentations.

options:
  -h, --help            show this help message and exit
  --template TEMPLATE, -t TEMPLATE
                        Path to .pptx template file (optional). Without it,
                        skips template assembly.
  --output OUTPUT, -o OUTPUT
                        Output filename (default: presentation_chunked.pptx).
  --prompt PROMPT, -p PROMPT
                        User prompt describing the presentation topic.
  --no-images           Skip AI image generation.
  --no-stream           Disable streaming mode for Claude agent.
  --min-images MIN_IMAGES
                        Minimum slides that must have AI-generated images
                        (default: 1).
  --visual-review       Enable visual QA with Gemini vision per chunk
                        (requires LibreOffice + template).
  --footer-text FOOTER_TEXT
                        Footer text for all slides (idx=11 placeholder).
  --date-text DATE_TEXT
                        Date text for footer date placeholder (idx=10).
  --show-slide-numbers  Preserve slide number placeholder (idx=12) on all
                        slides.
  --verbose, -v         Enable verbose/debug logging.
  --llm-provider {claude,openai,gemini}
                        LLM provider for swappable agents (brand analyzer,
                        query optimizer, fallback code gen, image planner,
                        visual reviewer). The Content Generator always uses
                        Claude (PPTX skill). Default: claude.
  --chunk-size CHUNK_SIZE
                        Number of slides per LLM API chunk call (default: 1).
                        Using 1 ensures each chunk sends only the single best-
                        matching template slide image, keeping prompts within
                        all model context windows.
  --max-retries MAX_RETRIES
                        Max retries per chunk on failure (default: 2).
  --visual-passes VISUAL_PASSES
                        Maximum visual inspection passes per chunk (default:
                        3).
  --start-tier {1,2,3}  Starting tier for chunk generation (default: 1).
                        1=Claude PPTX skill (best quality), 2=LLM code
                        generation (80-92% quality, faster, python-pptx native
                        charts), 3=text-only (structural, instant). Fallback
                        continues from selected tier.
  --inter-chunk-delay-min MS
                        Minimum inter-chunk delay in milliseconds (default:
                        provider-specific). A random value in [min, max] is
                        chosen between each chunk.
  --inter-chunk-delay-max MS
                        Maximum inter-chunk delay in milliseconds (default:
                        provider-specific). When a 429 rate-limit error is
                        detected, max_delay is used directly.
