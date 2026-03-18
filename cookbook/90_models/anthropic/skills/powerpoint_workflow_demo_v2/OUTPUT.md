usage: powerpoint_chunked_workflow.py
       [-h] [--template TEMPLATE]
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

Chunked PPTX generation workflow — overcomes
Claude API limits for large presentations.

options:
  -h, --help            show this help message
                        and exit
  --template TEMPLATE, -t TEMPLATE
                        Path to .pptx template
                        file (optional). Without
                        it, skips template
                        assembly.
  --output OUTPUT, -o OUTPUT
                        Output filename
                        (default: presentation_c
                        hunked.pptx).
  --prompt PROMPT, -p PROMPT
                        User prompt describing
                        the presentation topic.
  --no-images           Skip AI image
                        generation.
  --no-stream           Disable streaming mode
                        for Claude agent.
  --min-images MIN_IMAGES
                        Minimum slides that must
                        have AI-generated images
                        (default: 1).
  --visual-review       Enable visual QA with
                        Gemini vision per chunk
                        (requires LibreOffice +
                        template).
  --footer-text FOOTER_TEXT
                        Footer text for all
                        slides (idx=11
                        placeholder).
  --date-text DATE_TEXT
                        Date text for footer
                        date placeholder
                        (idx=10).
  --show-slide-numbers  Preserve slide number
                        placeholder (idx=12) on
                        all slides.
  --verbose, -v         Enable verbose/debug
                        logging.
  --llm-provider {claude,openai,gemini}
                        LLM provider for
                        swappable agents (brand
                        analyzer, query
                        optimizer, fallback code
                        gen, image planner,
                        visual reviewer). The
                        Content Generator always
                        uses Claude (PPTX
                        skill). Default: claude.
  --chunk-size CHUNK_SIZE
                        Number of slides per LLM
                        API chunk call (default:
                        1). Using 1 ensures each
                        chunk sends only the
                        single best-matching
                        template slide image,
                        keeping prompts within
                        all model context
                        windows.
  --max-retries MAX_RETRIES
                        Max retries per chunk on
                        failure (default: 2).
  --visual-passes VISUAL_PASSES
                        Maximum visual
                        inspection passes per
                        chunk (default: 3).
  --start-tier {1,2,3}  Starting tier for chunk
                        generation (default: 1).
                        1=Claude PPTX skill
                        (best quality), 2=LLM
                        code generation (80-92%
                        quality, faster, python-
                        pptx native charts),
                        3=text-only (structural,
                        instant). Fallback
                        continues from selected
                        tier.
  --inter-chunk-delay-min MS
                        Minimum inter-chunk
                        delay in milliseconds
                        (default: provider-
                        specific). A random
                        value in [min, max] is
                        chosen between each
                        chunk.
  --inter-chunk-delay-max MS
                        Maximum inter-chunk
                        delay in milliseconds
                        (default: provider-
                        specific). When a 429
                        rate-limit error is
                        detected, max_delay is
                        used directly.
[CHUNK 1 RETRY] Waiting... 59s remaining (89s total)
[CHUNK 7 RETRY] Waiting... 40s remaining (70s total)
[CHUNK 4 RETRY] Waiting... 58s remaining (88s total)
[CHUNK 3 RETRY] Waiting... 44s remaining (74s total)
[CHUNK 2 RETRY] Waiting... 32s remaining (62s total)
[CHUNK 5 RETRY] Waiting... 56s remaining (86s total)
[CHUNK 0 RETRY] Waiting... 32s remaining (62s total)
[CHUNK 6 RETRY] Waiting... 49s remaining (79s total)
[CHUNK 1 RETRY] Waiting... 44s remaining (89s total)
[CHUNK 7 RETRY] Waiting... 25s remaining (70s total)
[CHUNK 4 RETRY] Waiting... 43s remaining (88s total)
[CHUNK 3 RETRY] Waiting... 29s remaining (74s total)
[CHUNK 2 RETRY] Waiting... 17s remaining (62s total)
[CHUNK 5 RETRY] Waiting... 41s remaining (86s total)
[CHUNK 0 RETRY] Waiting... 17s remaining (62s total)
[CHUNK 6 RETRY] Waiting... 34s remaining (79s total)
[CHUNK 1 RETRY] Waiting... 29s remaining (89s total)
[CHUNK 7 RETRY] Final 10s...
[CHUNK 4 RETRY] Waiting... 28s remaining (88s total)
[CHUNK 3 RETRY] Final 14s...
[CHUNK 2 RETRY] Final 2s...
[CHUNK 5 RETRY] Waiting... 26s remaining (86s total)
[CHUNK 0 RETRY] Final 2s...
[CHUNK 6 RETRY] Waiting... 19s remaining (79s total)
[CHUNK 2] API call attempt 2/3 (slides 3-3)...

┌──────────────────────────────────────────────────
│ 🤖 AGENT: Chunk Generator 2
│ 📡 MODEL: claude-opus-4-6 [Anthropic]
│ 📋 STEP:  step_generate_chunks / Tier 1 PPTX Skill (attempt 2/3)
└──────────────────────────────────────────────────
[RATE TRACKER] [generate_chunk_pptx/Tier1] claude-opus-4-6 — ~1414 estimated input tokens | window so far: ~0 / 30000 tokens/min
[CHUNK 0] API call attempt 2/3 (slides 1-1)...

┌──────────────────────────────────────────────────
│ 🤖 AGENT: Chunk Generator 0
│ 📡 MODEL: claude-opus-4-6 [Anthropic]
│ 📋 STEP:  step_generate_chunks / Tier 1 PPTX Skill (attempt 2/3)
└──────────────────────────────────────────────────
[RATE TRACKER] [generate_chunk_pptx/Tier1] claude-opus-4-6 — ~1418 estimated input tokens | window so far: ~1414 / 30000 tokens/min
[1;31mERROR   [0m Claude API error [1m([0mstatus [1;36m200[0m[1m)[0m: [1m{[0m[32m'type'[0m:  
         [32m'error'[0m, [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m,      
         [32m'type'[0m: [32m'overloaded_error'[0m, [32m'message'[0m:   
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA66A2324e1XshAxpruA'[0m[1m}[0m          
[1;31mERROR   [0m Error in Agent run: [1m{[0m[32m'type'[0m: [32m'error'[0m,    
         [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m, [32m'type'[0m:       
         [32m'overloaded_error'[0m, [32m'message'[0m:           
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA66A2324e1XshAxpruA'[0m[1m}[0m          
[CHUNK 2] No RunOutput received after 1 events.
[TIMING] Chunk 2 attempt 2/3: 65.1s (no output)
[CHUNK 2] Retry 2/2 — cooling down for 68s (rate limit window reset)...
[CHUNK 2 RETRY] Waiting... 68s remaining (68s total)
[1;31mERROR   [0m Claude API error [1m([0mstatus [1;36m200[0m[1m)[0m: [1m{[0m[32m'type'[0m:  
         [32m'error'[0m, [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m,      
         [32m'type'[0m: [32m'overloaded_error'[0m, [32m'message'[0m:   
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA66CBGAtmiiRokhDsk7'[0m[1m}[0m          
[1;31mERROR   [0m Error in Agent run: [1m{[0m[32m'type'[0m: [32m'error'[0m,    
         [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m, [32m'type'[0m:       
         [32m'overloaded_error'[0m, [32m'message'[0m:           
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA66CBGAtmiiRokhDsk7'[0m[1m}[0m          
[CHUNK 0] No RunOutput received after 1 events.
[TIMING] Chunk 0 attempt 2/3: 66.0s (no output)
[CHUNK 0] Retry 2/2 — cooling down for 82s (rate limit window reset)...
[CHUNK 0 RETRY] Waiting... 82s remaining (82s total)
[CHUNK 7] API call attempt 2/3 (slides 8-8)...

┌──────────────────────────────────────────────────
│ 🤖 AGENT: Chunk Generator 7
│ 📡 MODEL: claude-opus-4-6 [Anthropic]
│ 📋 STEP:  step_generate_chunks / Tier 1 PPTX Skill (attempt 2/3)
└──────────────────────────────────────────────────
[RATE TRACKER] [generate_chunk_pptx/Tier1] claude-opus-4-6 — ~1402 estimated input tokens | window so far: ~2832 / 30000 tokens/min
[1;31mERROR   [0m Claude API error [1m([0mstatus [1;36m200[0m[1m)[0m: [1m{[0m[32m'type'[0m:  
         [32m'error'[0m, [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m,      
         [32m'type'[0m: [32m'overloaded_error'[0m, [32m'message'[0m:   
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA66iWeA5hyWXWs1qLhe'[0m[1m}[0m          
[1;31mERROR   [0m Error in Agent run: [1m{[0m[32m'type'[0m: [32m'error'[0m,    
         [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m, [32m'type'[0m:       
         [32m'overloaded_error'[0m, [32m'message'[0m:           
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA66iWeA5hyWXWs1qLhe'[0m[1m}[0m          
[CHUNK 7] No RunOutput received after 1 events.
[TIMING] Chunk 7 attempt 2/3: 72.7s (no output)
[CHUNK 7] Retry 2/2 — cooling down for 77s (rate limit window reset)...
[CHUNK 7 RETRY] Waiting... 77s remaining (77s total)
[CHUNK 3] API call attempt 2/3 (slides 4-4)...

┌──────────────────────────────────────────────────
│ 🤖 AGENT: Chunk Generator 3
│ 📡 MODEL: claude-opus-4-6 [Anthropic]
│ 📋 STEP:  step_generate_chunks / Tier 1 PPTX Skill (attempt 2/3)
└──────────────────────────────────────────────────
[RATE TRACKER] [generate_chunk_pptx/Tier1] claude-opus-4-6 — ~1431 estimated input tokens | window so far: ~4234 / 30000 tokens/min
[CHUNK 1 RETRY] Final 14s...
[CHUNK 4 RETRY] Final 13s...
[CHUNK 5 RETRY] Final 11s...
[CHUNK 6 RETRY] Final 4s...
[CHUNK 6] API call attempt 2/3 (slides 7-7)...

┌──────────────────────────────────────────────────
│ 🤖 AGENT: Chunk Generator 6
│ 📡 MODEL: claude-opus-4-6 [Anthropic]
│ 📋 STEP:  step_generate_chunks / Tier 1 PPTX Skill (attempt 2/3)
└──────────────────────────────────────────────────
[RATE TRACKER] [generate_chunk_pptx/Tier1] claude-opus-4-6 — ~1419 estimated input tokens | window so far: ~5665 / 30000 tokens/min
[CHUNK 2 RETRY] Waiting... 53s remaining (68s total)
[1;31mERROR   [0m Claude API error [1m([0mstatus [1;36m200[0m[1m)[0m: [1m{[0m[32m'type'[0m:  
         [32m'error'[0m, [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m,      
         [32m'type'[0m: [32m'overloaded_error'[0m, [32m'message'[0m:   
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA67HzHhXAjz9akg4fK4'[0m[1m}[0m          
[1;31mERROR   [0m Error in Agent run: [1m{[0m[32m'type'[0m: [32m'error'[0m,    
         [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m, [32m'type'[0m:       
         [32m'overloaded_error'[0m, [32m'message'[0m:           
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA67HzHhXAjz9akg4fK4'[0m[1m}[0m          
[CHUNK 3] No RunOutput received after 1 events.
[TIMING] Chunk 3 attempt 2/3: 80.4s (no output)
[CHUNK 3] Retry 2/2 — cooling down for 86s (rate limit window reset)...
[CHUNK 3 RETRY] Waiting... 86s remaining (86s total)
[CHUNK 0 RETRY] Waiting... 67s remaining (82s total)
[1;31mERROR   [0m Claude API error [1m([0mstatus [1;36m200[0m[1m)[0m: [1m{[0m[32m'type'[0m:  
         [32m'error'[0m, [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m,      
         [32m'type'[0m: [32m'overloaded_error'[0m, [32m'message'[0m:   
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA67TdidMVD2qYeP539A'[0m[1m}[0m          
[1;31mERROR   [0m Error in Agent run: [1m{[0m[32m'type'[0m: [32m'error'[0m,    
         [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m, [32m'type'[0m:       
         [32m'overloaded_error'[0m, [32m'message'[0m:           
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA67TdidMVD2qYeP539A'[0m[1m}[0m          
[CHUNK 6] No RunOutput received after 1 events.
[TIMING] Chunk 6 attempt 2/3: 82.5s (no output)
[CHUNK 6] Retry 2/2 — cooling down for 65s (rate limit window reset)...
[CHUNK 6 RETRY] Waiting... 65s remaining (65s total)
[CHUNK 5] API call attempt 2/3 (slides 6-6)...

┌──────────────────────────────────────────────────
│ 🤖 AGENT: Chunk Generator 5
│ 📡 MODEL: claude-opus-4-6 [Anthropic]
│ 📋 STEP:  step_generate_chunks / Tier 1 PPTX Skill (attempt 2/3)
└──────────────────────────────────────────────────
[RATE TRACKER] [generate_chunk_pptx/Tier1] claude-opus-4-6 — ~1437 estimated input tokens | window so far: ~7084 / 30000 tokens/min
[CHUNK 7 RETRY] Waiting... 62s remaining (77s total)
[CHUNK 4] API call attempt 2/3 (slides 5-5)...

┌──────────────────────────────────────────────────
│ 🤖 AGENT: Chunk Generator 4
│ 📡 MODEL: claude-opus-4-6 [Anthropic]
│ 📋 STEP:  step_generate_chunks / Tier 1 PPTX Skill (attempt 2/3)
└──────────────────────────────────────────────────
[RATE TRACKER] [generate_chunk_pptx/Tier1] claude-opus-4-6 — ~1429 estimated input tokens | window so far: ~8521 / 30000 tokens/min
[CHUNK 1] API call attempt 2/3 (slides 2-2)...

┌──────────────────────────────────────────────────
│ 🤖 AGENT: Chunk Generator 1
│ 📡 MODEL: claude-opus-4-6 [Anthropic]
│ 📋 STEP:  step_generate_chunks / Tier 1 PPTX Skill (attempt 2/3)
└──────────────────────────────────────────────────
[RATE TRACKER] [generate_chunk_pptx/Tier1] claude-opus-4-6 — ~1409 estimated input tokens | window so far: ~9950 / 30000 tokens/min
[1;31mERROR   [0m Claude API error [1m([0mstatus [1;36m200[0m[1m)[0m: [1m{[0m[32m'type'[0m:  
         [32m'error'[0m, [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m,      
         [32m'type'[0m: [32m'overloaded_error'[0m, [32m'message'[0m:   
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA67vJ6DCsjNMaR7uKwH'[0m[1m}[0m          
[1;31mERROR   [0m Error in Agent run: [1m{[0m[32m'type'[0m: [32m'error'[0m,    
         [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m, [32m'type'[0m:       
         [32m'overloaded_error'[0m, [32m'message'[0m:           
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA67vJ6DCsjNMaR7uKwH'[0m[1m}[0m          
[CHUNK 5] No RunOutput received after 1 events.
[TIMING] Chunk 5 attempt 2/3: 88.7s (no output)
[CHUNK 5] Retry 2/2 — cooling down for 62s (rate limit window reset)...
[CHUNK 5 RETRY] Waiting... 62s remaining (62s total)
[1;31mERROR   [0m Claude API error [1m([0mstatus [1;36m200[0m[1m)[0m: [1m{[0m[32m'type'[0m:  
         [32m'error'[0m, [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m,      
         [32m'type'[0m: [32m'overloaded_error'[0m, [32m'message'[0m:   
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA685GqhsffAfcUKhzsM'[0m[1m}[0m          
[1;31mERROR   [0m Error in Agent run: [1m{[0m[32m'type'[0m: [32m'error'[0m,    
         [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m, [32m'type'[0m:       
         [32m'overloaded_error'[0m, [32m'message'[0m:           
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA685GqhsffAfcUKhzsM'[0m[1m}[0m          
[CHUNK 4] No RunOutput received after 1 events.
[TIMING] Chunk 4 attempt 2/3: 91.0s (no output)
[CHUNK 4] Retry 2/2 — cooling down for 63s (rate limit window reset)...
[CHUNK 4 RETRY] Waiting... 63s remaining (63s total)
[1;31mERROR   [0m Claude API error [1m([0mstatus [1;36m200[0m[1m)[0m: [1m{[0m[32m'type'[0m:  
         [32m'error'[0m, [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m,      
         [32m'type'[0m: [32m'overloaded_error'[0m, [32m'message'[0m:   
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA687qNmoYsFU5VfPuSm'[0m[1m}[0m          
[1;31mERROR   [0m Error in Agent run: [1m{[0m[32m'type'[0m: [32m'error'[0m,    
         [32m'error'[0m: [1m{[0m[32m'details'[0m: [3;35mNone[0m, [32m'type'[0m:       
         [32m'overloaded_error'[0m, [32m'message'[0m:           
         [32m'Overloaded'[0m[1m}[0m, [32m'request_id'[0m:             
         [32m'req_011CZA687qNmoYsFU5VfPuSm'[0m[1m}[0m          
[CHUNK 1] No RunOutput received after 1 events.
[TIMING] Chunk 1 attempt 2/3: 91.9s (no output)
[CHUNK 1] Retry 2/2 — cooling down for 71s (rate limit window reset)...
[CHUNK 1 RETRY] Waiting... 71s remaining (71s total)
[CHUNK 2 RETRY] Waiting... 38s remaining (68s total)
[CHUNK 3 RETRY] Waiting... 71s remaining (86s total)
[CHUNK 0 RETRY] Waiting... 52s remaining (82s total)
[CHUNK 6 RETRY] Waiting... 50s remaining (65s total)
