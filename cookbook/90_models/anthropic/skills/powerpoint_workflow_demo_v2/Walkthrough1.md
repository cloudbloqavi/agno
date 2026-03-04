# PowerPoint Workflow Rate Limits Optimization — Walkthrough

## Summary of Fixes
The `powerpoint_chunked_workflow.py` script was running into a massive wall of HTTP 429 rate limit errors (from the Anthropic Claude API) because all functions were calling Anthropic models rapidly without any rate-limit awareness or cooldown. 

By analyzing the structure of the API calls and the individual models' constraints, I applied 8 targeted code structure changes to dramatically reduce token consumption, separate usage pools, and proactively abide by rate limits rather than just reacting to failures.

### What Was Changed & Validated:

1. **Two-Stage Brand/Style Parsing:**
   The original script called `claude-sonnet-4-6` for *every* prompt to extract branding intent, immediately eating into the 30K/min Anthropic token budget. 
   - *Fix:* Added a zero-cost Regex keyword pre-check. If a user asks a simple query (e.g. "Create a 5-slide deck about AI Trends"), it safely skips the LLM brand parsing completely (0 tokens used).
   - *Fix:* If branding *is* required, it now calls `gpt-4o-mini` (OpenAI), completely removing this load from the Anthropic quota.

2. **Smart Model Downgrades:**
   - **`query_optimizer`**: Downgraded from `claude-opus-4-6` to `claude-sonnet-4-6` (retains the same 30K limit but handles struct-outputs better with lower cost).
   - **`fallback_code_agent` (Tier 2)**: Downgraded from `claude-opus-4-6` to `claude-haiku-4-5`. This is a massive win because Haiku has its own separate **50,000 tokens/min** pool, meaning Tier 1 and Tier 2 Generation no longer cannibalize each other.

3. **Max Tokens Caps:**
   Reduced `max_tokens` on `query_optimizer` to 4096 and `fallback_code_agent` to 16384 (down from an excessive 128,000 which reserved too much space).

4. **Web Search Reduction:**
   Reduced the `query_optimizer` web search `max_uses` constraint from 5 to 2 to drastically cut down execution time and token consumption in Step 1.

5. **API Token Tracker Singleton:**
   Added a new `_RateLimitTracker` class to `powerpoint_chunked_workflow.py`. It tracks sequential Claude API calls, counts estimated input tokens, and aggregates them over a rolling 60-second window. It logs usages like: `[RATE TRACKER] claude-sonnet-4-6 — ~1072 estimated input tokens`. 
   
6. **Adaptive Inter-Chunk Delay:**
   The rigid 1.0 second delay between chunk generations was replaced with a random `60.0 to 120.0` second backoff `_inter_chunk_sleep()`. A visual interactive countdown tells the user exactly how long it is waiting.

7. **429 Rate-Limit Error Distinguisher:**
   Previously, a single 429 rate limit hit permanently disabled the Tier 1 generator for *all remaining chunks*. It now correctly tags 429s as "transient" errors, logging it, forcing a max-delay block, and safely retrying without breaking the state machine.

8. **Gemini Key Validation Guard:**
   Visual review steps previously crashed with a blind `400 INVALID_ARGUMENT: API key expired` and needlessly burned execution time repeating the failure. An upfront key validation guard automatically detects if `GOOGLE_API_KEY` is missing and disables visual review dynamically.

9. **Documentation Sync & Template Workflow Guard:**
   The `GOOGLE_API_KEY` validation guard has also been synchronized into the `powerpoint_template_workflow.py` script. The CLI flags in the docstrings, `PRODUCT.md`, and `ARCHITECTURE_powerpoint_chunked_workflow.md` have all been successfully batch-updated to reflect the exact models, logic, and rate trackers deployed.
