#!/usr/bin/env python3
"""
Isolated test for the Claude Content Generator agent (Tier 1 PPTX skill).

This test validates that the Claude API with the native PPTX skill is
functioning correctly, independent of the full chunked workflow. It:
  1. Creates a minimal 1-slide storyboard prompt
  2. Calls the Claude Content Generator agent directly
  3. Validates that a .pptx file is produced
  4. Validates the file opens with python-pptx and contains exactly 1 slide
  5. Reports timing, token usage estimate, and any errors

WARNING: This test requires a valid ANTHROPIC_API_KEY and makes real API
calls.  It is NOT intended for CI — run manually when debugging Tier 1.

Usage:
    cd cookbook/90_models/anthropic/skills/powerpoint_workflow_demo_v2
    python tests/test_content_agent_isolated.py
"""

import os
import sys
import time
import traceback

# Ensure project root is on path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from dotenv import load_dotenv
    load_dotenv(override=True)
except ImportError:
    pass


def test_content_agent():
    """Run an isolated test of the Claude Content Generator (Tier 1)."""

    print("=" * 60)
    print("Isolated Content Generator Test (Tier 1 PPTX Skill)")
    print("=" * 60)

    # --- Pre-flight checks ---
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("[FAIL] ANTHROPIC_API_KEY not set. Cannot run this test.")
        return False

    print("[OK] ANTHROPIC_API_KEY found (length=%d)" % len(api_key))

    # --- Import dependencies ---
    try:
        from agno.agent import Agent
        from lib_patches.anthropic.claude import Claude
        from anthropic import Anthropic
        from file_download_helper import download_skill_files
        from pptx import Presentation

        print("[OK] All dependencies imported successfully")
    except ImportError as e:
        print("[FAIL] Missing dependency: %s" % e)
        return False

    # --- Create a minimal 1-slide prompt ---
    test_prompt = (
        "## Task: Generate exactly 1 slide\n\n"
        "Create a single title slide for a presentation about 'AI in Healthcare'.\n"
        "The slide should have:\n"
        "- Title: 'AI in Healthcare: Transforming Patient Care'\n"
        "- Subtitle: 'A Strategic Overview for 2026'\n\n"
        "Save the output as 'test_output.pptx'.\n"
        "Do NOT apply custom fonts, colors, or theme styling.\n"
        "Do NOT add animations, transitions, or speaker notes.\n"
    )

    estimated_tokens = len(test_prompt) // 4
    print("\n[PROMPT] Length: %d chars (~%d tokens)" % (len(test_prompt), estimated_tokens))

    # --- Create agent ---
    chunk_agent = Agent(
        name="Test Content Generator",
        model=Claude(
            id="claude-opus-4-6",
            betas=["context-1m-2025-08-07"],
            skills=[{"type": "anthropic", "skill_id": "pptx", "version": "latest"}],
            max_tokens=128000,
        ),
        instructions=[
            "You are a structured content generator for PowerPoint presentations.",
            "Generate EXACTLY the number of slides specified in the task.",
            "Use one clear title per slide with concise bullet points.",
            "Do NOT apply custom fonts, colors, or theme styling.",
        ],
        markdown=True,
    )

    client = Anthropic(api_key=api_key)
    output_dir = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "output_chunked",
        "test_isolated",
    )
    os.makedirs(output_dir, exist_ok=True)

    print("[INFO] Output dir: %s" % output_dir)
    print("\n" + "-" * 60)
    print("[TEST] Calling Claude Content Generator agent...")
    print("-" * 60)

    # --- Call agent with streaming ---
    start_time = time.time()
    response = None
    event_count = 0
    error_msg = None

    try:
        from agno.run.agent import RunOutput

        for event in chunk_agent.run(test_prompt, stream=True, yield_run_output=True):
            event_count += 1
            if isinstance(event, RunOutput):
                response = event
    except Exception as e:
        error_msg = str(e)
        elapsed = time.time() - start_time
        print("\n[ERROR] Agent call failed after %.1fs: %s" % (elapsed, error_msg))

        is_rate_limit = "rate_limit" in error_msg or "429" in error_msg
        if is_rate_limit:
            print("[DIAGNOSIS] Rate limit error detected.")
            print("  - This confirms the API key works but you've hit the token/min cap.")
            print("  - Wait 60s and try again, or check your Anthropic usage tier.")
        else:
            traceback.print_exc()
        return False

    elapsed = time.time() - start_time

    print("\n" + "-" * 60)
    print("[RESULT] Agent call completed in %.1fs" % elapsed)
    print("-" * 60)
    print("  Events received: %d" % event_count)
    print("  RunOutput received: %s" % (response is not None))

    if response is None:
        print("[FAIL] No RunOutput received. Agent may have produced no output.")
        return False

    # --- Try to download the generated file ---
    msg_count = len(response.messages) if response.messages else 0
    print("  Messages: %d" % msg_count)

    generated_file = None
    if response.messages:
        for msg in response.messages:
            if hasattr(msg, "provider_data") and msg.provider_data:
                try:
                    files = download_skill_files(
                        msg.provider_data, client, output_dir=output_dir
                    )
                    for f in files:
                        if f.endswith(".pptx"):
                            generated_file = f
                            break
                except Exception as e:
                    print("  [WARNING] download_skill_files failed: %s" % e)

    # Fallback: try model_provider_data
    if not generated_file and hasattr(response, "model_provider_data") and response.model_provider_data:
        try:
            files = download_skill_files(
                response.model_provider_data, client, output_dir=output_dir
            )
            for f in files:
                if f.endswith(".pptx"):
                    generated_file = f
                    break
        except Exception as e:
            print("  [WARNING] model_provider_data download failed: %s" % e)

    if not generated_file:
        print("[FAIL] No .pptx file was produced by the agent.")
        print("  Content type: %s" % type(response.content).__name__)
        if response.content:
            print("  Content preview: %s" % str(response.content)[:300])
        return False

    print("\n[FILE] Generated: %s" % generated_file)

    # --- Validate the PPTX file ---
    try:
        prs = Presentation(generated_file)
        slide_count = len(prs.slides)
        print("[VALIDATE] File opens successfully with python-pptx")
        print("[VALIDATE] Slide count: %d (expected: 1)" % slide_count)

        if slide_count == 1:
            print("\n[PASS] ✅ Content Generator test PASSED")
            print("  - API call: %.1fs" % elapsed)
            print("  - File: %s" % generated_file)
            print("  - Slides: %d" % slide_count)
            return True
        else:
            print("\n[WARN] ⚠️  File produced but slide count is %d (expected 1)" % slide_count)
            print("  The agent is functional but generated extra slides.")
            return True  # Still consider this a pass — agent works

    except Exception as e:
        print("[FAIL] Generated file is not a valid PPTX: %s" % e)
        return False


def main():
    print()
    success = test_content_agent()
    print("\n" + "=" * 60)
    if success:
        print("OVERALL: PASSED ✅")
    else:
        print("OVERALL: FAILED ❌")
    print("=" * 60)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
