#!/usr/bin/env python3
"""
List all available theme presets.
Returns a JSON array of themes with their basic configuration.
"""

import json
import os
import sys
from pathlib import Path


def list_themes():
    # Execute relative to the location of this script
    script_dir = Path(__file__).resolve().parent
    references_dir = script_dir.parent / "references"

    if not references_dir.exists():
        return json.dumps(
            {"error": f"References directory not found at {references_dir}"}
        )

    themes = []

    for file in references_dir.glob("*.md"):
        try:
            content = file.read_text(encoding="utf-8")
            lines = content.splitlines()
            name = (
                lines[0].replace("# ", "").strip()
                if lines and lines[0].startswith("# ")
                else file.stem
            )

            # Simple summarization for the LLM
            content_str = str(content)
            themes.append(
                {"id": file.stem, "name": name, "summary": content_str[:250] + "..."}
            )
        except Exception as e:
            pass

    return json.dumps({"themes": themes}, indent=2)


if __name__ == "__main__":
    print(list_themes())
