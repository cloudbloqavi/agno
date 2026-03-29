---
name: theme-factory
description: Provides curated themes with colors and fonts to apply to PowerPoint slides. Also provides custom generative capabilities for specific brands.
license: Apache-2.0
---

# Theme Factory Skill

This skill provides a curated collection of professional font and color themes, each with carefully selected color palettes and font pairings. Once a theme is mapped, its structure is supplied to the PowerPoint code generator.

## When to Use

- User asks for a presentation in a specific branding style (Colors, Fonts).
- User mentions a brand (e.g., "Coca-Cola", "Stripe").
- You are requested to pick an appropriate style for a deck.

## Process

1. **Analyze Intent**: Review the provided `BrandStyleIntent` and user prompt. 
2. **List Available Themes**: Run the `scripts/list_themes.py` script to see an overview of the 10 built-in themes.
3. **Compare Constraints**: 
   - If the request allows for generic, cohesive aesthetics, read one of the matching themes via `get_skill_reference('theme-factory', 'references/theme-name.md')`.
   - If the request mentions a strict real-world brand (e.g., Coca-Cola's distinct Red/White/Black) AND no preset is a perfect 1:1 match, you MUST generate a completely new custom theme payload yourself.
4. **Output Final Payload**: Output the final valid JSON containing the `name`, `description`, `palette`, and `typography` exactly matching the reference style.

## Script Details

- `scripts/list_themes.py`: Returns a compressed JSON catalog of all standard available themes inside `references/`.

## Best Practices

- Always prioritize reading from standard `references/` via `get_skill_reference` if the intent is broad (like "modern dark theme").
- For specific brands, ensure custom generated JSON payloads have excellent contrast metrics.
