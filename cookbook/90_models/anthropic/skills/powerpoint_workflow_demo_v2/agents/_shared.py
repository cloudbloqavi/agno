"""
Shared instructions, Pydantic models, and constants used across all provider agent modules.

This file is the SINGLE SOURCE OF TRUTH for agent prompts and output schemas.
Provider modules (claude_agents.py, openai_agents.py, gemini_agents.py) import
from here so that instructions stay consistent regardless of which LLM provider
is selected at runtime.
"""

import os
import sys
from typing import List

from pydantic import BaseModel, Field

# ---------------------------------------------------------------------------
# Pydantic Models — imported by each provider module for output_schema
# ---------------------------------------------------------------------------

# Re-export from the main workflow files.  We import them lazily because the
# workflow files are large and import-heavy.  For the models needed at import
# time (BrandStyleIntent, ImagePlan, SlideQualityReport) we define slim copies
# here so the agents/ package can be imported without pulling in the full
# workflow chain.
#
# These classes MUST stay in sync with the canonical definitions in:
#   - powerpoint_chunked_workflow.py  (BrandStyleIntent)
#   - powerpoint_template_workflow.py (ImagePlan, SlideQualityReport, etc.)


class BrandStyleIntent(BaseModel):
    """Parsed branding/styling intent extracted from a user query or template file."""

    has_branding: bool = Field(
        False,
        description=(
            "True when the user query contains an identifiable branding or styling "
            "intent (e.g. 'Nike branding', 'in the style of Apple').  False when "
            "the query is purely topical with no brand/style directive."
        ),
    )
    brand_name: str = Field(
        "",
        description="Brand name extracted from query (e.g. 'Nike', 'Tesla').",
    )
    style_keywords: List[str] = Field(
        default_factory=list,
        description=(
            "Style descriptors inferred or researched for the brand "
            "(e.g. ['bold', 'sporty', 'minimalist'])."
        ),
    )
    color_palette: List[str] = Field(
        default_factory=list,
        description=(
            "Specific color names or hex codes associated with the brand "
            "(e.g. ['#FF6600', 'black', 'white']).  Prefer hex when known."
        ),
    )
    tone_override: str = Field(
        "",
        description=(
            "Tone suggested by the brand identity "
            "(e.g. 'empowering', 'innovative', 'luxurious')."
        ),
    )
    typography_hints: List[str] = Field(
        default_factory=list,
        description=(
            "Font families or typographic style hints associated with the brand "
            "(e.g. ['Futura', 'Helvetica Neue', 'sans-serif bold'])."
        ),
    )
    content_query: str = Field(
        "",
        description=(
            "The user's original query with branding clauses removed, preserving "
            "only the content/topic portion.  Empty when has_branding is False."
        ),
    )
    source: str = Field(
        "query",
        description="'query' or 'template' — where the intent was derived.",
    )
    source_detail: str = Field(
        "",
        description=(
            "Human-readable detail about the source "
            "(e.g. template filename, or 'user query')."
        ),
    )


class SlideImageDecision(BaseModel):
    """Decision about whether a slide needs an AI-generated image."""

    slide_index: int = Field(description="Zero-based index of the slide")
    needs_image: bool = Field(
        description="Whether this slide should have an AI-generated image"
    )
    image_prompt: str = Field(
        default="",
        description=(
            "Detailed prompt for image generation. Required if needs_image is True. "
            "If needs_image is False, leave empty."
        ),
    )
    reasoning: str = Field(
        default="",
        description="Brief explanation of why the slide does or does not need an image",
    )


class ImagePlan(BaseModel):
    """Plan for which slides need AI-generated images."""

    decisions: List[SlideImageDecision] = Field(
        description="List of image decisions, one per slide"
    )


class ShapeIssue(BaseModel):
    """A single design issue detected on a rendered slide image."""

    issue_type: str = Field(
        description=(
            "Category of issue: 'text_overflow', 'overlap', 'ghost_text', "
            "'low_contrast', 'element_clipped', 'empty_placeholder', "
            "'footer_inconsistent', 'poor_spacing', 'alignment_off', "
            "'typography_hierarchy', 'color_underutilized', "
            "'visual_enrichment_needed', 'font_inconsistency'"
        )
    )
    severity: str = Field(
        description="Severity level: 'critical', 'moderate', or 'minor'"
    )
    description: str = Field(
        description="Detailed, specific description of the issue"
    )
    programmatic_fix: str = Field(
        default="none",
        description=(
            "Recommended programmatic fix type: 'reduce_font_size', "
            "'increase_contrast', 'remove_element', 'clear_placeholder', "
            "'fix_spacing', 'fix_alignment', 'fix_body_paragraph_alignment', "
            "'enforce_typography_hierarchy', 'increase_title_font_size', "
            "'apply_accent_color_title', 'apply_accent_color_body', "
            "'apply_body_accent_border', 'enrich_header_bar', 'enrich_title_card', "
            "'enrich_divider', 'enrich_accent_strip', or 'none'"
        ),
    )
    shape_description: str = Field(
        default="",
        description=(
            "Brief description of the affected element (e.g., 'title text box', "
            "'bullet list', 'slide background'). Leave empty if slide-wide."
        ),
    )


class SlideQualityReport(BaseModel):
    """Design and quality assessment for a single rendered slide image."""

    slide_index: int = Field(description="Zero-based index of the slide")
    design_score: int = Field(
        description="Overall design quality score from 1-10"
    )
    is_visually_bland: bool = Field(
        default=False,
        description=(
            "True if the slide fails to use the template's visual vocabulary "
            "(monochrome text, no accent colors, no visual hierarchy)."
        ),
    )
    issues: List[ShapeIssue] = Field(
        default_factory=list,
        description=(
            "List of detected issues ordered by severity. Include both structural "
            "defects and design quality improvements. Be specific and actionable."
        ),
    )


# ---------------------------------------------------------------------------
# Shared Instruction Constants
# ---------------------------------------------------------------------------

BRAND_STYLE_ANALYZER_INSTRUCTIONS = [
    "Analyze the user's presentation request for branding or styling directives.",
    "Look for patterns like: 'using X branding', 'X-branded', 'in the style of X', "
    "'with X theme', 'X corporate style', 'like X would present it'.",
    "If NO branding intent is found, return has_branding=false with all other fields empty.",
    "If branding IS found, extract the brand_name and decide:",
    "  - For well-known global brands (Nike, Apple, Google, etc.) where you are "
    "    confident about colors and tone, you MAY skip web search.",
    "  - For less familiar brands, regional brands, or when you are unsure about "
    "    specific hex colors or typography, USE web search to look up "
    "    '<brand_name> brand guidelines colors typography'.",
    "Fill in color_palette with specific hex codes when possible.",
    "Fill in content_query with the user's query MINUS the branding clause.",
    "Keep style_keywords to 3-5 concise descriptors.",
    "Keep typography_hints to 1-3 font families.",
    "Set source='query' and source_detail='user query'.",
]

IMAGE_PLANNER_INSTRUCTIONS = [
    "You are an image planning specialist for PowerPoint presentations.",
    "You will receive slide metadata AND the user's presentation topic/request.",
    "",
    "RULES (follow strictly):",
    "- If has_image_placeholder is true: ALWAYS set needs_image to true.",
    "- Title slides (index 0): ALWAYS generate an image — first impressions matter.",
    "- Slides with existing images: NEVER generate (already have visuals).",
    "- Data slides (has_table, has_chart, or has_data_vis is true): NEVER generate an image.",
    "  These slides already contain native data visualizations (charts, tables, or infographics).",
    "  Adding an external AI-generated image would collide with the native visual and degrade clarity.",
    "- Slides where visual_suggestion contains 'chart', 'table', 'infographic', 'diagram', or 'graph':",
    "  NEVER generate an image. These slides will have native data visualizations from python-pptx,",
    "  not external images. Setting needs_image=true here would insert an unrelated photo/illustration",
    "  that clashes with the chart or table.",
    "- All other content slides: Default to YES. Visuals enhance every presentation.",
    "- If the user explicitly requested 'visuals', 'images', or 'with pictures',",
    "  generate images for at LEAST half of ELIGIBLE (non-data-vis) slides.",
    "",
    "IMPORTANT: When in doubt about non-data-vis slides, generate an image. It is better",
    "to have too many images than too few. Empty picture placeholders look unprofessional.",
    "But NEVER add images to data-vis slides (charts, tables, infographics, diagrams).",
    "",
    "When writing image prompts:",
    "- Use the PRESENTATION TOPIC to create relevant imagery.",
    "- Describe professional, clean, modern illustrations.",
    "- Use abstract or metaphorical imagery, not literal text depictions.",
    "- Specify style: 'minimalist corporate illustration', 'flat design', etc.",
    "- Keep prompts under 100 words.",
    "- Make images suitable for a professional business presentation.",
]

PPTX_CODE_GEN_INSTRUCTIONS = [
    "You are a Python code generator that creates PowerPoint presentations using python-pptx.",
    "When given slide specifications, write a COMPLETE, SELF-CONTAINED Python script that generates a .pptx file.",
    "The script must import all required libraries at the top.",
    "ALLOWED imports only: pptx, pptx.util, pptx.chart.data, pptx.enum.chart, pptx.dml.color, pptx.enum.text, io, os, os.path, collections, math.",
    "FORBIDDEN imports: matplotlib, matplotlib.pyplot, subprocess, socket, requests, urllib, httpx, shutil, glob, sys, importlib, __import__.",
    "For each slide, create one slide in the presentation using prs.slides.add_slide(prs.slide_layouts[N]).",
    "Slide layout indices: 0=Title Slide, 1=Title and Content, 2=Section Header, 3=Two Content.",
    "For CHARTS: ALWAYS use python-pptx native CategoryChartData or ChartData (creates editable Office charts). NEVER use matplotlib, PIL, or any image-based approach for charts.",
    "For python-pptx ChartData bar/column charts: from pptx.chart.data import ChartData; from pptx.enum.chart import XL_CHART_TYPE.",
    "ChartData example: chart_data = ChartData(); chart_data.categories = ['A','B','C']; chart_data.add_series('Series1', (10, 20, 30)); slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(1), Inches(1.5), Inches(8), Inches(4.5), chart_data).",
    "For line charts use XL_CHART_TYPE.LINE, for pie charts use XL_CHART_TYPE.PIE — always via ChartData, never via image embedding.",
    "For TABLES: ALWAYS use slide.shapes.add_table(rows, cols, Inches(1), Inches(1.5), Inches(8), Inches(4.5)). Fill cells via table.cell(row, col).text = 'value'. NEVER embed a table as an image or use matplotlib/PIL to render one.",
    "For INFOGRAPHICS or DIAGRAMS: use native python-pptx shapes (add_shape with MSO_AUTO_SHAPE_TYPE.RECTANGLE, add_textbox, add_connector) or a native table to approximate the layout. NEVER insert an image for an infographic or diagram.",
    "Treat charts/tables/infographics/diagrams as native data-vis. Preserve data-vis intent even when exact styling cannot be replicated.",
    "Synthesize plausible, specific data values from the visual_suggestion and key_points descriptions. Do NOT use generic placeholder data.",
    "CHART AXIS SAFETY: When styling chart axes, always wrap axis.format.line.fill.background(), axis.format.line.color.rgb, and axis.major_gridlines.format.line.color.rgb calls in try/except blocks to handle cases where the underlying XML element does not yet exist. Example: try: chart.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0x80,0x80,0x80)\\nexcept Exception: pass",
    "CHART AXIS SAFETY: Similarly wrap axis.tick_labels.font.color.rgb, chart.legend.font.color.rgb, and series.data_labels.font.color.rgb in try/except blocks — these can raise 'NoneType object has no attribute attrib' if the XML element is absent.",
    "Save the final presentation using prs.save('EXACT_OUTPUT_PATH') where EXACT_OUTPUT_PATH is the path given in the task.",
    "Execute the script using the save_to_file_and_run tool immediately after writing it.",
    "If the script has an error, fix it and re-run. Maximum 2 fix attempts.",
    "If a visual cannot be implemented exactly, keep the slide and add a concise native textbox note. Do not skip slides.",
    "Do not add speaker notes, animations, or transitions.",
    "Do not print to stdout or write any files other than the final prs.save() call.",
    "COLOR AND CONTRAST RULES:",
    "- Default slide background: white (#FFFFFF) or very light colors (luminance > 0.9).",
    "- Default body text: dark colors (#333333 or #000000) for readability.",
    "- If using a dark background (luminance < 0.3), use white (#FFFFFF) or very light text.",
    "- If using colored shape fills for headers/accents, ensure text color has sufficient contrast.",
    "- For charts: use medium-to-dark accent colors; data labels should contrast against their background.",
    "- For tables: header rows with dark fills should have white text; body rows with light fills should have dark text.",
    "- NEVER use dark text (#000000-#666666) on dark backgrounds (#000000-#555555).",
    "- NEVER use light text (#AAAAAA-#FFFFFF) on light backgrounds (#CCCCCC-#FFFFFF).",
    "- When in doubt, use white background with black text — readability is paramount.",
]

SLIDE_QUALITY_REVIEWER_INSTRUCTIONS = [
    "You are a world-class senior UI/UX designer and presentation design expert",
    "with 15+ years of experience creating award-winning corporate presentations.",
    "You combine the precision of a quality inspector with the creative eye of a",
    "professional designer who understands visual hierarchy, typography, color theory,",
    "whitespace, and the importance of brand consistency.",
    "",
    "Analyze the provided slide screenshot from BOTH a structural AND aesthetic",
    "design quality perspective. Your goal is to elevate every slide to the",
    "standard of a professionally designed McKinsey or Apple-quality presentation.",
    "",
    "=== STRUCTURAL DEFECTS (always report if present) ===",
    "text_overflow      - Text extends beyond its container or is cut off.",
    "overlap            - Two elements visibly overlap in a way that hurts readability.",
    "ghost_text         - 'Click to add' placeholder text is still visible.",
    "low_contrast       - Text is nearly indistinguishable from background.",
    "element_clipped    - Content is cut off by the slide boundary.",
    "empty_placeholder  - Visible empty frame with no content.",
    "footer_inconsistent - Footer text missing, truncated, or misaligned.",
    "",
    "=== DESIGN QUALITY ISSUES (report when they significantly impact visual quality) ===",
    "poor_spacing       - Insufficient whitespace, cramped layout, or unbalanced margins.",
    "alignment_off      - Elements are visibly misaligned (not on a consistent grid).",
    "typography_hierarchy - Title and body have similar weight/size, lacking visual hierarchy.",
    "color_underutilized - Template accent colors are not used; everything looks monochrome.",
    "visual_enrichment_needed - Slide uses only plain text when the template's design",
    "                           vocabulary (colored shapes, accent bars, icons) could",
    "                           dramatically improve visual interest.",
    "font_inconsistency - Inconsistent font sizes or weights across similar content types.",
    "",
    "=== SEVERITY GUIDE ===",
    "critical  - Broken or unreadable: text cut off, ghost text, major overlap.",
    "moderate  - Clearly suboptimal: a professional would notice and want to fix it.",
    "minor     - A refinement that would polish the slide.",
    "",
    "=== PROGRAMMATIC FIX SELECTION ===",
    "For structural issues:",
    "  reduce_font_size      -> text_overflow",
    "  increase_contrast     -> low_contrast",
    "  remove_element        -> overlap (remove the offending element)",
    "  clear_placeholder     -> ghost_text, empty_placeholder",
    "",
    "For spacing/alignment (safe, failsafe implementations):",
    "  fix_spacing                  -> poor_spacing: clamps shapes that overflow safe",
    "                                  margins back within 5% edge boundary.",
    "  fix_alignment                -> alignment_off: snaps outlier shapes to the",
    "                                  majority left edge (max 2% slide width).",
    "  fix_body_paragraph_alignment -> alignment_off / poor_spacing: sets all body",
    "                                  text paragraphs to left alignment.",
    "",
    "For typography hierarchy (ensures title is visually dominant):",
    "  enforce_typography_hierarchy -> typography_hierarchy: ensures title font is at",
    "                                  least 6pt larger than body. Increases title if",
    "                                  needed (cap: 36pt).",
    "  increase_title_font_size     -> typography_hierarchy: forces title to 28pt.",
    "",
    "For color scheme (applies template accent colors):",
    "  apply_accent_color_title -> color_underutilized / typography_hierarchy:",
    "                              applies primary accent color to title text runs.",
    "  apply_accent_color_body  -> color_underutilized: applies primary accent color",
    "                              to the first (lead) paragraph of body text.",
    "",
    "For visual enrichment (native PPTX shapes, NO image generation):",
    "  apply_body_accent_border  -> mild enrichment: thin vertical accent bar left of body.",
    "  enrich_header_bar    -> prominent: full-width accent bar across top 8% of slide.",
    "  enrich_title_card    -> structured: lightly tinted card behind title area.",
    "  enrich_divider       -> separator: thin horizontal accent rule at 25% height.",
    "  enrich_accent_strip  -> minimal: thin vertical accent strip on far left edge.",
    "",
    "SELECTION GUIDE:",
    "  poor_spacing                   -> fix_spacing",
    "  alignment_off (shapes)         -> fix_alignment",
    "  alignment_off (text)           -> fix_body_paragraph_alignment",
    "  typography_hierarchy (severe)  -> enforce_typography_hierarchy",
    "  typography_hierarchy (mild)    -> increase_title_font_size",
    "  color_underutilized (title)    -> apply_accent_color_title",
    "  color_underutilized (overall)  -> apply_accent_color_body + enrich_*",
    "  visual_enrichment_needed       -> choose the best enrich_* type",
    "",
    "Use 'none' ONLY when the issue requires AI-generated images, human content",
    "editing, or a completely different slide layout.",
    "",
    "=== DATA VISUALIZATION SLIDES (charts, tables, infographics) ===",
    "If the rendered slide clearly contains a native chart, table, or infographic",
    "(a data visualization element that fills the visual region), do NOT penalize it",
    "for lacking a photo or illustration. The data visual IS the intended element —",
    "it occupies the visual region on purpose. Specifically:",
    "- Do NOT set is_visually_bland=True purely because no AI-generated image is present.",
    "- Do NOT report visual_enrichment_needed with programmatic_fix='none' citing absence",
    "  of a photo/illustration on a slide that already has a chart or table.",
    "- Do NOT suggest adding an AI-generated photo to a slide that already has data visuals.",
    "You SHOULD still report genuine structural defects on these slides (text_overflow,",
    "ghost_text, overlap, low_contrast, element_clipped) and design issues unrelated to",
    "image absence (e.g. typography_hierarchy, poor_spacing, color_underutilized in text",
    "regions). When the prompt includes '[Data vis: chart/table/infographic present]',",
    "apply this rule strictly — the slide is intentionally structured around its data visual.",
    "",
    "=== VISUAL BLANDNESS ===",
    "A slide is visually bland (is_visually_bland=True) when it fails to use",
    "the template's visual vocabulary — for example: all text in one plain color,",
    "no accent colors applied anywhere, no visual hierarchy, large empty whitespace",
    "regions, or it looks like an unformatted draft document.",
    "Be generous with this flag: flag it if a professional designer would look at",
    "the slide and immediately want to add some visual structure or color accent.",
    "",
    "=== DESIGN SCORE (design_score field, 1-10) ===",
    "10: Stunning. Could appear in a top-tier consulting pitch deck.",
    "8-9: Strong design with minor polish opportunities.",
    "6-7: Functional and readable, could use visual enrichment.",
    "4-5: Plain / generic; lacks visual identity or hierarchy.",
    "2-3: Noticeably poorly designed or has significant issues.",
    "1: Broken or completely unusable.",
    "",
    "=== IMPORTANT ===",
    "- The template_context field in the prompt tells you the available accent",
    "  colors and fonts from the template. Reference these when suggesting fixes.",
    "- Be specific in descriptions: not 'title could be bigger' but 'title is 20pt,",
    "  indistinguishable from body text at 18pt; increase to 28pt for hierarchy'.",
    "- Always return design_score even if the slide looks perfect.",
]
