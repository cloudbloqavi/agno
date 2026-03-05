import re

with open("powerpoint_template_workflow.py", "r") as f:
    content = f.read()

# 1. Update _make_high_contrast_fill definition and logic
old_func = """def _make_high_contrast_fill(rPr, bg_hex: str):
    \"\"\"Set text color to black or white for maximum contrast against background.\"\"\"
    ns_a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    bg_rgb = _hex_to_rgb(bg_hex)
    bg_lum = _relative_luminance(*bg_rgb)
    # Use white text on dark backgrounds, black text on light backgrounds
    text_hex = "FFFFFF" if bg_lum < 0.4 else "000000"

    solid_fill = etree.SubElement(rPr, ns_a + "solidFill")
    srgb_clr = etree.SubElement(solid_fill, ns_a + "srgbClr")
    srgb_clr.set("val", text_hex)"""

new_func = """def _make_high_contrast_fill(rPr, bg_hex: str, existing_solidFill=None):
    \"\"\"Set text color to black or white for maximum contrast against background.\"\"\"
    ns_a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    bg_rgb = _hex_to_rgb(bg_hex)
    bg_lum = _relative_luminance(*bg_rgb)
    # Use white text on dark backgrounds, black text on light backgrounds
    text_hex = "FFFFFF" if bg_lum < 0.4 else "000000"

    if existing_solidFill is not None:
        # Clear existing children to preserve XML order of solidFill within rPr
        for child in list(existing_solidFill):
            existing_solidFill.remove(child)
        srgb_clr = etree.SubElement(existing_solidFill, ns_a + "srgbClr")
        srgb_clr.set("val", text_hex)
    else:
        # Create new solidFill and insert at beginning of rPr to avoid sequence violations
        solid_fill = etree.Element(ns_a + "solidFill")
        srgb_clr = etree.SubElement(solid_fill, ns_a + "srgbClr")
        srgb_clr.set("val", text_hex)
        rPr.insert(0, solid_fill)"""

if old_func in content:
    content = content.replace(old_func, new_func)
    print("Updated _make_high_contrast_fill function")
else:
    print("Could not find _make_high_contrast_fill function")

# 2. Update callers that remove solidFill first
# Pattern: rPr.remove(solidFill)\n {whitespace} _make_high_contrast_fill(rPr, bg_hex)
# Sometimes it's cell_bg_hex instead of bg_hex
pattern1 = r'rPr\.remove\(solidFill\)\n(\s+)_make_high_contrast_fill\(rPr, (bg_hex|cell_bg_hex)\)'
content, count = re.subn(pattern1, r'_make_high_contrast_fill(rPr, \2, existing_solidFill=solidFill)', content)
print(f"Replaced {count} instances of rPr.remove(solidFill)")

# 3. Update SubElement(run._r, ns_a + "rPr") to insert(0)
# Look for: new_rPr = _etree.SubElement(run._r, ns_a + "rPr")
pattern2 = r'new_rPr\s*=\s*_etree\.SubElement\(run\._r,\s*ns_a\s*\+\s*"rPr"\)'
replacement2 = 'new_rPr = _etree.Element(ns_a + "rPr")\n                                    run._r.insert(0, new_rPr)'
content, count2 = re.subn(pattern2, replacement2, content)
print(f"Replaced {count2} instances of SubElement rPr")

# 4. _set_chart_text_color function needs the same treatment to not use SubElement
old_chart = """def _set_chart_text_color(rPr_elem, ns_a: str, color_hex: str):
    \"\"\"Set or replace solidFill color on a chart rPr/defRPr element.\"\"\"
    existing = rPr_elem.find(ns_a + "solidFill")
    if existing is not None:
        rPr_elem.remove(existing)
    new_fill = etree.SubElement(rPr_elem, ns_a + "solidFill")
    srgb = etree.SubElement(new_fill, ns_a + "srgbClr")
    srgb.set("val", color_hex)"""

new_chart = """def _set_chart_text_color(rPr_elem, ns_a: str, color_hex: str):
    \"\"\"Set or replace solidFill color on a chart rPr/defRPr element.\"\"\"
    existing = rPr_elem.find(ns_a + "solidFill")
    if existing is not None:
        for child in list(existing):
            existing.remove(child)
        srgb = etree.SubElement(existing, ns_a + "srgbClr")
        srgb.set("val", color_hex)
    else:
        new_fill = etree.Element(ns_a + "solidFill")
        srgb = etree.SubElement(new_fill, ns_a + "srgbClr")
        srgb.set("val", color_hex)
        rPr_elem.insert(0, new_fill)"""

if old_chart in content:
    content = content.replace(old_chart, new_chart)
    print("Updated _set_chart_text_color function")
else:
    print("Could not find _set_chart_text_color function")

# 5. Fix overlapping charts
# Update enforce_final_contrast to stagger overlapping charts
# we need to track shapes resized
overlap_fix = """
    for slide in prs.slides:
        resized_charts = []
        for shape in slide.shapes:"""

old_loop = """
    for slide in prs.slides:
        for shape in slide.shapes:"""

content = content.replace(old_loop, overlap_fix)

overlap_logic_old = """                    if shape.width < _MIN_W:
                        shape.width = _MIN_W
                        corrections += 1
                    if shape.height < _MIN_H:
                        shape.height = _MIN_H
                        corrections += 1"""
overlap_logic_new = """                    if shape.width < _MIN_W:
                        shape.width = _MIN_W
                        corrections += 1
                    if shape.height < _MIN_H:
                        shape.height = _MIN_H
                        corrections += 1
                    
                    # Prevent overlap
                    for pc in resized_charts:
                        if abs(shape.left - pc.left) < _Inches(2.0) and abs(shape.top - pc.top) < _Inches(2.0):
                            shape.top = pc.top + pc.height + _Inches(0.2)
                            corrections += 1
                    resized_charts.append(shape)"""

if overlap_logic_old in content:
    content = content.replace(overlap_logic_old, overlap_logic_new)
    print("Updated overlap logic")
else:
    print("Could not find overlap logic")


with open("powerpoint_template_workflow.py", "w") as f:
    f.write(content)

