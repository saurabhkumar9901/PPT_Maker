import os
import re

file_path = "compiler.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix Slide 6 Divider
# The new logic for divider:
# 1. Do NOT force font color. Let python-pptx inherit template color.
# 2. If len(placeholders) < 2 and there is a subtitle, manually draw it!

old_divider_logic = r"""        # ─── Divider ───────────────────────────────────────────
        if layout_type == 'divider':
            placeholders = sorted(list(slide.placeholders), key=lambda x: x.placeholder_format.idx)
            if len(placeholders) > 0:
                placeholders[0].text = slide_def.get('title', '')
                for p in placeholders[0].text_frame.paragraphs:
                    p.font.name = tokens\['fonts'\]\['heading'\]
                    p.font.size = Pt\(36\)
                    p.font.bold = True
                    p.alignment = PP_ALIGN.CENTER
                    if p.runs: p.runs\[0\].font.color.rgb = hex_to_rgb\(tokens\['colors'\]\['lt1'\]\)
            
            if len\(placeholders\) > 1 and slide_def.get\('subtitle'\):
                placeholders\[1\].text = slide_def.get\('subtitle', ''\)
                for p in placeholders\[1\].text_frame.paragraphs:
                    p.font.name = tokens\['fonts'\]\['body'\]
                    p.font.size = Pt\(20\)
                    p.alignment = PP_ALIGN.CENTER
                    if p.runs: p.runs\[0\].font.color.rgb = hex_to_rgb\(tokens\['colors'\]\['accent2'\]\)
            
            add_footer\(slide, tokens, idx \+ 1, report_name\)"""

new_divider_logic = """        # ─── Divider ───────────────────────────────────────────
        if layout_type == 'divider':
            placeholders = sorted(list(slide.placeholders), key=lambda x: x.placeholder_format.idx)
            if len(placeholders) > 0:
                placeholders[0].text = slide_def.get('title', '')
                for p in placeholders[0].text_frame.paragraphs:
                    p.font.name = tokens['fonts']['heading']
                    p.font.size = Pt(36)
                    p.font.bold = True
                    p.alignment = PP_ALIGN.CENTER
                    # Removed hardcoded font color to inherit the Master Slide's native aesthetic.
            
            if slide_def.get('subtitle'):
                if len(placeholders) > 1:
                    placeholders[1].text = slide_def.get('subtitle', '')
                    for p in placeholders[1].text_frame.paragraphs:
                        p.font.name = tokens['fonts']['body']
                        p.font.size = Pt(22)
                        p.alignment = PP_ALIGN.CENTER
                else:
                    # Manually inject subtitle below title!
                    title_top = placeholders[0].top
                    title_h = placeholders[0].height
                    sub_y = title_top + title_h + Inches(0.5)
                    add_text_box(slide, margin_left, sub_y, tokens['dimensions']['width'] * 914400 - (margin_left*2), Inches(1.5),
                        slide_def.get('subtitle', ''), tokens['fonts']['body'], Pt(22), 
                        hex_to_rgb(tokens['colors']['accent2']), align=PP_ALIGN.CENTER, zero_margin=True)
                        
            add_footer(slide, tokens, idx + 1, report_name)"""

# I'll just use a more resilient replacement for the entire Divider block 
# by searching between '# ─── Divider' and 'add_footer...' 

def update_file():
    global content
    
    # regex search
    pattern = r'# ─── Divider ───────────────────────────────────────────.*?add_footer\(slide, tokens, idx \+ 1, report_name\)'
    content = re.sub(pattern, new_divider_logic, content, flags=re.DOTALL)
    
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)

update_file()
print("Divider slide patched!")
