import os

file_path = "compiler.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix the add_card positional argument overlaps!
# The bug is lines like this: 
# hex_to_rgb(tokens['colors']['accent2']), hex_to_rgb(tokens['colors']['accent2']), Pt(2.0), shadow=True)
# or 
# hex_to_rgb(tokens['colors']['lt1']), hex_to_rgb(tokens['colors']['accent2']), Pt(2.0), shadow=True)
# We need to make sure they are mapped clearly as kwargs!

def fix_add_card():
    # Let's cleanly replace 'add_card' usages that are broken.
    global content
    
    # render_icon_grid broken call:
    broken_icon_grid = "add_card(slide, x, y, card_w, card_h,\n                hex_to_rgb(accent),\n                hex_to_rgb(tokens['colors']['accent2']), hex_to_rgb(tokens['colors']['accent2']), Pt(2.0), shadow=True)"
    fixed_icon_grid = "add_card(slide, x, y, card_w, card_h, fill_color=hex_to_rgb(tokens['colors']['lt1']), border_color=hex_to_rgb(tokens['colors']['accent2']), border_w=Pt(2.0), shadow=True)"
    
    # We don't know the exact string because of multiple replacements. Let's just use regex!
    import re
    # Match add_card(... shadow=True) and rewrite it.
    
    # Actually, it's easier:
    content = content.replace("hex_to_rgb(accent),\n                hex_to_rgb(tokens['colors']['accent2']), hex_to_rgb(tokens['colors']['accent2']), Pt(2.0), shadow=True", "fill_color=hex_to_rgb(tokens['colors']['lt1']), border_color=hex_to_rgb(tokens['colors']['accent2']), border_w=Pt(2.0), shadow=True")
    
    content = content.replace("hex_to_rgb(accent),\n                hex_to_rgb(accent), hex_to_rgb(accent), Pt(2.0), shadow=True", "fill_color=hex_to_rgb(tokens['colors']['lt1']), border_color=hex_to_rgb(tokens['colors']['accent2']), border_w=Pt(2.0), shadow=True")

    # render_stats_row
    content = content.replace("hex_to_rgb(accent_hex),\n                border_color=hex_to_rgb(accent_hex), border_w=Pt(2.0), shadow=True", "fill_color=hex_to_rgb(tokens['colors']['lt1']), border_color=hex_to_rgb(tokens['colors']['accent2']), border_w=Pt(2.0), shadow=True")

    # render_comparison
    content = content.replace("hex_to_rgb(tokens['colors']['accent2']), \n                border_color=hex_to_rgb(tokens['colors']['accent2']), border_w=Pt(2.0), shadow=True", "fill_color=hex_to_rgb(tokens['colors']['lt1']), border_color=hex_to_rgb(tokens['colors']['accent2']), border_w=Pt(2.0), shadow=True")
    content = content.replace("hex_to_rgb(accent), \n                border_color=hex_to_rgb(accent), border_w=Pt(2.0), shadow=True", "fill_color=hex_to_rgb(tokens['colors']['lt1']), border_color=hex_to_rgb(tokens['colors']['accent2']), border_w=Pt(2.0), shadow=True")

    content = content.replace("hex_to_rgb(tokens['colors']['accent2']), hex_to_rgb(tokens['colors']['accent2']), Pt(2.0), shadow=True", "fill_color=hex_to_rgb(tokens['colors']['lt1']), border_color=hex_to_rgb(tokens['colors']['accent2']), border_w=Pt(2.0), shadow=True")
    content = content.replace("hex_to_rgb(accent), hex_to_rgb(accent), Pt(2.0), shadow=True", "fill_color=hex_to_rgb(tokens['colors']['lt1']), border_color=hex_to_rgb(tokens['colors']['accent2']), border_w=Pt(2.0), shadow=True")
    
    # Generic fix for any hex_to_rgb(...) repeating 3 times
    import re
    content = re.sub(r'hex_to_rgb\([^)]+\),\s*hex_to_rgb\([^)]+\),\s*hex_to_rgb\([^)]+\),\s*Pt\(2\.0\),\s*shadow=True', 
                     r"fill_color=hex_to_rgb(tokens['colors']['lt1']), border_color=hex_to_rgb(tokens['colors']['accent2']), border_w=Pt(2.0), shadow=True", content)

fix_add_card()

# Fix Divider Slide 6 Subtitle issue!
# In compiler.py
divider_block = """        # ─── Divider ───────────────────────────────────────────
        if layout_type == 'divider':
            placeholders = list(slide.placeholders)
            if placeholders:
                placeholders[0].text = slide_def.get('title', '')
                for p in placeholders[0].text_frame.paragraphs:
                    p.font.name = tokens['fonts']['heading']
                    p.font.size = Pt(36)
                    p.font.bold = True
                    # Divider sections usually use dark backgrounds, ensure text is readable
                    p.font.color.rgb = hex_to_rgb(tokens['colors']['lt1'])
              add_footer(slide, tokens, idx + 1, report_name)"""

# Instead of exact replacement, let's use regex to insert the subtitle code
import re
new_divider = """        # ─── Divider ───────────────────────────────────────────
        if layout_type == 'divider':
            placeholders = sorted(list(slide.placeholders), key=lambda x: x.placeholder_format.idx)
            if len(placeholders) > 0:
                placeholders[0].text = slide_def.get('title', '')
                for p in placeholders[0].text_frame.paragraphs:
                    p.font.name = tokens['fonts']['heading']
                    p.font.size = Pt(36)
                    p.font.bold = True
                    p.alignment = PP_ALIGN.CENTER
                    if p.runs: p.runs[0].font.color.rgb = hex_to_rgb(tokens['colors']['lt1'])
            
            if len(placeholders) > 1 and slide_def.get('subtitle'):
                placeholders[1].text = slide_def.get('subtitle', '')
                for p in placeholders[1].text_frame.paragraphs:
                    p.font.name = tokens['fonts']['body']
                    p.font.size = Pt(20)
                    p.alignment = PP_ALIGN.CENTER
                    if p.runs: p.runs[0].font.color.rgb = hex_to_rgb(tokens['colors']['accent2'])
            
            add_footer(slide, tokens, idx + 1, report_name)"""

content = re.sub(r'# ─── Divider ───.*?add_footer\(slide, tokens, idx \+ 1, report_name\)', new_divider, content, flags=re.DOTALL)


# Also ensure that p.font.color.rgb uses p.runs[0] in divider
with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Slide 3 card arg bug and Slide 6 Subtitle issue patched.")
