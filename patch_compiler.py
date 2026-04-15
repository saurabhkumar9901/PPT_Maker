import os

file_path = "compiler.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: Faded colors
# Throughout compiler.py, there are calls like:
# add_card(slide, x, y, w, h, hex_to_rgb(accent_hex), border_color=..., alpha=10000...)
# Let's replace the fill color with lt2 to make it solid light, and keep the border solid accent.
content = content.replace("hex_to_rgb(accent_hex),\n                border_color=hex_to_rgb(accent_hex), border_w=Pt(1.5), alpha=10000", "hex_to_rgb(tokens['colors']['lt2']),\n                border_color=hex_to_rgb(accent_hex), border_w=Pt(2.0)")

content = content.replace("hex_to_rgb(accent_hex), border_color=None, alpha=10000", "hex_to_rgb(tokens['colors']['lt2']), border_color=hex_to_rgb(accent_hex), border_w=Pt(2.0)")

content = content.replace("hex_to_rgb(accent), \n                border_color=hex_to_rgb(accent), border_w=Pt(1.5), alpha=15000", "hex_to_rgb(tokens['colors']['lt2']), \n                border_color=hex_to_rgb(accent), border_w=Pt(2.0)")

content = content.replace("hex_to_rgb(accent), Pt(1.5), alpha=15000", "hex_to_rgb(tokens['colors']['lt2']), hex_to_rgb(accent), Pt(2.0), shadow=True")

content = content.replace("hex_to_rgb(colors[idx]),\n                border_color=hex_to_rgb(colors[idx]), border_w=Pt(2), alpha=15000", "hex_to_rgb(tokens['colors']['lt2']),\n                border_color=hex_to_rgb(colors[idx]), border_w=Pt(2), shadow=True")

content = content.replace("alpha=50000", "") # Remove generic heavy alphas
content = content.replace("alpha=90000", "")

# Fix 2: Slide 12 Cycle Overflow
old_cycle = """def render_cycle(slide, element, tokens):
    steps = element.get('steps', [])
    if not steps: return
    n = len(steps)
    center_x = int(MARGIN_L) + int(CONTENT_W / 2)
    center_y = int(CONTENT_TOP) + int(CONTENT_H / 2)
    radius = min(int(CONTENT_W), int(CONTENT_H)) / 2 - Inches(0.70)
    node_w = Inches(1.80)
    node_h = Inches(0.80)"""

new_cycle = """def render_cycle(slide, element, tokens):
    from content_fitter import calculate_fit_font_size
    steps = element.get('steps', [])
    if not steps: return
    n = len(steps)
    center_x = int(MARGIN_L) + int(CONTENT_W / 2)
    center_y = int(CONTENT_TOP) + int(CONTENT_H / 2)
    radius = min(int(CONTENT_W), int(CONTENT_H)) / 2 - Inches(1.20)
    node_w = Inches(2.60)
    node_h = Inches(1.80)"""

content = content.replace(old_cycle, new_cycle)

old_fonts = """        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = hex_to_rgb(tokens['colors']['lt1'])
        p.alignment = PP_ALIGN.CENTER

        if step.get('description'):
            p2 = tf.add_paragraph()
            r2 = p2.add_run()
            r2.text = step['description']
            r2.font.name = tokens['fonts']['body']
            r2.font.size = Pt(8)"""

new_fonts = """        
        # dynamic fitting
        fit_title = calculate_fit_font_size(step.get('title',''), node_w/914400, node_h/914400, 11)
        
        run.font.size = Pt(fit_title)
        run.font.bold = True
        run.font.color.rgb = hex_to_rgb(tokens['colors']['lt1'])
        p.alignment = PP_ALIGN.CENTER

        if step.get('description'):
            p2 = tf.add_paragraph()
            r2 = p2.add_run()
            r2.text = step['description']
            r2.font.name = tokens['fonts']['body']
            r2.font.size = Pt(max(7, fit_title - 2))"""

content = content.replace(old_fonts, new_fonts)

# Validate autofixer logic.
with open("auto_fixer.py", "r") as f:
    fix_content = f.read()

# Auto-fixer failsafe: ensure it dynamically shrinks overflowing node text bounds too!
# Currently Rule 3 uses left + width > slide_w. That's good.
# Let's write the modified compiler back.
with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Patch applied.")
