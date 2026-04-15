import os

file_path = "compiler.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: Divider Slide Text Color (Slide 6 fix)
# In compiler.py under # ─── Divider ───
divider_block_old = """        # ─── Divider ───────────────────────────────────────────
        if layout_type == 'divider':
            placeholders = list(slide.placeholders)
            if placeholders:
                placeholders[0].text = slide_def.get('title', '')
                for p in placeholders[0].text_frame.paragraphs:
                    p.font.name = tokens['fonts']['heading']
                    p.font.size = Pt(28)
                    p.font.bold = True
                    p.font.color.rgb = hex_to_rgb(tokens['colors']['dk1'])"""

divider_block_new = """        # ─── Divider ───────────────────────────────────────────
        if layout_type == 'divider':
            placeholders = list(slide.placeholders)
            if placeholders:
                placeholders[0].text = slide_def.get('title', '')
                for p in placeholders[0].text_frame.paragraphs:
                    p.font.name = tokens['fonts']['heading']
                    p.font.size = Pt(36)
                    p.font.bold = True
                    # Divider sections usually use dark backgrounds, ensure text is readable
                    p.font.color.rgb = hex_to_rgb(tokens['colors']['lt1'])"""

content = content.replace(divider_block_old, divider_block_new)

# Fix 2: "Faded" boxes and lines (Slide 3 etc)
# Substitute accent1 with accent2 across layout renderers for borders, line traces, badges etc.
content = content.replace("tokens['colors']['accent1']", "tokens['colors']['accent2']")
# Also there are some generic occurrences like 'accent = tokens['colors']['accent1']' which will be updated by the above.
content = content.replace("hex_to_rgb(tokens['colors']['accent1'])", "hex_to_rgb(tokens['colors']['accent2'])")

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Slide 3 and 6 color patches applied.")
