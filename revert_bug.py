import os

file_path = "compiler.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# I previously ran replacements like:
# content.replace("hex_to_rgb(accent_hex),\n                border_color=hex_to_rgb(accent_hex), border_w=Pt(1.5), alpha=10000", "hex_to_rgb(tokens['colors']['lt2']),\n                border_color=hex_to_rgb(accent_hex), border_w=Pt(2.0)")
# So I need to carefully restore 'lt2' back to 'accent_hex' or 'accent' or 'lt1'.

# Let's cleanly replace tokens['colors']['lt2'] with accent where appropriate.
# 1. render_stats_row (uses accent_hex)
content = content.replace("hex_to_rgb(tokens['colors']['lt2']),\n                border_color=hex_to_rgb(accent_hex), border_w=Pt(2.0)", "hex_to_rgb(accent_hex),\n                border_color=hex_to_rgb(accent_hex), border_w=Pt(2.0), shadow=True")

# 2. render_comparison / matrix
content = content.replace("hex_to_rgb(tokens['colors']['lt2']), \n                border_color=hex_to_rgb(accent), border_w=Pt(2.0)", "hex_to_rgb(accent), \n                border_color=hex_to_rgb(accent), border_w=Pt(2.0), shadow=True")

content = content.replace("hex_to_rgb(tokens['colors']['lt2']), hex_to_rgb(accent), Pt(2.0), shadow=True", "hex_to_rgb(accent), hex_to_rgb(accent), Pt(2.0), shadow=True")

# 3. render_matrix (uses colors[idx])
content = content.replace("hex_to_rgb(tokens['colors']['lt2']),\n                border_color=hex_to_rgb(colors[idx]), border_w=Pt(2), shadow=True", "hex_to_rgb(tokens['colors']['lt1']),\n                border_color=hex_to_rgb(colors[idx]), border_w=Pt(2), shadow=True")

# Fix add_card at line 1076 (which I didn't regex properly maybe, or I missed).
# Let's change any generic lt2 add_card backgrounds to lt1 (white) just in case! 
# lt2 is supposed to be light, but some templates map it to dark. lt1 is always white.
# In `add_card(..., hex_to_rgb(tokens['colors']['lt2'])` we should use lt1.
content = content.replace("hex_to_rgb(tokens['colors']['lt2'])", "hex_to_rgb(tokens['colors']['lt1'])")

# Also, the color of text in cycle
# run.font.color.rgb = hex_to_rgb(tokens['colors']['lt1']) -> white text.
# The dark background cycle nodes were green! But colors array uses accents! 
# colors = [tokens['colors']['accent1'], tokens['colors']['accent2'], tokens['colors']['accent3'], ...]
# If accent1 is light (#EFF3E5), then white text on it will be invisible!
# So for cycle nodes, cycle text should default to 'dk1' (black), not 'lt1' (white)!
old_cycle_font_1 = "run.font.color.rgb = hex_to_rgb(tokens['colors']['lt1'])"
new_cycle_font_1 = "run.font.color.rgb = hex_to_rgb(tokens['colors']['dk1'])"
content = content.replace(old_cycle_font_1, new_cycle_font_1)

# Save
with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Restoration complete.")
