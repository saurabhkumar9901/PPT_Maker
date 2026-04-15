import os

file_path = "compiler.py"
with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace accent2 globally with accent1 to remove the grey overrides
content = content.replace("tokens['colors']['accent2']", "tokens['colors']['accent1']")

# Now safely re-introduce secondary, tertiary colors properly in the loops!
content = content.replace(
    "colors = [tokens['colors']['accent1'], tokens['colors']['accent1'], tokens['colors']['accent3'],",
    "colors = [tokens['colors']['accent1'], tokens['colors']['accent2'], tokens['colors']['accent3'],"
)
content = content.replace(
    "colors = [tokens['colors']['accent1'], tokens['colors']['accent1'],",
    "colors = [tokens['colors']['accent1'], tokens['colors']['accent2'],"
)

# This was the generic matrix colors line
content = content.replace(
    "colors = [tokens['colors']['accent1'], tokens['colors'].get('accent3', tokens['colors']['accent1']),\n              tokens['colors']['accent1'], tokens['colors']['accent4']]",
    "colors = [tokens['colors']['accent1'], tokens['colors'].get('accent2', tokens['colors']['accent1']),\n              tokens['colors'].get('accent3', tokens['colors']['accent1']), tokens['colors']['accent4']]"
)

# Line 528 fallback
content = content.replace(
    "colors = [tokens['colors'].get(f'accent{i}', tokens['colors']['accent1']) for i in range(1, 7)]",
    "colors = [tokens['colors'].get(f'accent{i}', tokens['colors']['accent1']) for i in range(1, 7)]"
)

# Other explicit restorations (if missing, it's harmless as accent1 is primary)
content = content.replace("accent1 if is_total", "accent2 if is_total")
content = content.replace("hex_to_rgb(tokens['colors']['lt1']), hex_to_rgb(tokens['colors']['accent1']), Pt(1),", "hex_to_rgb(tokens['colors']['lt1']), hex_to_rgb(tokens['colors']['accent2']), Pt(1),")


with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Restored accent1 mapping.")
