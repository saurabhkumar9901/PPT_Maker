"""
Content Fitting Engine for PPT Maker v2.

Provides intelligent text fitting within placeholder and shape bounds:
  - Measures text against available space
  - Calculates optimal font sizes to prevent overflow
  - Splits overflowing content into continuation slides

Ported and adapted from pptx-from-layouts-skill.
"""

import math
from pptx.util import Pt, Emu, Inches


# ─── Font Metrics (approximate) ──────────────────────────────────────────────
# Since we can't render fonts server-side, use conservative character-width
# estimates based on common proportional fonts at various sizes.

# Average character width as fraction of font size (em units)
# Calibri/Arial average ~0.52em per character
AVG_CHAR_WIDTH_RATIO = 0.52

# Average line height as fraction of font size
LINE_HEIGHT_RATIO = 1.35


def estimate_text_lines(text: str, font_size_pt: float, box_width_inches: float) -> int:
    """Estimate how many lines a block of text will occupy in a given box width.
    
    Args:
        text: The text content
        font_size_pt: Font size in points
        box_width_inches: Width of the text container in inches
    
    Returns:
        Estimated number of lines
    """
    if not text:
        return 0
    
    # Convert box width to points (1 inch = 72 points)
    box_width_pt = box_width_inches * 72
    
    # Average character width at this font size
    char_width_pt = font_size_pt * AVG_CHAR_WIDTH_RATIO
    
    # Characters per line
    chars_per_line = max(1, int(box_width_pt / char_width_pt))
    
    # Count lines (accounting for word wrap — words don't break mid-word)
    total_lines = 0
    for paragraph in text.split('\n'):
        if not paragraph.strip():
            total_lines += 1
            continue
        words = paragraph.split()
        current_line_len = 0
        lines_in_para = 1
        for word in words:
            word_len = len(word) + 1  # +1 for space
            if current_line_len + word_len > chars_per_line and current_line_len > 0:
                lines_in_para += 1
                current_line_len = len(word)
            else:
                current_line_len += word_len
        total_lines += lines_in_para
    
    return total_lines


def calculate_fit_font_size(text: str, box_width_inches: float, box_height_inches: float,
                             max_font_pt: float = 16, min_font_pt: float = 8,
                             line_spacing: float = 1.35) -> float:
    """Calculate the largest font size that fits text within a bounding box.
    
    Uses binary search to find the optimal font size between min and max.
    
    Args:
        text: The text to fit
        box_width_inches: Available width in inches
        box_height_inches: Available height in inches
        max_font_pt: Maximum allowed font size
        min_font_pt: Minimum allowed font size (below this, text is unreadable)
        line_spacing: Line height multiplier
    
    Returns:
        Optimal font size in points
    """
    if not text:
        return max_font_pt
    
    box_height_pt = box_height_inches * 72
    
    # Binary search for the best font size
    low, high = min_font_pt, max_font_pt
    best = min_font_pt
    
    for _ in range(10):  # 10 iterations gives ~0.05pt precision
        mid = (low + high) / 2
        num_lines = estimate_text_lines(text, mid, box_width_inches)
        total_height = num_lines * mid * line_spacing
        
        if total_height <= box_height_pt:
            best = mid
            low = mid
        else:
            high = mid
    
    return round(best, 1)


def calculate_bullet_fit(items: list, box_width_inches: float, box_height_inches: float,
                          max_font_pt: float = 12, min_font_pt: float = 8,
                          spacing_pt: float = 6) -> float:
    """Calculate optimal font size for a bulleted list.
    
    Args:
        items: List of bullet text strings or dicts with 'text' key
        box_width_inches: Available width
        box_height_inches: Available height
        max_font_pt: Maximum font size
        min_font_pt: Minimum font size
        spacing_pt: Space between bullets in points
    
    Returns:
        Optimal font size in points
    """
    if not items:
        return max_font_pt
    
    # Flatten items to strings
    texts = []
    for item in items:
        if isinstance(item, str):
            texts.append(item)
        elif isinstance(item, dict):
            prefix = item.get('bold_prefix', '')
            text = item.get('text', '')
            texts.append(f"{prefix} {text}" if prefix else text)
    
    combined = '\n'.join(texts)
    box_height_pt = box_height_inches * 72
    
    low, high = min_font_pt, max_font_pt
    best = min_font_pt
    
    for _ in range(10):
        mid = (low + high) / 2
        total_height = 0
        for t in texts:
            lines = estimate_text_lines(t, mid, box_width_inches)
            total_height += lines * mid * LINE_HEIGHT_RATIO + spacing_pt
        
        if total_height <= box_height_pt:
            best = mid
            low = mid
        else:
            high = mid
    
    return round(best, 1)


# ─── Content Splitting ────────────────────────────────────────────────────────

def detect_content_overflow(items: list, box_width_inches: float, box_height_inches: float,
                             font_size_pt: float) -> bool:
    """Check if a list of content items would overflow the given box.
    
    Args:
        items: List of text strings or dicts
        box_width_inches: Available width
        box_height_inches: Available height
        font_size_pt: Font size being used
    
    Returns:
        True if content would overflow
    """
    box_height_pt = box_height_inches * 72
    total_height = 0
    
    for item in items:
        text = item if isinstance(item, str) else item.get('text', '')
        lines = estimate_text_lines(text, font_size_pt, box_width_inches)
        total_height += lines * font_size_pt * LINE_HEIGHT_RATIO + 6  # 6pt spacing
    
    return total_height > box_height_pt


def split_content_for_overflow(items: list, box_width_inches: float, box_height_inches: float,
                                font_size_pt: float) -> list:
    """Split a list of items into chunks that each fit within the box.
    
    Args:
        items: List of text strings or dicts
        box_width_inches: Available width  
        box_height_inches: Available height
        font_size_pt: Font size being used
    
    Returns:
        List of lists, where each sub-list fits within the box
    """
    if not items:
        return [[]]
    
    box_height_pt = box_height_inches * 72
    chunks = []
    current_chunk = []
    current_height = 0
    
    for item in items:
        text = item if isinstance(item, str) else item.get('text', '')
        lines = estimate_text_lines(text, font_size_pt, box_width_inches)
        item_height = lines * font_size_pt * LINE_HEIGHT_RATIO + 6
        
        if current_height + item_height > box_height_pt and current_chunk:
            chunks.append(current_chunk)
            current_chunk = []
            current_height = 0
        
        current_chunk.append(item)
        current_height += item_height
    
    if current_chunk:
        chunks.append(current_chunk)
    
    return chunks if chunks else [[]]
