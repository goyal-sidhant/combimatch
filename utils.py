"""
utils.py - Utility Functions for CombiMatch

Contains helper functions for:
- Color generation and management
- Number parsing from text
- Input validation
"""

from typing import List, Tuple, Optional
from models import NumberItem, ItemSource


# Predefined colors for finalized combinations
# Soft, distinguishable colors that are easy on the eyes
HIGHLIGHT_COLORS: List[Tuple[Tuple[int, int, int], str]] = [
    ((173, 216, 230), "Light Blue"),
    ((144, 238, 144), "Light Green"),
    ((255, 218, 185), "Peach"),
    ((221, 160, 221), "Plum"),
    ((176, 224, 230), "Powder Blue"),
    ((255, 255, 224), "Light Yellow"),
    ((240, 128, 128), "Light Coral"),
    ((175, 238, 238), "Pale Turquoise"),
    ((255, 182, 193), "Light Pink"),
    ((216, 191, 216), "Thistle"),
    ((152, 251, 152), "Pale Green"),
    ((255, 228, 196), "Bisque"),
    ((230, 230, 250), "Lavender"),
    ((189, 183, 107), "Dark Khaki"),
    ((250, 250, 210), "Light Goldenrod"),
    ((135, 206, 235), "Sky Blue"),
    ((245, 222, 179), "Wheat"),
    ((188, 143, 143), "Rosy Brown"),
    ((176, 196, 222), "Light Steel Blue"),
    ((255, 240, 245), "Lavender Blush"),
]


class ColorManager:
    """
    Manages color assignment for finalized combinations.
    
    Cycles through predefined colors, ensuring each finalized
    combination gets a unique color.
    """
    
    def __init__(self):
        """Initialize the color manager."""
        self._current_index = 0
        self._assigned_colors: List[Tuple[Tuple[int, int, int], str]] = []
    
    def get_next_color(self) -> Tuple[Tuple[int, int, int], str]:
        """
        Get the next available color.
        
        Returns:
            Tuple of (RGB tuple, color name)
        """
        color = HIGHLIGHT_COLORS[self._current_index % len(HIGHLIGHT_COLORS)]
        self._assigned_colors.append(color)
        self._current_index += 1
        return color
    
    def reset(self):
        """Reset color assignment."""
        self._current_index = 0
        self._assigned_colors.clear()
    
    @property
    def colors_used(self) -> int:
        """Number of colors assigned so far."""
        return len(self._assigned_colors)


def parse_numbers_line_separated(text: str) -> Tuple[List[NumberItem], List[str]]:
    """
    Parse numbers from line-separated text.
    
    Args:
        text: Input text with one number per line
        
    Returns:
        Tuple of (list of NumberItem, list of error messages)
    """
    items = []
    errors = []
    
    lines = text.strip().split('\n')
    
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
        
        try:
            # Handle thousand separators and different decimal formats
            cleaned = line.replace(',', '').replace(' ', '')
            value = float(cleaned)
            
            item = NumberItem(
                value=value,
                index=len(items),
                source=ItemSource.MANUAL
            )
            items.append(item)
            
        except ValueError:
            errors.append(f"Line {i + 1}: '{line}' is not a valid number")
    
    return items, errors


def parse_numbers_comma_separated(text: str) -> Tuple[List[NumberItem], List[str]]:
    """
    Parse numbers from comma-separated text.
    
    Args:
        text: Input text with numbers separated by commas
        
    Returns:
        Tuple of (list of NumberItem, list of error messages)
    """
    items = []
    errors = []
    
    # Split by comma
    parts = text.strip().split(',')
    
    for i, part in enumerate(parts):
        part = part.strip()
        if not part:
            continue
        
        try:
            # Remove any thousand separators (spaces)
            cleaned = part.replace(' ', '')
            value = float(cleaned)
            
            item = NumberItem(
                value=value,
                index=len(items),
                source=ItemSource.MANUAL
            )
            items.append(item)
            
        except ValueError:
            errors.append(f"Item {i + 1}: '{part}' is not a valid number")
    
    return items, errors


def validate_parameters(
    target: str,
    buffer: str,
    min_size: str,
    max_size: str,
    max_results: str
) -> Tuple[Optional[dict], List[str]]:
    """
    Validate and parse search parameters.
    
    Args:
        target: Target sum string
        buffer: Buffer value string
        min_size: Minimum combo size string
        max_size: Maximum combo size string
        max_results: Maximum results string
        
    Returns:
        Tuple of (dict of parsed values or None, list of errors)
    """
    errors = []
    params = {}
    
    # Target sum
    try:
        params['target'] = round(float(target), 2)
    except ValueError:
        errors.append("Target sum must be a valid number")
    
    # Buffer
    try:
        buffer_val = round(float(buffer), 2)
        if buffer_val < 0:
            errors.append("Buffer cannot be negative")
        else:
            params['buffer'] = buffer_val
    except ValueError:
        errors.append("Buffer must be a valid number")
    
    # Min size
    try:
        min_val = int(min_size)
        if min_val < 1:
            errors.append("Minimum size must be at least 1")
        else:
            params['min_size'] = min_val
    except ValueError:
        errors.append("Minimum size must be a whole number")
    
    # Max size
    try:
        max_val = int(max_size)
        if max_val < 1:
            errors.append("Maximum size must be at least 1")
        else:
            params['max_size'] = max_val
    except ValueError:
        errors.append("Maximum size must be a whole number")
    
    # Max results
    try:
        results_val = int(max_results)
        if results_val < 1:
            errors.append("Maximum results must be at least 1")
        else:
            params['max_results'] = results_val
    except ValueError:
        errors.append("Maximum results must be a whole number")
    
    # Cross-validation
    if 'min_size' in params and 'max_size' in params:
        if params['min_size'] > params['max_size']:
            errors.append("Minimum size cannot be greater than maximum size")
    
    if errors:
        return None, errors
    
    return params, []


def format_number(value: float) -> str:
    """
    Format a number for display with 2 decimal places.
    
    Args:
        value: Number to format
        
    Returns:
        Formatted string
    """
    return f"{value:,.2f}"


def rgb_to_hex(rgb: Tuple[int, int, int]) -> str:
    """
    Convert RGB tuple to hex color string.
    
    Args:
        rgb: Tuple of (r, g, b) values 0-255
        
    Returns:
        Hex string like '#AABBCC'
    """
    return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"


def hex_to_rgb(hex_color: str) -> Tuple[int, int, int]:
    """
    Convert hex color string to RGB tuple.
    
    Args:
        hex_color: Hex string like '#AABBCC' or 'AABBCC'
        
    Returns:
        Tuple of (r, g, b) values 0-255
    """
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def get_contrast_color(rgb: Tuple[int, int, int]) -> Tuple[int, int, int]:
    """
    Get a contrasting text color (black or white) for a background.
    
    Args:
        rgb: Background color RGB tuple
        
    Returns:
        Black (0,0,0) or white (255,255,255) for best contrast
    """
    # Calculate luminance
    r, g, b = rgb
    luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    
    return (0, 0, 0) if luminance > 0.5 else (255, 255, 255)
