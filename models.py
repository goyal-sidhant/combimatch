"""
models.py - Data models for CombiMatch

Contains dataclasses and type definitions for:
- NumberItem: Individual number with source reference
- Combination: A valid combination of numbers
- CombinationResult: Search results container
"""

from dataclasses import dataclass, field
from typing import List, Optional, Tuple
from enum import Enum


class InputMode(Enum):
    """Input mode for number entry."""
    LINE_SEPARATED = "line"
    COMMA_SEPARATED = "comma"
    EXCEL = "excel"


class ItemSource(Enum):
    """Source of a number item."""
    MANUAL = "manual"
    EXCEL = "excel"


@dataclass
class NumberItem:
    """
    Represents a single number item from the input list.
    
    Attributes:
        value: The numeric value (rounded to 2 decimals)
        index: Position in the original input (0-based)
        row: Excel row number (if from Excel), None otherwise
        column: Excel column letter (if from Excel), None otherwise
        is_finalized: Whether this item is part of a finalized combination
        finalized_color: RGB tuple for highlight color if finalized
        source: Where this item came from (manual or excel)
    """
    value: float
    index: int
    row: Optional[int] = None
    column: Optional[str] = None
    is_finalized: bool = False
    finalized_color: Optional[Tuple[int, int, int]] = None
    source: ItemSource = ItemSource.MANUAL
    
    def __post_init__(self):
        """Round value to 2 decimal places."""
        self.value = round(self.value, 2)
    
    @property
    def display_label(self) -> str:
        """Get display label for this item."""
        if self.source == ItemSource.EXCEL and self.row is not None:
            col_str = f"{self.column}" if self.column else ""
            return f"Row {self.row}{col_str}: {self.value}"
        return f"#{self.index + 1}: {self.value}"
    
    def __hash__(self):
        """Hash based on index and source for set operations."""
        return hash((self.index, self.row, self.column))
    
    def __eq__(self, other):
        """Equality based on index and source."""
        if not isinstance(other, NumberItem):
            return False
        return (self.index == other.index and 
                self.row == other.row and 
                self.column == other.column)


@dataclass
class Combination:
    """
    Represents a valid combination of numbers.
    
    Attributes:
        items: List of NumberItem objects in this combination
        target: The target sum that was searched for
    """
    items: List[NumberItem]
    target: float
    
    @property
    def sum(self) -> float:
        """Calculate the sum of all items."""
        return round(sum(item.value for item in self.items), 2)
    
    @property
    def difference(self) -> float:
        """Calculate difference from target."""
        return round(self.sum - self.target, 2)
    
    @property
    def size(self) -> int:
        """Number of items in this combination."""
        return len(self.items)
    
    @property
    def difference_display(self) -> str:
        """Formatted difference string with +/- sign."""
        diff = self.difference
        if diff >= 0:
            return f"+{diff:.2f}"
        return f"{diff:.2f}"
    
    def items_in_original_order(self) -> List[NumberItem]:
        """Return items sorted by their original index."""
        return sorted(self.items, key=lambda x: (x.row or 0, x.index))
    
    def __str__(self) -> str:
        values = [str(item.value) for item in self.items]
        return f"[{', '.join(values)}] = {self.sum} ({self.difference_display})"


@dataclass
class CombinationResult:
    """
    Container for search results.
    
    Attributes:
        combinations: List of valid combinations found
        target: The target sum searched for
        buffer: The buffer/tolerance used
        min_size: Minimum combination size
        max_size: Maximum combination size
        max_results: Maximum results requested
        total_found: Total combinations found (may be more than returned)
    """
    combinations: List[Combination] = field(default_factory=list)
    target: float = 0.0
    buffer: float = 0.0
    min_size: int = 1
    max_size: int = 10
    max_results: int = 100
    total_found: int = 0
    
    def add_combination(self, combo: Combination) -> bool:
        """
        Add a combination to results if under max_results limit.
        Returns True if added, False if limit reached.
        """
        self.total_found += 1
        if len(self.combinations) < self.max_results:
            self.combinations.append(combo)
            return True
        return False
    
    def sort_combinations(self):
        """Sort combinations by size (ascending), then by closeness to target."""
        self.combinations.sort(key=lambda c: (c.size, abs(c.difference)))


@dataclass
class FinalizedCombination:
    """
    Represents a finalized (highlighted) combination.
    
    Attributes:
        combination: The original combination
        color: RGB color tuple assigned to this combination
        color_name: Human-readable color name
        finalized_at: Index/order of finalization
    """
    combination: Combination
    color: Tuple[int, int, int]
    color_name: str
    finalized_at: int
    
    @property
    def hex_color(self) -> str:
        """Get hex color string."""
        return f"#{self.color[0]:02x}{self.color[1]:02x}{self.color[2]:02x}"


@dataclass 
class ExcelConnection:
    """
    Represents connection to an Excel workbook.
    
    Attributes:
        workbook_name: Name of the connected workbook
        sheet_name: Name of the active sheet
        file_path: Full path to the file
        is_connected: Whether connection is active
        is_saved: Whether file has been auto-saved
    """
    workbook_name: str = ""
    sheet_name: str = ""
    file_path: str = ""
    is_connected: bool = False
    is_saved: bool = False
