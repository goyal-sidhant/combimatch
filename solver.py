"""
solver.py - Subset Sum Solver for CombiMatch

Implements the combination finding algorithm with:
- Buffer tolerance support
- Min/max combination size constraints
- Progressive result yielding for UI responsiveness
- Sorted output by combination size
"""

from typing import List, Generator, Optional, Callable
from itertools import combinations
from models import NumberItem, Combination, CombinationResult


class SubsetSumSolver:
    """
    Finds combinations of numbers that sum to a target within tolerance.
    
    Uses itertools.combinations for generating subsets, iterating from
    smallest to largest combination sizes for optimal result ordering.
    """
    
    def __init__(
        self,
        items: List[NumberItem],
        target: float,
        buffer: float = 0.0,
        min_size: int = 1,
        max_size: int = 10,
        max_results: int = 100,
        progress_callback: Optional[Callable[[int, int], None]] = None
    ):
        """
        Initialize the solver.
        
        Args:
            items: List of NumberItem objects to search through
            target: Target sum to find
            buffer: Tolerance (+/-) for accepting sums
            min_size: Minimum numbers in a combination
            max_size: Maximum numbers in a combination
            max_results: Maximum combinations to return
            progress_callback: Optional callback(current_size, total_checked)
        """
        # Filter out finalized items
        self.items = [item for item in items if not item.is_finalized]
        self.target = round(target, 2)
        self.buffer = round(abs(buffer), 2)
        self.min_size = max(1, min_size)
        self.max_size = min(max_size, len(self.items))
        self.max_results = max_results
        self.progress_callback = progress_callback
        
        # Calculate bounds
        self.lower_bound = round(self.target - self.buffer, 2)
        self.upper_bound = round(self.target + self.buffer, 2)
        
        # Results container
        self.result = CombinationResult(
            target=self.target,
            buffer=self.buffer,
            min_size=self.min_size,
            max_size=self.max_size,
            max_results=self.max_results
        )
        
        # Control flag for stopping
        self._stop_requested = False
    
    def stop(self):
        """Request the solver to stop."""
        self._stop_requested = True
    
    def _is_valid_sum(self, total: float) -> bool:
        """Check if a sum falls within the acceptable range."""
        total = round(total, 2)
        return self.lower_bound <= total <= self.upper_bound
    
    def find_combinations(self) -> CombinationResult:
        """
        Find all valid combinations.
        
        Returns:
            CombinationResult containing all found combinations,
            sorted by size then by closeness to target.
        """
        total_checked = 0
        
        # Iterate from smallest to largest combination size
        for size in range(self.min_size, self.max_size + 1):
            if self._stop_requested:
                break
                
            # Generate all combinations of current size
            for combo_items in combinations(self.items, size):
                if self._stop_requested:
                    break
                    
                total_checked += 1
                
                # Calculate sum
                combo_sum = sum(item.value for item in combo_items)
                
                # Check if valid
                if self._is_valid_sum(combo_sum):
                    combo = Combination(
                        items=list(combo_items),
                        target=self.target
                    )
                    
                    # Try to add to results
                    if not self.result.add_combination(combo):
                        # Max results reached, but continue counting
                        pass
                
                # Progress callback every 1000 iterations
                if self.progress_callback and total_checked % 1000 == 0:
                    self.progress_callback(size, total_checked)
            
            # Check if we have enough results
            if len(self.result.combinations) >= self.max_results:
                break
        
        # Sort results
        self.result.sort_combinations()
        
        return self.result
    
    def find_combinations_generator(self) -> Generator[Combination, None, None]:
        """
        Generator version that yields combinations as they are found.
        
        Useful for real-time UI updates.
        
        Yields:
            Combination objects as they are discovered
        """
        total_checked = 0
        found_count = 0
        
        for size in range(self.min_size, self.max_size + 1):
            if self._stop_requested or found_count >= self.max_results:
                break
                
            for combo_items in combinations(self.items, size):
                if self._stop_requested or found_count >= self.max_results:
                    break
                    
                total_checked += 1
                combo_sum = sum(item.value for item in combo_items)
                
                if self._is_valid_sum(combo_sum):
                    combo = Combination(
                        items=list(combo_items),
                        target=self.target
                    )
                    found_count += 1
                    yield combo
                
                if self.progress_callback and total_checked % 1000 == 0:
                    self.progress_callback(size, total_checked)


def estimate_combinations(n: int, min_k: int, max_k: int) -> int:
    """
    Estimate total number of combinations to check.
    
    Args:
        n: Total number of items
        min_k: Minimum combination size
        max_k: Maximum combination size
    
    Returns:
        Estimated total combinations
    """
    from math import comb
    total = 0
    for k in range(min_k, min(max_k, n) + 1):
        total += comb(n, k)
    return total


def quick_check_possible(
    items: List[NumberItem],
    target: float,
    buffer: float
) -> bool:
    """
    Quick check if any valid combination is possible.
    
    Checks basic bounds to avoid unnecessary computation.
    
    Args:
        items: Available items
        target: Target sum
        buffer: Tolerance
    
    Returns:
        True if a valid combination might exist
    """
    if not items:
        return False
    
    available = [item for item in items if not item.is_finalized]
    if not available:
        return False
    
    values = [item.value for item in available]
    min_val = min(values)
    max_val = max(values)
    total_sum = sum(values)
    
    lower = target - buffer
    upper = target + buffer
    
    # If smallest value is larger than upper bound, impossible
    if min_val > upper:
        return False
    
    # If sum of all is less than lower bound, impossible
    if total_sum < lower:
        return False
    
    return True
