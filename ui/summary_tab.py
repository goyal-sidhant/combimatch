"""
summary_tab.py - Highlighted Summary Tab for CombiMatch

Displays all finalized combinations with their:
- Assigned colors
- Numbers involved
- Sums and differences
"""

from typing import List
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
    QListWidget, QListWidgetItem, QGroupBox,
    QFrame, QScrollArea
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QFont

from models import FinalizedCombination
from utils import rgb_to_hex, get_contrast_color


class SummaryCard(QFrame):
    """
    A card widget displaying a single finalized combination.
    """
    
    def __init__(self, finalized: FinalizedCombination, parent=None):
        super().__init__(parent)
        self.finalized = finalized
        self._init_ui()
    
    def _init_ui(self):
        """Initialize the card UI."""
        self.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        self.setLineWidth(1)
        
        # Set background color
        hex_color = rgb_to_hex(self.finalized.color)
        contrast = get_contrast_color(self.finalized.color)
        contrast_hex = rgb_to_hex(contrast)
        
        self.setStyleSheet(f"""
            SummaryCard {{
                background-color: {hex_color};
                border-radius: 6px;
                padding: 8px;
            }}
            QLabel {{
                color: {contrast_hex};
            }}
        """)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        
        # Header with color name and combo number
        header = QLabel(f"#{self.finalized.finalized_at + 1} - {self.finalized.color_name}")
        header_font = QFont()
        header_font.setBold(True)
        header_font.setPointSize(10)
        header.setFont(header_font)
        layout.addWidget(header)
        
        # Combination info
        combo = self.finalized.combination
        
        # Sum and difference
        sum_text = f"Sum: {combo.sum:.2f} ({combo.difference_display})"
        sum_label = QLabel(sum_text)
        layout.addWidget(sum_label)
        
        # Number of items
        count_label = QLabel(f"Items: {combo.size}")
        layout.addWidget(count_label)
        
        # Values (in original order)
        items_in_order = combo.items_in_original_order()
        values_text = ", ".join(
            f"{item.value:.2f}" for item in items_in_order
        )
        values_label = QLabel(f"Values: {values_text}")
        values_label.setWordWrap(True)
        layout.addWidget(values_label)
        
        # Row references if from Excel
        if any(item.row for item in items_in_order):
            rows_text = ", ".join(
                f"R{item.row}" for item in items_in_order if item.row
            )
            rows_label = QLabel(f"Rows: {rows_text}")
            rows_label.setWordWrap(True)
            layout.addWidget(rows_label)


class SummaryTab(QWidget):
    """
    Summary tab showing all finalized combinations.
    """
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.finalized_list: List[FinalizedCombination] = []
        self._init_ui()
    
    def _init_ui(self):
        """Initialize the user interface."""
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        
        # Header
        header_layout = QHBoxLayout()
        
        header = QLabel("Finalized Combinations")
        header.setProperty("heading", True)
        header_layout.addWidget(header)
        
        header_layout.addStretch()
        
        self.count_label = QLabel("0 combinations")
        self.count_label.setProperty("subheading", True)
        header_layout.addWidget(self.count_label)
        
        layout.addLayout(header_layout)
        
        # Stats row
        self.stats_label = QLabel("Total: $0.00")
        self.stats_label.setProperty("subheading", True)
        layout.addWidget(self.stats_label)
        
        # Scroll area for cards
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        
        # Container for cards
        self.cards_container = QWidget()
        self.cards_layout = QVBoxLayout(self.cards_container)
        self.cards_layout.setSpacing(12)
        self.cards_layout.setAlignment(Qt.AlignTop)
        
        scroll.setWidget(self.cards_container)
        layout.addWidget(scroll)
        
        # Empty state
        self.empty_label = QLabel(
            "No combinations finalized yet.\n"
            "Use the Find tab to find and finalize combinations."
        )
        self.empty_label.setAlignment(Qt.AlignCenter)
        self.empty_label.setProperty("subheading", True)
        self.cards_layout.addWidget(self.empty_label)
    
    def add_finalized(self, finalized: FinalizedCombination):
        """
        Add a newly finalized combination.
        
        Args:
            finalized: The finalized combination to add
        """
        # Hide empty state
        if self.empty_label.isVisible():
            self.empty_label.setVisible(False)
        
        self.finalized_list.append(finalized)
        
        # Create and add card
        card = SummaryCard(finalized)
        self.cards_layout.addWidget(card)
        
        # Update stats
        self._update_stats()
    
    def _update_stats(self):
        """Update the statistics display."""
        count = len(self.finalized_list)
        self.count_label.setText(f"{count} combination{'s' if count != 1 else ''}")
        
        # Calculate total sum of all finalized
        total = sum(f.combination.sum for f in self.finalized_list)
        self.stats_label.setText(f"Total Sum: {total:,.2f}")
    
    def clear_all(self):
        """Clear all finalized combinations."""
        # Remove all cards
        while self.cards_layout.count() > 1:
            item = self.cards_layout.takeAt(1)
            if item.widget():
                item.widget().deleteLater()
        
        self.finalized_list.clear()
        self.empty_label.setVisible(True)
        self._update_stats()
    
    def get_all_finalized(self) -> List[FinalizedCombination]:
        """Get all finalized combinations."""
        return self.finalized_list.copy()
