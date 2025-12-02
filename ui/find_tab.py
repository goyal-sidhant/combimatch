"""
find_tab.py - Find Combinations Tab for CombiMatch

The main workspace containing:
- Input area (manual or Excel)
- Parameter fields
- Results list (split into Exact and Approximate matches)
- Source list with highlighting
- Selected combination info panel
"""

from typing import List, Optional, Set
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QTextEdit, QPushButton, QListWidget,
    QListWidgetItem, QGroupBox, QComboBox, QSplitter,
    QMessageBox, QFrame, QSpinBox, QSizePolicy
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QColor, QFont, QBrush

from models import NumberItem, Combination, FinalizedCombination, InputMode
from solver import SubsetSumSolver, quick_check_possible
from excel_handler import get_excel_handler
from utils import (
    parse_numbers_line_separated, 
    parse_numbers_comma_separated,
    validate_parameters,
    ColorManager,
    rgb_to_hex,
    get_contrast_color
)


# Colors for UI
SELECTION_HIGHLIGHT_COLOR = (255, 165, 0)  # Orange - prominent selection highlight
FINALIZED_TEXT_COLOR = (150, 150, 150)  # Grey for finalized items
DISABLED_COMBO_COLOR = (240, 240, 240)  # Light grey for invalid combos


class SolverThread(QThread):
    """Background thread for running the solver."""
    
    result_found = pyqtSignal(object)  # Emits Combination
    finished_signal = pyqtSignal(int, int)  # total_found, total_checked
    progress = pyqtSignal(int, int)  # current_size, total_checked
    
    def __init__(self, solver: SubsetSumSolver):
        super().__init__()
        self.solver = solver
    
    def run(self):
        """Run the solver and emit results."""
        total_found = 0
        total_checked = 0
        
        for combo in self.solver.find_combinations_generator():
            total_found += 1
            self.result_found.emit(combo)
        
        self.finished_signal.emit(total_found, total_checked)
    
    def stop(self):
        """Request solver to stop."""
        self.solver.stop()


class SelectedComboInfoPanel(QFrame):
    """Panel showing details of the currently selected combination."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        self.setStyleSheet("""
            SelectedComboInfoPanel {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 6px;
                padding: 8px;
            }
        """)
        self._init_ui()
    
    def _init_ui(self):
        """Initialize the panel UI."""
        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        layout.setContentsMargins(12, 12, 12, 12)
        
        # Title
        title = QLabel("Selected Combination")
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(10)
        title.setFont(title_font)
        layout.addWidget(title)
        
        # Info grid
        info_layout = QGridLayout()
        info_layout.setSpacing(4)
        
        info_layout.addWidget(QLabel("Sum:"), 0, 0)
        self.sum_label = QLabel("-")
        self.sum_label.setStyleSheet("font-weight: bold;")
        info_layout.addWidget(self.sum_label, 0, 1)
        
        info_layout.addWidget(QLabel("Difference:"), 1, 0)
        self.diff_label = QLabel("-")
        info_layout.addWidget(self.diff_label, 1, 1)
        
        info_layout.addWidget(QLabel("Items:"), 2, 0)
        self.count_label = QLabel("-")
        info_layout.addWidget(self.count_label, 2, 1)
        
        layout.addLayout(info_layout)
        
        # Values list
        layout.addWidget(QLabel("Values:"))
        self.values_label = QLabel("-")
        self.values_label.setWordWrap(True)
        self.values_label.setStyleSheet("color: #495057; padding-left: 8px;")
        layout.addWidget(self.values_label)
        
        # Placeholder for empty state
        self.placeholder = QLabel("Select a combination to see details")
        self.placeholder.setStyleSheet("color: #868e96; font-style: italic;")
        self.placeholder.setAlignment(Qt.AlignCenter)
    
    def update_info(self, combo: Optional[Combination]):
        """Update the panel with combination info."""
        if combo is None:
            self.sum_label.setText("-")
            self.diff_label.setText("-")
            self.count_label.setText("-")
            self.values_label.setText("-")
        else:
            self.sum_label.setText(f"{combo.sum:,.2f}")
            
            diff = combo.difference
            diff_color = "#51cf66" if diff == 0 else "#fcc419"
            self.diff_label.setText(f"{combo.difference_display}")
            self.diff_label.setStyleSheet(f"color: {diff_color}; font-weight: bold;")
            
            self.count_label.setText(f"{combo.size} numbers")
            
            # Show values in original order
            items_ordered = combo.items_in_original_order()
            values_text = "\n".join(f"• {item.value:,.2f}" for item in items_ordered)
            self.values_label.setText(values_text)


class FindTab(QWidget):
    """
    Main workspace tab for finding combinations.
    """
    
    # Signal emitted when a combination is finalized
    combination_finalized = pyqtSignal(object)  # FinalizedCombination
    
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.items: List[NumberItem] = []
        self.exact_combinations: List[Combination] = []
        self.approx_combinations: List[Combination] = []
        self.finalized: List[FinalizedCombination] = []
        self.color_manager = ColorManager()
        self.selected_combo: Optional[Combination] = None
        self.solver_thread: Optional[SolverThread] = None
        self.input_mode = InputMode.LINE_SEPARATED
        self.current_target: float = 0.0
        
        self._init_ui()
        self._connect_signals()
    
    def _init_ui(self):
        """Initialize the user interface."""
        layout = QHBoxLayout(self)
        layout.setSpacing(16)
        
        # Create splitter for resizable panels
        splitter = QSplitter(Qt.Horizontal)
        
        # Left panel - Input and Parameters
        left_panel = self._create_left_panel()
        splitter.addWidget(left_panel)
        
        # Middle panel - Results (Exact + Approx)
        middle_panel = self._create_middle_panel()
        splitter.addWidget(middle_panel)
        
        # Right panel - Source List + Info Panel
        right_panel = self._create_right_panel()
        splitter.addWidget(right_panel)
        
        # Set initial sizes
        splitter.setSizes([280, 380, 320])
        
        layout.addWidget(splitter)
    
    def _create_left_panel(self) -> QWidget:
        """Create the input and parameters panel."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(12)
        
        # Input Mode Selection
        mode_group = QGroupBox("Input Mode")
        mode_layout = QHBoxLayout(mode_group)
        
        self.mode_combo = QComboBox()
        self.mode_combo.addItems([
            "Line Separated",
            "Comma Separated",
            "Excel Selection"
        ])
        mode_layout.addWidget(self.mode_combo)
        
        layout.addWidget(mode_group)
        
        # Input Area
        input_group = QGroupBox("Numbers Input")
        input_layout = QVBoxLayout(input_group)
        
        self.input_text = QTextEdit()
        self.input_text.setPlaceholderText(
            "Enter numbers here...\n"
            "One per line (Line mode)\n"
            "or separated by commas (Comma mode)"
        )
        input_layout.addWidget(self.input_text)
        
        # Excel button (hidden initially)
        self.excel_btn = QPushButton("📊 Grab from Excel")
        self.excel_btn.setVisible(False)
        input_layout.addWidget(self.excel_btn)
        
        self.load_btn = QPushButton("Load Numbers")
        input_layout.addWidget(self.load_btn)
        
        layout.addWidget(input_group)
        
        # Parameters
        params_group = QGroupBox("Search Parameters")
        params_layout = QGridLayout(params_group)
        
        # Target Sum
        params_layout.addWidget(QLabel("Target Sum:"), 0, 0)
        self.target_input = QLineEdit()
        self.target_input.setPlaceholderText("e.g., 1000.00")
        params_layout.addWidget(self.target_input, 0, 1)
        
        # Buffer
        params_layout.addWidget(QLabel("Buffer (±):"), 1, 0)
        self.buffer_input = QLineEdit()
        self.buffer_input.setPlaceholderText("e.g., 0.50")
        self.buffer_input.setText("0")
        params_layout.addWidget(self.buffer_input, 1, 1)
        
        # Min Size
        params_layout.addWidget(QLabel("Min Numbers:"), 2, 0)
        self.min_size_input = QSpinBox()
        self.min_size_input.setRange(1, 100)
        self.min_size_input.setValue(1)
        params_layout.addWidget(self.min_size_input, 2, 1)
        
        # Max Size
        params_layout.addWidget(QLabel("Max Numbers:"), 3, 0)
        self.max_size_input = QSpinBox()
        self.max_size_input.setRange(1, 100)
        self.max_size_input.setValue(10)
        params_layout.addWidget(self.max_size_input, 3, 1)
        
        # Max Results
        params_layout.addWidget(QLabel("Max Results:"), 4, 0)
        self.max_results_input = QSpinBox()
        self.max_results_input.setRange(1, 10000)
        self.max_results_input.setValue(100)
        params_layout.addWidget(self.max_results_input, 4, 1)
        
        layout.addWidget(params_group)
        
        # Find Button
        self.find_btn = QPushButton("🔍 Find Combinations")
        self.find_btn.setMinimumHeight(40)
        layout.addWidget(self.find_btn)
        
        # Status
        self.status_label = QLabel("")
        self.status_label.setProperty("subheading", True)
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
        
        layout.addStretch()
        
        return panel
    
    def _create_middle_panel(self) -> QWidget:
        """Create the results panel with Exact and Approximate sections."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(12)
        
        # Exact Matches Section
        exact_group = QGroupBox("Exact Matches")
        exact_layout = QVBoxLayout(exact_group)
        
        self.exact_count_label = QLabel("0 exact matches")
        self.exact_count_label.setProperty("subheading", True)
        exact_layout.addWidget(self.exact_count_label)
        
        self.exact_list = QListWidget()
        self.exact_list.setSelectionMode(QListWidget.SingleSelection)
        exact_layout.addWidget(self.exact_list)
        
        layout.addWidget(exact_group)
        
        # Approximate Matches Section
        approx_group = QGroupBox("Approximate Matches")
        approx_layout = QVBoxLayout(approx_group)
        
        self.approx_count_label = QLabel("0 approximate matches")
        self.approx_count_label.setProperty("subheading", True)
        approx_layout.addWidget(self.approx_count_label)
        
        self.approx_list = QListWidget()
        self.approx_list.setSelectionMode(QListWidget.SingleSelection)
        approx_layout.addWidget(self.approx_list)
        
        layout.addWidget(approx_group)
        
        # Finalize button
        self.finalize_btn = QPushButton("✓ Finalize Selected")
        self.finalize_btn.setEnabled(False)
        self.finalize_btn.setMinimumHeight(36)
        layout.addWidget(self.finalize_btn)
        
        return panel
    
    def _create_right_panel(self) -> QWidget:
        """Create the source list and info panel."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(12)
        
        # Selected Combo Info Panel
        self.info_panel = SelectedComboInfoPanel()
        layout.addWidget(self.info_panel)
        
        # Source Numbers Section
        source_group = QGroupBox("Source Numbers")
        source_layout = QVBoxLayout(source_group)
        
        self.source_count_label = QLabel("0 numbers loaded")
        self.source_count_label.setProperty("subheading", True)
        source_layout.addWidget(self.source_count_label)
        
        self.source_list = QListWidget()
        self.source_list.setSelectionMode(QListWidget.NoSelection)
        source_layout.addWidget(self.source_list)
        
        layout.addWidget(source_group)
        
        return panel
    
    def _connect_signals(self):
        """Connect UI signals to slots."""
        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        self.load_btn.clicked.connect(self._on_load_numbers)
        self.excel_btn.clicked.connect(self._on_grab_from_excel)
        self.find_btn.clicked.connect(self._on_find_combinations)
        
        # Connect both lists for selection
        self.exact_list.currentRowChanged.connect(self._on_exact_selected)
        self.approx_list.currentRowChanged.connect(self._on_approx_selected)
        
        self.finalize_btn.clicked.connect(self._on_finalize)
    
    def _on_mode_changed(self, index: int):
        """Handle input mode change."""
        modes = [InputMode.LINE_SEPARATED, InputMode.COMMA_SEPARATED, InputMode.EXCEL]
        self.input_mode = modes[index]
        
        is_excel = self.input_mode == InputMode.EXCEL
        self.input_text.setVisible(not is_excel)
        self.excel_btn.setVisible(is_excel)
        self.load_btn.setVisible(not is_excel)
        
        if is_excel:
            self.input_text.setPlaceholderText("")
        else:
            placeholder = (
                "Enter numbers, one per line" if index == 0
                else "Enter numbers, separated by commas"
            )
            self.input_text.setPlaceholderText(placeholder)
    
    def _on_load_numbers(self):
        """Load numbers from text input."""
        text = self.input_text.toPlainText().strip()
        
        if not text:
            self._show_status("Please enter some numbers", error=True)
            return
        
        # Parse based on mode
        if self.input_mode == InputMode.LINE_SEPARATED:
            items, errors = parse_numbers_line_separated(text)
        else:
            items, errors = parse_numbers_comma_separated(text)
        
        if errors:
            self._show_status(f"Errors: {'; '.join(errors[:3])}", error=True)
            if not items:
                return
        
        self._set_items(items)
        self._show_status(f"Loaded {len(items)} numbers")
    
    def _on_grab_from_excel(self):
        """Grab numbers from Excel selection."""
        handler = get_excel_handler()
        
        if not handler.is_connected():
            if not handler.connect():
                QMessageBox.warning(
                    self, "Excel Not Found",
                    "Could not connect to Excel.\n"
                    "Make sure Excel is open with a workbook."
                )
                return
        
        # Get active workbook info
        wb_info = handler.get_active_workbook_info()
        if not wb_info:
            QMessageBox.warning(
                self, "No Workbook",
                "No active workbook found in Excel."
            )
            return
        
        # Select the active workbook
        handler.select_workbook(wb_info['name'])
        
        # Auto-save first
        if not handler.auto_save():
            reply = QMessageBox.question(
                self, "Save Failed",
                "Could not auto-save the workbook.\n"
                "Continue anyway?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.No:
                return
        
        # Read selection
        items, warnings = handler.read_selection()
        
        if warnings:
            for warning in warnings:
                QMessageBox.information(self, "Note", warning)
        
        if not items:
            self._show_status("No valid numbers in selection", error=True)
            return
        
        self._set_items(items)
        self._show_status(
            f"Loaded {len(items)} numbers from {wb_info['name']}"
        )
    
    def _set_items(self, items: List[NumberItem]):
        """Set the source items and update UI."""
        self.items = items
        self.exact_combinations.clear()
        self.approx_combinations.clear()
        self.selected_combo = None
        
        # Update source list
        self.source_list.clear()
        for item in items:
            list_item = QListWidgetItem(item.display_label)
            list_item.setData(Qt.UserRole, item)
            self.source_list.addItem(list_item)
        
        self.source_count_label.setText(f"{len(items)} numbers loaded")
        
        # Clear results
        self.exact_list.clear()
        self.approx_list.clear()
        self.exact_count_label.setText("0 exact matches")
        self.approx_count_label.setText("0 approximate matches")
        self.finalize_btn.setEnabled(False)
        self.info_panel.update_info(None)
    
    def _on_find_combinations(self):
        """Start searching for combinations."""
        if not self.items:
            self._show_status("Please load numbers first", error=True)
            return
        
        # Get available items (not finalized)
        available = [item for item in self.items if not item.is_finalized]
        
        if not available:
            self._show_status("All numbers are finalized", error=True)
            return
        
        # Validate parameters
        params, errors = validate_parameters(
            self.target_input.text(),
            self.buffer_input.text(),
            str(self.min_size_input.value()),
            str(self.max_size_input.value()),
            str(self.max_results_input.value())
        )
        
        if errors:
            self._show_status(f"Error: {errors[0]}", error=True)
            return
        
        self.current_target = params['target']
        
        # Quick check
        if not quick_check_possible(available, params['target'], params['buffer']):
            self._show_status("No valid combinations possible", error=True)
            return
        
        # Clear previous results
        self.exact_combinations.clear()
        self.approx_combinations.clear()
        self.exact_list.clear()
        self.approx_list.clear()
        self.selected_combo = None
        self.finalize_btn.setEnabled(False)
        self.info_panel.update_info(None)
        self._clear_highlights()
        
        # Create solver
        solver = SubsetSumSolver(
            items=self.items,
            target=params['target'],
            buffer=params['buffer'],
            min_size=params['min_size'],
            max_size=params['max_size'],
            max_results=params['max_results']
        )
        
        # Run in thread
        self.find_btn.setEnabled(False)
        self.find_btn.setText("Searching...")
        
        self.solver_thread = SolverThread(solver)
        self.solver_thread.result_found.connect(self._on_combo_found)
        self.solver_thread.finished_signal.connect(self._on_search_finished)
        self.solver_thread.start()
    
    def _on_combo_found(self, combo: Combination):
        """Handle a new combination found."""
        is_exact = abs(combo.difference) < 0.001  # Effectively zero
        
        # Create list item text
        item_text = (
            f"[{combo.size}] "
            f"{', '.join(str(i.value) for i in combo.items)} "
            f"= {combo.sum:.2f}"
        )
        
        if not is_exact:
            item_text += f" ({combo.difference_display})"
        
        list_item = QListWidgetItem(item_text)
        list_item.setData(Qt.UserRole, combo)
        
        if is_exact:
            self.exact_combinations.append(combo)
            self.exact_list.addItem(list_item)
            self.exact_count_label.setText(f"{len(self.exact_combinations)} exact matches")
        else:
            self.approx_combinations.append(combo)
            self.approx_list.addItem(list_item)
            self.approx_count_label.setText(f"{len(self.approx_combinations)} approximate matches")
    
    def _on_search_finished(self, total_found: int, total_checked: int):
        """Handle search completion."""
        self.find_btn.setEnabled(True)
        self.find_btn.setText("🔍 Find Combinations")
        
        total = len(self.exact_combinations) + len(self.approx_combinations)
        if total > 0:
            self._show_status(
                f"Found {len(self.exact_combinations)} exact, "
                f"{len(self.approx_combinations)} approximate"
            )
        else:
            self._show_status("No combinations found", error=True)
    
    def _on_exact_selected(self, row: int):
        """Handle exact match selection."""
        # Clear approx selection
        self.approx_list.blockSignals(True)
        self.approx_list.clearSelection()
        self.approx_list.setCurrentRow(-1)
        self.approx_list.blockSignals(False)
        
        self._handle_selection(self.exact_combinations, row)
    
    def _on_approx_selected(self, row: int):
        """Handle approximate match selection."""
        # Clear exact selection
        self.exact_list.blockSignals(True)
        self.exact_list.clearSelection()
        self.exact_list.setCurrentRow(-1)
        self.exact_list.blockSignals(False)
        
        self._handle_selection(self.approx_combinations, row)
    
    def _handle_selection(self, combo_list: List[Combination], row: int):
        """Handle combination selection from either list."""
        if row < 0 or row >= len(combo_list):
            self.selected_combo = None
            self.finalize_btn.setEnabled(False)
            self.info_panel.update_info(None)
            self._clear_highlights()
            return
        
        self.selected_combo = combo_list[row]
        self.finalize_btn.setEnabled(True)
        
        # Update info panel
        self.info_panel.update_info(self.selected_combo)
        
        # Highlight source items
        self._highlight_combination(self.selected_combo)
    
    def _highlight_combination(self, combo: Combination):
        """Highlight the items in a combination in the source list with prominent color."""
        combo_indices = {item.index for item in combo.items}
        
        for i in range(self.source_list.count()):
            list_item = self.source_list.item(i)
            item: NumberItem = list_item.data(Qt.UserRole)
            
            if item.is_finalized:
                # Keep finalized color with grey text
                if item.finalized_color:
                    color = QColor(*item.finalized_color)
                    list_item.setBackground(color)
                    contrast = get_contrast_color(item.finalized_color)
                    list_item.setForeground(QColor(*contrast))
            elif item.index in combo_indices:
                # Prominent orange highlight for selected combo
                list_item.setBackground(QColor(*SELECTION_HIGHLIGHT_COLOR))
                list_item.setForeground(QColor(0, 0, 0))
                
                # Make font bold
                font = list_item.font()
                font.setBold(True)
                list_item.setFont(font)
            else:
                # Clear highlight
                list_item.setBackground(QColor("transparent"))
                list_item.setForeground(QColor("#343a40"))
                
                # Remove bold
                font = list_item.font()
                font.setBold(False)
                list_item.setFont(font)
    
    def _clear_highlights(self):
        """Clear temporary highlights (keep finalized colors)."""
        for i in range(self.source_list.count()):
            list_item = self.source_list.item(i)
            item: NumberItem = list_item.data(Qt.UserRole)
            
            font = list_item.font()
            font.setBold(False)
            list_item.setFont(font)
            
            if item.is_finalized:
                if item.finalized_color:
                    color = QColor(*item.finalized_color)
                    list_item.setBackground(color)
                    contrast = get_contrast_color(item.finalized_color)
                    list_item.setForeground(QColor(*contrast))
                else:
                    # Grey out finalized without specific color
                    list_item.setBackground(QColor(*DISABLED_COMBO_COLOR))
                    list_item.setForeground(QColor(*FINALIZED_TEXT_COLOR))
            else:
                list_item.setBackground(QColor("transparent"))
                list_item.setForeground(QColor("#343a40"))
    
    def _on_finalize(self):
        """Finalize the selected combination."""
        if not self.selected_combo:
            return
        
        # Get next color
        color, color_name = self.color_manager.get_next_color()
        
        # Get indices of finalized items
        finalized_indices: Set[int] = set()
        for item in self.selected_combo.items:
            item.is_finalized = True
            item.finalized_color = color
            finalized_indices.add(item.index)
        
        # Create finalized record
        finalized = FinalizedCombination(
            combination=self.selected_combo,
            color=color,
            color_name=color_name,
            finalized_at=len(self.finalized)
        )
        self.finalized.append(finalized)
        
        # Update source list colors (grey out finalized)
        self._update_source_colors()
        
        # Highlight in Excel if connected
        handler = get_excel_handler()
        if handler.is_connected():
            handler.highlight_items(self.selected_combo.items, color)
        
        # Emit signal for summary tab
        self.combination_finalized.emit(finalized)
        
        # Remove/disable combinations that use any finalized numbers
        self._remove_invalid_combinations(finalized_indices)
        
        # Clear selection
        self.selected_combo = None
        self.finalize_btn.setEnabled(False)
        self.info_panel.update_info(None)
        
        self._show_status(f"Finalized with {color_name}")
    
    def _remove_invalid_combinations(self, finalized_indices: Set[int]):
        """Remove or grey out combinations that use any of the finalized numbers."""
        # Process exact matches
        self._filter_combination_list(
            self.exact_list, 
            self.exact_combinations, 
            finalized_indices
        )
        self.exact_count_label.setText(f"{len(self.exact_combinations)} exact matches")
        
        # Process approximate matches
        self._filter_combination_list(
            self.approx_list,
            self.approx_combinations,
            finalized_indices
        )
        self.approx_count_label.setText(f"{len(self.approx_combinations)} approximate matches")
    
    def _filter_combination_list(
        self,
        list_widget: QListWidget,
        combo_list: List[Combination],
        finalized_indices: Set[int]
    ):
        """Remove combinations that contain any finalized indices."""
        # Find indices to remove (iterate in reverse to avoid index shifting)
        indices_to_remove = []
        
        for i, combo in enumerate(combo_list):
            combo_item_indices = {item.index for item in combo.items}
            if combo_item_indices & finalized_indices:  # If there's any overlap
                indices_to_remove.append(i)
        
        # Remove in reverse order
        for i in reversed(indices_to_remove):
            list_widget.takeItem(i)
            combo_list.pop(i)
    
    def _update_source_colors(self):
        """Update all source list item colors - grey out finalized items."""
        for i in range(self.source_list.count()):
            list_item = self.source_list.item(i)
            item: NumberItem = list_item.data(Qt.UserRole)
            
            # Remove bold from all
            font = list_item.font()
            font.setBold(False)
            list_item.setFont(font)
            
            if item.is_finalized and item.finalized_color:
                color = QColor(*item.finalized_color)
                list_item.setBackground(color)
                
                # Set text color for contrast
                contrast = get_contrast_color(item.finalized_color)
                list_item.setForeground(QColor(*contrast))
                
                # Update display text to show finalized status
                original_text = item.display_label
                list_item.setText(f"✓ {original_text}")
            elif item.is_finalized:
                # Grey out without color
                list_item.setBackground(QColor(*DISABLED_COMBO_COLOR))
                list_item.setForeground(QColor(*FINALIZED_TEXT_COLOR))
            else:
                list_item.setBackground(QColor("transparent"))
                list_item.setForeground(QColor("#343a40"))
    
    def _show_status(self, message: str, error: bool = False):
        """Show a status message."""
        self.status_label.setText(message)
        color = "#ff6b6b" if error else "#51cf66"
        self.status_label.setStyleSheet(f"color: {color};")
    
    def get_finalized_combinations(self) -> List[FinalizedCombination]:
        """Get all finalized combinations."""
        return self.finalized.copy()
    
    def check_excel_connection(self) -> bool:
        """Check if Excel is still connected."""
        handler = get_excel_handler()
        return handler.is_connected()
