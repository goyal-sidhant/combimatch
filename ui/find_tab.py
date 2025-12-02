"""
find_tab.py - Find Combinations Tab for CombiMatch

The main workspace containing:
- Input area (manual or Excel)
- Parameter fields
- Results list
- Source list with highlighting
"""

from typing import List, Optional
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QTextEdit, QPushButton, QListWidget,
    QListWidgetItem, QGroupBox, QComboBox, QSplitter,
    QMessageBox, QFrame, QSpinBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QColor

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


class FindTab(QWidget):
    """
    Main workspace tab for finding combinations.
    """
    
    # Signal emitted when a combination is finalized
    combination_finalized = pyqtSignal(object)  # FinalizedCombination
    
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.items: List[NumberItem] = []
        self.combinations: List[Combination] = []
        self.finalized: List[FinalizedCombination] = []
        self.color_manager = ColorManager()
        self.selected_combo: Optional[Combination] = None
        self.solver_thread: Optional[SolverThread] = None
        self.input_mode = InputMode.LINE_SEPARATED
        
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
        
        # Middle panel - Results
        middle_panel = self._create_middle_panel()
        splitter.addWidget(middle_panel)
        
        # Right panel - Source List
        right_panel = self._create_right_panel()
        splitter.addWidget(right_panel)
        
        # Set initial sizes (roughly equal)
        splitter.setSizes([300, 350, 300])
        
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
        """Create the results panel."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        
        # Header
        header = QLabel("Combinations Found")
        header.setProperty("heading", True)
        layout.addWidget(header)
        
        # Results count
        self.results_count_label = QLabel("No results yet")
        self.results_count_label.setProperty("subheading", True)
        layout.addWidget(self.results_count_label)
        
        # Results list
        self.results_list = QListWidget()
        self.results_list.setSelectionMode(QListWidget.SingleSelection)
        layout.addWidget(self.results_list)
        
        # Finalize button
        self.finalize_btn = QPushButton("✓ Finalize Selected")
        self.finalize_btn.setEnabled(False)
        layout.addWidget(self.finalize_btn)
        
        return panel
    
    def _create_right_panel(self) -> QWidget:
        """Create the source list panel."""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        
        # Header
        header = QLabel("Source Numbers")
        header.setProperty("heading", True)
        layout.addWidget(header)
        
        # Count
        self.source_count_label = QLabel("0 numbers loaded")
        self.source_count_label.setProperty("subheading", True)
        layout.addWidget(self.source_count_label)
        
        # Source list
        self.source_list = QListWidget()
        self.source_list.setSelectionMode(QListWidget.NoSelection)
        layout.addWidget(self.source_list)
        
        return panel
    
    def _connect_signals(self):
        """Connect UI signals to slots."""
        self.mode_combo.currentIndexChanged.connect(self._on_mode_changed)
        self.load_btn.clicked.connect(self._on_load_numbers)
        self.excel_btn.clicked.connect(self._on_grab_from_excel)
        self.find_btn.clicked.connect(self._on_find_combinations)
        self.results_list.currentRowChanged.connect(self._on_result_selected)
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
        self.combinations.clear()
        self.selected_combo = None
        
        # Update source list
        self.source_list.clear()
        for item in items:
            list_item = QListWidgetItem(item.display_label)
            list_item.setData(Qt.UserRole, item)
            self.source_list.addItem(list_item)
        
        self.source_count_label.setText(f"{len(items)} numbers loaded")
        
        # Clear results
        self.results_list.clear()
        self.results_count_label.setText("No results yet")
        self.finalize_btn.setEnabled(False)
    
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
        
        # Quick check
        if not quick_check_possible(available, params['target'], params['buffer']):
            self._show_status("No valid combinations possible", error=True)
            return
        
        # Clear previous results
        self.combinations.clear()
        self.results_list.clear()
        self.selected_combo = None
        self.finalize_btn.setEnabled(False)
        
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
        self.combinations.append(combo)
        
        # Add to list
        item_text = (
            f"[{combo.size}] "
            f"{', '.join(str(i.value) for i in combo.items)} "
            f"= {combo.sum:.2f} ({combo.difference_display})"
        )
        list_item = QListWidgetItem(item_text)
        list_item.setData(Qt.UserRole, combo)
        self.results_list.addItem(list_item)
        
        self.results_count_label.setText(f"{len(self.combinations)} combinations found")
    
    def _on_search_finished(self, total_found: int, total_checked: int):
        """Handle search completion."""
        self.find_btn.setEnabled(True)
        self.find_btn.setText("🔍 Find Combinations")
        
        if self.combinations:
            self._show_status(f"Found {len(self.combinations)} combinations")
        else:
            self._show_status("No combinations found", error=True)
    
    def _on_result_selected(self, row: int):
        """Handle result selection."""
        if row < 0 or row >= len(self.combinations):
            self.selected_combo = None
            self.finalize_btn.setEnabled(False)
            self._clear_highlights()
            return
        
        self.selected_combo = self.combinations[row]
        self.finalize_btn.setEnabled(True)
        
        # Highlight source items
        self._highlight_combination(self.selected_combo)
    
    def _highlight_combination(self, combo: Combination):
        """Highlight the items in a combination in the source list."""
        combo_indices = {item.index for item in combo.items}
        
        for i in range(self.source_list.count()):
            list_item = self.source_list.item(i)
            item: NumberItem = list_item.data(Qt.UserRole)
            
            if item.is_finalized:
                # Keep finalized color
                color = item.finalized_color
                list_item.setBackground(QColor(*color))
            elif item.index in combo_indices:
                # Highlight as selected
                list_item.setBackground(QColor("#fff3bf"))
            else:
                # Clear highlight
                list_item.setBackground(QColor("transparent"))
    
    def _clear_highlights(self):
        """Clear temporary highlights (keep finalized colors)."""
        for i in range(self.source_list.count()):
            list_item = self.source_list.item(i)
            item: NumberItem = list_item.data(Qt.UserRole)
            
            if item.is_finalized:
                color = item.finalized_color
                list_item.setBackground(QColor(*color))
            else:
                list_item.setBackground(QColor("transparent"))
    
    def _on_finalize(self):
        """Finalize the selected combination."""
        if not self.selected_combo:
            return
        
        # Get next color
        color, color_name = self.color_manager.get_next_color()
        
        # Mark items as finalized
        for item in self.selected_combo.items:
            item.is_finalized = True
            item.finalized_color = color
        
        # Create finalized record
        finalized = FinalizedCombination(
            combination=self.selected_combo,
            color=color,
            color_name=color_name,
            finalized_at=len(self.finalized)
        )
        self.finalized.append(finalized)
        
        # Update source list colors
        self._update_source_colors()
        
        # Highlight in Excel if connected
        handler = get_excel_handler()
        if handler.is_connected():
            handler.highlight_items(self.selected_combo.items, color)
        
        # Emit signal for summary tab
        self.combination_finalized.emit(finalized)
        
        # Remove finalized combo from results
        current_row = self.results_list.currentRow()
        self.results_list.takeItem(current_row)
        self.combinations.pop(current_row)
        
        # Clear selection
        self.selected_combo = None
        self.finalize_btn.setEnabled(False)
        
        self._show_status(f"Finalized with {color_name}")
        self.results_count_label.setText(f"{len(self.combinations)} combinations remaining")
    
    def _update_source_colors(self):
        """Update all source list item colors."""
        for i in range(self.source_list.count()):
            list_item = self.source_list.item(i)
            item: NumberItem = list_item.data(Qt.UserRole)
            
            if item.is_finalized and item.finalized_color:
                color = QColor(*item.finalized_color)
                list_item.setBackground(color)
                
                # Set text color for contrast
                contrast = get_contrast_color(item.finalized_color)
                list_item.setForeground(QColor(*contrast))
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
