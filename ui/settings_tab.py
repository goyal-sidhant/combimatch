"""
settings_tab.py - Settings Tab for CombiMatch

Contains:
- Excel connection status and controls
- Color palette preview
- Application information
"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
    QPushButton, QGroupBox, QComboBox, QFrame,
    QGridLayout, QMessageBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

from excel_handler import get_excel_handler, ExcelHandler
from utils import HIGHLIGHT_COLORS, rgb_to_hex


class ColorSwatch(QFrame):
    """A small color preview square."""
    
    def __init__(self, color: tuple, name: str, parent=None):
        super().__init__(parent)
        self.setFixedSize(30, 30)
        self.setToolTip(name)
        hex_color = rgb_to_hex(color)
        self.setStyleSheet(f"""
            ColorSwatch {{
                background-color: {hex_color};
                border: 1px solid #dee2e6;
                border-radius: 4px;
            }}
        """)


class SettingsTab(QWidget):
    """
    Settings and configuration tab.
    """
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._init_ui()
        self._refresh_excel_status()
    
    def _init_ui(self):
        """Initialize the user interface."""
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        
        # Excel Connection Section
        excel_group = QGroupBox("Excel Connection")
        excel_layout = QVBoxLayout(excel_group)
        
        # Status row
        status_row = QHBoxLayout()
        
        status_row.addWidget(QLabel("Status:"))
        self.excel_status = QLabel("Not connected")
        self.excel_status.setStyleSheet("color: #ff6b6b; font-weight: bold;")
        status_row.addWidget(self.excel_status)
        status_row.addStretch()
        
        excel_layout.addLayout(status_row)
        
        # Workbook info
        self.workbook_label = QLabel("No workbook selected")
        self.workbook_label.setProperty("subheading", True)
        excel_layout.addWidget(self.workbook_label)
        
        # Workbook selection
        select_row = QHBoxLayout()
        
        select_row.addWidget(QLabel("Workbook:"))
        self.workbook_combo = QComboBox()
        self.workbook_combo.setMinimumWidth(200)
        select_row.addWidget(self.workbook_combo)
        
        self.refresh_btn = QPushButton("↻ Refresh")
        self.refresh_btn.setProperty("secondary", True)
        self.refresh_btn.clicked.connect(self._refresh_workbooks)
        select_row.addWidget(self.refresh_btn)
        
        select_row.addStretch()
        
        excel_layout.addLayout(select_row)
        
        # Connect/Disconnect buttons
        btn_row = QHBoxLayout()
        
        self.connect_btn = QPushButton("Connect to Excel")
        self.connect_btn.clicked.connect(self._connect_excel)
        btn_row.addWidget(self.connect_btn)
        
        self.disconnect_btn = QPushButton("Disconnect")
        self.disconnect_btn.setProperty("secondary", True)
        self.disconnect_btn.clicked.connect(self._disconnect_excel)
        self.disconnect_btn.setEnabled(False)
        btn_row.addWidget(self.disconnect_btn)
        
        btn_row.addStretch()
        
        excel_layout.addLayout(btn_row)
        
        layout.addWidget(excel_group)
        
        # Color Palette Section
        colors_group = QGroupBox("Highlight Colors")
        colors_layout = QVBoxLayout(colors_group)
        
        colors_info = QLabel(
            "Colors are automatically assigned to finalized combinations.\n"
            "Colors cycle through the palette below:"
        )
        colors_info.setProperty("subheading", True)
        colors_layout.addWidget(colors_info)
        
        # Color swatches grid
        swatches_layout = QGridLayout()
        swatches_layout.setSpacing(6)
        
        cols = 10
        for i, (color, name) in enumerate(HIGHLIGHT_COLORS):
            row = i // cols
            col = i % cols
            swatch = ColorSwatch(color, name)
            swatches_layout.addWidget(swatch, row, col)
        
        colors_layout.addLayout(swatches_layout)
        
        layout.addWidget(colors_group)
        
        # About Section
        about_group = QGroupBox("About")
        about_layout = QVBoxLayout(about_group)
        
        about_text = QLabel(
            "<b>CombiMatch</b> - Invoice Reconciliation Tool<br><br>"
            "Find combinations of numbers that sum to a target value "
            "with tolerance support.<br><br>"
            "<b>Features:</b><br>"
            "• Manual input (line or comma separated)<br>"
            "• Excel integration with filter-aware selection<br>"
            "• Automatic highlighting in Excel<br>"
            "• Multiple finalized combinations with unique colors<br><br>"
            "<b>Tip:</b> Start by loading numbers, then set your target "
            "sum and buffer to find matching combinations."
        )
        about_text.setWordWrap(True)
        about_text.setTextFormat(Qt.RichText)
        about_layout.addWidget(about_text)
        
        layout.addWidget(about_group)
        
        # Stretch at bottom
        layout.addStretch()
    
    def _refresh_excel_status(self):
        """Refresh the Excel connection status."""
        handler = get_excel_handler()
        
        if handler.is_connected():
            self.excel_status.setText("Connected")
            self.excel_status.setStyleSheet("color: #51cf66; font-weight: bold;")
            self.connect_btn.setEnabled(False)
            self.disconnect_btn.setEnabled(True)
            
            # Update workbook info
            if handler.connection.workbook_name:
                self.workbook_label.setText(
                    f"Active: {handler.connection.workbook_name} "
                    f"({handler.connection.sheet_name})"
                )
        else:
            self.excel_status.setText("Not connected")
            self.excel_status.setStyleSheet("color: #ff6b6b; font-weight: bold;")
            self.connect_btn.setEnabled(True)
            self.disconnect_btn.setEnabled(False)
            self.workbook_label.setText("No workbook selected")
    
    def _refresh_workbooks(self):
        """Refresh the list of open workbooks."""
        handler = get_excel_handler()
        
        if not handler.is_connected():
            if not handler.connect():
                QMessageBox.information(
                    self, "Excel Not Running",
                    "Excel is not running or no workbooks are open."
                )
                return
        
        workbooks = handler.get_open_workbooks()
        
        self.workbook_combo.clear()
        if workbooks:
            self.workbook_combo.addItems(workbooks)
        else:
            self.workbook_combo.addItem("(No workbooks open)")
        
        self._refresh_excel_status()
    
    def _connect_excel(self):
        """Connect to Excel."""
        handler = get_excel_handler()
        
        if not ExcelHandler.is_available():
            QMessageBox.warning(
                self, "Not Available",
                "Excel automation is not available.\n"
                "This feature requires Windows and pywin32."
            )
            return
        
        if handler.connect():
            self._refresh_workbooks()
            
            # Auto-select first workbook
            workbooks = handler.get_open_workbooks()
            if workbooks:
                handler.select_workbook(workbooks[0])
            
            self._refresh_excel_status()
            
            QMessageBox.information(
                self, "Connected",
                "Successfully connected to Excel."
            )
        else:
            QMessageBox.warning(
                self, "Connection Failed",
                "Could not connect to Excel.\n"
                "Make sure Excel is running with at least one workbook open."
            )
    
    def _disconnect_excel(self):
        """Disconnect from Excel."""
        handler = get_excel_handler()
        handler.disconnect()
        
        self.workbook_combo.clear()
        self._refresh_excel_status()
    
    def check_excel_closed(self) -> bool:
        """
        Check if Excel has been closed unexpectedly.
        
        Returns:
            True if Excel was connected but is now closed
        """
        handler = get_excel_handler()
        
        if handler.connection.is_connected and not handler.is_connected():
            # Was connected but now isn't
            self._refresh_excel_status()
            return True
        
        return False
