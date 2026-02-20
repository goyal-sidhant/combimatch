"""
main_window.py - Main Application Window for CombiMatch

The main window containing:
- Tab widget with Find, Summary, and Settings tabs
- Status bar
- Close event handling with save prompt
"""

from PyQt5.QtWidgets import (
    QMainWindow, QTabWidget, QStatusBar, QMessageBox,
    QWidget, QVBoxLayout
)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QCloseEvent

from .find_tab import FindTab
from .summary_tab import SummaryTab
from .settings_tab import SettingsTab
from excel_handler import get_excel_handler


class MainWindow(QMainWindow):
    """
    Main application window.
    """
    
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("CombiMatch - Invoice Reconciliation")
        self.setMinimumSize(1000, 700)
        
        # Center on screen
        self._center_on_screen()
        
        self._init_ui()
        self._connect_signals()
        self._start_excel_monitor()
    
    def _center_on_screen(self):
        """Center the window on the screen."""
        from PyQt5.QtWidgets import QDesktopWidget
        
        screen = QDesktopWidget().availableGeometry()
        size = self.geometry()
        
        x = (screen.width() - size.width()) // 2
        y = (screen.height() - size.height()) // 2
        
        self.move(x, y)
    
    def _init_ui(self):
        """Initialize the user interface."""
        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        
        layout = QVBoxLayout(central)
        layout.setContentsMargins(16, 16, 16, 16)
        
        # Tab widget
        self.tabs = QTabWidget()
        
        # Create tabs
        self.find_tab = FindTab()
        self.summary_tab = SummaryTab()
        self.settings_tab = SettingsTab()
        
        self.tabs.addTab(self.find_tab, "🔍 Find Combinations")
        self.tabs.addTab(self.summary_tab, "📋 Summary")
        self.tabs.addTab(self.settings_tab, "⚙️ Settings")
        
        layout.addWidget(self.tabs)
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")
    
    def _connect_signals(self):
        """Connect signals between components."""
        # Connect finalization signal to summary tab
        self.find_tab.combination_finalized.connect(
            self.summary_tab.add_finalized
        )

        # Auto-switch find tab to Excel mode when connected
        self.settings_tab.excel_connected.connect(
            self.find_tab.switch_to_excel_mode
        )

        # Update status on tab change
        self.tabs.currentChanged.connect(self._on_tab_changed)
    
    def _on_tab_changed(self, index: int):
        """Handle tab change."""
        tab_names = ["Find Combinations", "Summary", "Settings"]
        if 0 <= index < len(tab_names):
            self.status_bar.showMessage(f"{tab_names[index]} tab")
        
        # Refresh Excel status when going to settings
        if index == 2:  # Settings tab
            self.settings_tab._refresh_excel_status()
    
    def _start_excel_monitor(self):
        """Start monitoring Excel connection."""
        self.excel_timer = QTimer()
        self.excel_timer.timeout.connect(self._check_excel)
        self.excel_timer.start(5000)  # Check every 5 seconds
    
    def _check_excel(self):
        """Check if Excel is still connected."""
        if self.settings_tab.check_excel_closed():
            QMessageBox.warning(
                self, "Excel Disconnected",
                "Excel has been closed.\n"
                "Highlighting will no longer be applied to Excel."
            )
    
    def closeEvent(self, event: QCloseEvent):
        """Handle window close event."""
        handler = get_excel_handler()
        finalized = self.find_tab.get_finalized_combinations()
        
        # Check if there are highlights to save
        if handler.is_connected() and finalized:
            reply = QMessageBox.question(
                self,
                "Save Changes?",
                "Do you want to save the Excel file with highlights?",
                QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel
            )
            
            if reply == QMessageBox.Save:
                if handler.save_workbook():
                    self.status_bar.showMessage("Excel file saved")
                else:
                    QMessageBox.warning(
                        self, "Save Failed",
                        "Could not save the Excel file."
                    )
                    event.ignore()
                    return
            elif reply == QMessageBox.Cancel:
                event.ignore()
                return
        
        # Confirm close
        if finalized:
            confirm = QMessageBox.question(
                self,
                "Close Application?",
                f"You have {len(finalized)} finalized combination(s).\n"
                "Are you sure you want to close?",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if confirm == QMessageBox.No:
                event.ignore()
                return
        
        # Cleanup
        handler.disconnect()
        self.excel_timer.stop()
        
        event.accept()
