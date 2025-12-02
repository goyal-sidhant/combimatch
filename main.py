"""
CombiMatch - Invoice Reconciliation Subset Sum Solver

Main entry point for the application.

Usage:
    python main.py

Requirements:
    - Python 3.8+
    - PyQt5
    - pywin32 (for Excel integration on Windows)

Install dependencies:
    pip install PyQt5 pywin32
"""

import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt

from ui import MainWindow, apply_styles


def main():
    """Main entry point."""
    # Enable high DPI scaling
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    
    # Create application
    app = QApplication(sys.argv)
    app.setApplicationName("CombiMatch")
    app.setOrganizationName("CombiMatch")
    
    # Apply styles
    apply_styles(app)
    
    # Create and show main window
    window = MainWindow()
    window.show()
    
    # Run event loop
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
