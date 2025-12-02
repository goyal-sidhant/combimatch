"""
styles.py - UI Styling for CombiMatch

Defines the visual theme with soft, muted colors
that are easy on the eyes.
"""

# Color palette - soft, muted tones
COLORS = {
    'background': '#f5f6f8',
    'surface': '#ffffff',
    'surface_dark': '#e8eaed',
    'primary': '#5c7cfa',
    'primary_hover': '#4263eb',
    'secondary': '#748ffc',
    'text': '#343a40',
    'text_light': '#868e96',
    'text_muted': '#adb5bd',
    'border': '#dee2e6',
    'success': '#51cf66',
    'warning': '#fcc419',
    'error': '#ff6b6b',
    'highlight': '#fff3bf',
    'selection': '#d0ebff',
}


def get_stylesheet() -> str:
    """
    Get the main application stylesheet.
    
    Returns:
        CSS-like stylesheet string for PyQt
    """
    return f"""
        /* Main Window */
        QMainWindow {{
            background-color: {COLORS['background']};
        }}
        
        /* General Widget */
        QWidget {{
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 10pt;
            color: {COLORS['text']};
        }}
        
        /* Tab Widget */
        QTabWidget::pane {{
            border: 1px solid {COLORS['border']};
            background-color: {COLORS['surface']};
            border-radius: 4px;
        }}
        
        QTabBar::tab {{
            background-color: {COLORS['surface_dark']};
            border: 1px solid {COLORS['border']};
            border-bottom: none;
            padding: 8px 16px;
            margin-right: 2px;
            border-top-left-radius: 4px;
            border-top-right-radius: 4px;
        }}
        
        QTabBar::tab:selected {{
            background-color: {COLORS['surface']};
            border-bottom: 1px solid {COLORS['surface']};
        }}
        
        QTabBar::tab:hover:!selected {{
            background-color: {COLORS['selection']};
        }}
        
        /* Buttons */
        QPushButton {{
            background-color: {COLORS['primary']};
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            font-weight: 500;
        }}
        
        QPushButton:hover {{
            background-color: {COLORS['primary_hover']};
        }}
        
        QPushButton:pressed {{
            background-color: {COLORS['secondary']};
        }}
        
        QPushButton:disabled {{
            background-color: {COLORS['text_muted']};
        }}
        
        QPushButton[secondary="true"] {{
            background-color: {COLORS['surface']};
            color: {COLORS['text']};
            border: 1px solid {COLORS['border']};
        }}
        
        QPushButton[secondary="true"]:hover {{
            background-color: {COLORS['surface_dark']};
        }}
        
        /* Input Fields */
        QLineEdit, QSpinBox, QDoubleSpinBox {{
            background-color: {COLORS['surface']};
            border: 1px solid {COLORS['border']};
            border-radius: 4px;
            padding: 6px 10px;
        }}
        
        QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus {{
            border-color: {COLORS['primary']};
        }}
        
        /* Text Areas */
        QTextEdit, QPlainTextEdit {{
            background-color: {COLORS['surface']};
            border: 1px solid {COLORS['border']};
            border-radius: 4px;
            padding: 8px;
        }}
        
        QTextEdit:focus, QPlainTextEdit:focus {{
            border-color: {COLORS['primary']};
        }}
        
        /* Lists */
        QListWidget {{
            background-color: {COLORS['surface']};
            border: 1px solid {COLORS['border']};
            border-radius: 4px;
            padding: 4px;
        }}
        
        QListWidget::item {{
            padding: 6px 8px;
            border-radius: 2px;
        }}
        
        QListWidget::item:selected {{
            background-color: {COLORS['selection']};
            color: {COLORS['text']};
        }}
        
        QListWidget::item:hover:!selected {{
            background-color: {COLORS['surface_dark']};
        }}
        
        /* Combo Box */
        QComboBox {{
            background-color: {COLORS['surface']};
            border: 1px solid {COLORS['border']};
            border-radius: 4px;
            padding: 6px 10px;
            min-width: 100px;
        }}
        
        QComboBox:hover {{
            border-color: {COLORS['primary']};
        }}
        
        QComboBox::drop-down {{
            border: none;
            width: 24px;
        }}
        
        QComboBox::down-arrow {{
            width: 12px;
            height: 12px;
        }}
        
        /* Labels */
        QLabel {{
            color: {COLORS['text']};
        }}
        
        QLabel[heading="true"] {{
            font-size: 12pt;
            font-weight: bold;
            color: {COLORS['text']};
        }}
        
        QLabel[subheading="true"] {{
            font-size: 9pt;
            color: {COLORS['text_light']};
        }}
        
        /* Group Box */
        QGroupBox {{
            font-weight: bold;
            border: 1px solid {COLORS['border']};
            border-radius: 4px;
            margin-top: 12px;
            padding-top: 8px;
        }}
        
        QGroupBox::title {{
            subcontrol-origin: margin;
            subcontrol-position: top left;
            left: 10px;
            padding: 0 6px;
            background-color: {COLORS['surface']};
        }}
        
        /* Scroll Bars */
        QScrollBar:vertical {{
            background-color: {COLORS['surface_dark']};
            width: 12px;
            border-radius: 6px;
        }}
        
        QScrollBar::handle:vertical {{
            background-color: {COLORS['text_muted']};
            border-radius: 6px;
            min-height: 30px;
        }}
        
        QScrollBar::handle:vertical:hover {{
            background-color: {COLORS['text_light']};
        }}
        
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
            height: 0;
        }}
        
        QScrollBar:horizontal {{
            background-color: {COLORS['surface_dark']};
            height: 12px;
            border-radius: 6px;
        }}
        
        QScrollBar::handle:horizontal {{
            background-color: {COLORS['text_muted']};
            border-radius: 6px;
            min-width: 30px;
        }}
        
        QScrollBar::handle:horizontal:hover {{
            background-color: {COLORS['text_light']};
        }}
        
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
            width: 0;
        }}
        
        /* Status Bar */
        QStatusBar {{
            background-color: {COLORS['surface_dark']};
            border-top: 1px solid {COLORS['border']};
        }}
        
        /* Message Box */
        QMessageBox {{
            background-color: {COLORS['surface']};
        }}
        
        /* Radio and Checkbox */
        QRadioButton, QCheckBox {{
            spacing: 8px;
        }}
        
        QRadioButton::indicator, QCheckBox::indicator {{
            width: 16px;
            height: 16px;
        }}
        
        /* Progress Bar */
        QProgressBar {{
            border: 1px solid {COLORS['border']};
            border-radius: 4px;
            background-color: {COLORS['surface_dark']};
            text-align: center;
        }}
        
        QProgressBar::chunk {{
            background-color: {COLORS['primary']};
            border-radius: 3px;
        }}
        
        /* Splitter */
        QSplitter::handle {{
            background-color: {COLORS['border']};
        }}
        
        QSplitter::handle:horizontal {{
            width: 2px;
        }}
        
        QSplitter::handle:vertical {{
            height: 2px;
        }}
    """


def apply_styles(app):
    """
    Apply the stylesheet to the application.
    
    Args:
        app: QApplication instance
    """
    app.setStyleSheet(get_stylesheet())


def get_color(name: str) -> str:
    """
    Get a color value by name.
    
    Args:
        name: Color name from COLORS dict
        
    Returns:
        Color hex string
    """
    return COLORS.get(name, COLORS['text'])
