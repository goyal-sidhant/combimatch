# CombiMatch - Invoice Reconciliation Tool

A Windows desktop application for finding combinations of numbers that sum to a target value. Perfect for invoice reconciliation, bank statement matching, and similar tasks.

## Features

- **Subset Sum Solver**: Find combinations of numbers matching a target sum with buffer tolerance
- **Split Results View**: Exact matches and approximate matches displayed separately
- **Selected Combination Info Panel**: Detailed view of selected combination before finalizing
- **Multiple Input Modes**: 
  - Line-separated numbers
  - Comma-separated numbers  
  - Direct Excel selection (filter-aware)
- **Excel Integration**:
  - Read selected cells from open Excel files
  - Select specific workbook and sheet from Settings
  - Respects filtered views (only reads visible cells)
  - Auto-save before modifications
  - Highlight matched cells directly in Excel
- **Visual Workflow**:
  - Results sorted by combination size
  - Prominent orange highlight when selecting a combination
  - Bold text for selected numbers in source list
  - Preview combinations before finalizing
  - Unique colors for each finalized combination
  - Finalized numbers marked with ✓ and greyed out
  - Invalid combinations automatically removed after finalization
- **Clean UI**: Soft colors, multi-tab layout, resizable window

## Installation

### Prerequisites
- Python 3.8 or higher
- Windows (for Excel integration)

### Install Dependencies

```bash
pip install -r requirements.txt
```

Or manually:
```bash
pip install PyQt5 pywin32
```

## Usage

### Running the Application

```bash
python main.py
```

### Basic Workflow

1. **Load Numbers**
   - Type/paste numbers in the input area, OR
   - Select cells in Excel and click "Grab from Excel"

2. **Set Parameters**
   - Target Sum: The value you're trying to match
   - Buffer: Tolerance (±) for accepting close matches
   - Min/Max Numbers: Combination size limits
   - Max Results: How many combinations to find

3. **Find Combinations**
   - Click "Find Combinations"
   - Exact matches appear in top box, approximate in bottom box
   - Results sorted by size (smallest first)

4. **Review & Finalize**
   - Click a result to see details in the info panel
   - Selected numbers highlighted in orange with bold text
   - Click "Finalize Selected" to lock in the match
   - Finalized numbers get a unique color and ✓ marker
   - Combinations using finalized numbers are automatically removed

5. **Repeat**
   - Continue finding combinations with remaining numbers
   - Each finalized set gets a different color

### Input Modes

| Mode | Format |
|------|--------|
| Line Separated | One number per line |
| Comma Separated | Numbers separated by commas |
| Excel | Select cells in open Excel file |

### Excel Tips

- Go to Settings tab to select specific workbook and sheet
- Only visible cells are read (respects filters)
- Multi-column selections are read as a flat list (with warning)
- The app auto-saves your workbook before making changes
- Highlights are applied directly to Excel cells
- On close, you're prompted to save the Excel file

## File Structure

```
combimatch/
├── main.py              # Application entry point
├── models.py            # Data classes and types
├── solver.py            # Subset sum algorithm
├── excel_handler.py     # Excel COM automation
├── utils.py             # Utilities and color management
├── requirements.txt     # Python dependencies
├── README.md            # This file
├── .gitignore           # Git ignore file
└── ui/
    ├── __init__.py      # UI package
    ├── main_window.py   # Main application window
    ├── find_tab.py      # Find combinations tab
    ├── summary_tab.py   # Finalized combinations summary
    ├── settings_tab.py  # Settings and Excel connection
    └── styles.py        # UI styling and theme
```

## What's New in This Version

- **Split Results**: Exact and approximate matches in separate boxes
- **Info Panel**: Shows sum, difference, item count, and values before finalizing
- **Prominent Highlighting**: Orange background + bold text for selected items
- **Smart Cleanup**: Combinations using finalized numbers auto-removed
- **Sheet Selection**: Choose specific sheet in Excel from Settings
- **Visual Indicators**: ✓ marker and color for finalized numbers

## Keyboard Shortcuts

- Tab navigation: `Ctrl+Tab` / `Ctrl+Shift+Tab`
- Standard copy/paste in text fields

## Limitations

- Excel integration requires Windows
- Designed for ~100 numbers (larger lists may be slow)
- No undo for finalized combinations (by design)
- Session data is not saved between runs

## Troubleshooting

### "Could not connect to Excel"
- Make sure Excel is running with at least one workbook open
- Try clicking "Refresh" in Settings tab

### "No valid numbers in selection"
- Ensure selected cells contain numeric values
- Text, blanks, and error cells are skipped

### Slow performance
- Reduce "Max Numbers" in combination size
- Reduce "Max Results"
- Work with smaller number sets

## License

MIT License - Free to use and modify.
