# CombiMatch - Invoice Reconciliation Tool

A Windows desktop application for finding combinations of numbers that sum to a target value. Perfect for invoice reconciliation, bank statement matching, and similar tasks.

## Features

- **Subset Sum Solver**: Find combinations of numbers matching a target sum with buffer tolerance
- **Multiple Input Modes**: 
  - Line-separated numbers
  - Comma-separated numbers  
  - Direct Excel selection (filter-aware)
- **Excel Integration**:
  - Read selected cells from open Excel files
  - Respects filtered views (only reads visible cells)
  - Auto-save before modifications
  - Highlight matched cells directly in Excel
- **Visual Workflow**:
  - Results sorted by combination size
  - Preview combinations before finalizing
  - Unique colors for each finalized combination
  - Source list shows original order of numbers
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
   - Results appear sorted by size (smallest first)

4. **Review & Finalize**
   - Click a result to highlight its numbers in the source list
   - Click "Finalize Selected" to lock in the match
   - Finalized numbers get a unique color and are excluded from future searches

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
└── ui/
    ├── __init__.py      # UI package
    ├── main_window.py   # Main application window
    ├── find_tab.py      # Find combinations tab
    ├── summary_tab.py   # Finalized combinations summary
    ├── settings_tab.py  # Settings and Excel connection
    └── styles.py        # UI styling and theme
```

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
