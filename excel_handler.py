"""
excel_handler.py - Excel COM Automation for CombiMatch

Handles all Excel interactions via win32com:
- Detecting open workbooks
- Reading selected cells (filter-aware)
- Highlighting cells
- Auto-saving workbooks
"""

import pythoncom
from typing import List, Optional, Tuple, Dict, Any
from models import NumberItem, ItemSource, ExcelConnection

# Try to import win32com - will fail on non-Windows
try:
    import win32com.client as win32
    from pywintypes import com_error
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False
    com_error = Exception


class ExcelHandler:
    """
    Handles Excel COM automation for reading and highlighting cells.
    
    Attributes:
        app: Excel Application COM object
        connection: Current connection info
    """
    
    def __init__(self):
        """Initialize the handler."""
        self.app = None
        self.connection = ExcelConnection()
        self._workbook = None
        self._sheet = None
    
    @staticmethod
    def is_available() -> bool:
        """Check if Excel automation is available."""
        return EXCEL_AVAILABLE
    
    def connect(self) -> bool:
        """
        Connect to a running Excel instance.
        
        Returns:
            True if connected successfully
        """
        if not EXCEL_AVAILABLE:
            return False
        
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            # Get running Excel instance
            self.app = win32.GetActiveObject("Excel.Application")
            self.connection.is_connected = True
            return True
            
        except com_error:
            self.connection.is_connected = False
            return False
    
    def disconnect(self):
        """Disconnect from Excel."""
        self.app = None
        self._workbook = None
        self._sheet = None
        self.connection = ExcelConnection()
        
        try:
            pythoncom.CoUninitialize()
        except:
            pass
    
    def is_connected(self) -> bool:
        """Check if still connected to Excel."""
        if not self.app:
            return False
        
        try:
            # Try to access a property to verify connection
            _ = self.app.Visible
            return True
        except:
            self.connection.is_connected = False
            return False
    
    def get_open_workbooks(self) -> List[str]:
        """
        Get list of currently open workbook names.
        
        Returns:
            List of workbook names
        """
        if not self.is_connected():
            return []
        
        try:
            workbooks = []
            for wb in self.app.Workbooks:
                workbooks.append(wb.Name)
            return workbooks
        except:
            return []
    
    def select_workbook(self, workbook_name: str) -> bool:
        """
        Select a specific workbook to work with.
        
        Args:
            workbook_name: Name of the workbook
            
        Returns:
            True if selected successfully
        """
        if not self.is_connected():
            return False
        
        try:
            self._workbook = self.app.Workbooks(workbook_name)
            self._sheet = self._workbook.ActiveSheet
            
            self.connection.workbook_name = workbook_name
            self.connection.sheet_name = self._sheet.Name
            self.connection.file_path = self._workbook.FullName
            
            return True
        except:
            return False
    
    def get_sheets(self, workbook_name: str = None) -> List[str]:
        """
        Get list of sheet names in a workbook.
        
        Args:
            workbook_name: Name of workbook, or None for current
            
        Returns:
            List of sheet names
        """
        if not self.is_connected():
            return []
        
        try:
            if workbook_name:
                wb = self.app.Workbooks(workbook_name)
            elif self._workbook:
                wb = self._workbook
            else:
                wb = self.app.ActiveWorkbook
            
            if wb:
                return [sheet.Name for sheet in wb.Sheets]
        except:
            pass
        return []
    
    def select_sheet(self, sheet_name: str) -> bool:
        """
        Select a specific sheet in the current workbook.
        
        Args:
            sheet_name: Name of the sheet to select
            
        Returns:
            True if selected successfully
        """
        if not self._workbook:
            return False
        
        try:
            self._sheet = self._workbook.Sheets(sheet_name)
            self._sheet.Activate()
            self.connection.sheet_name = sheet_name
            return True
        except:
            return False
    
    def get_active_sheet_name(self) -> Optional[str]:
        """
        Get the name of the currently active sheet.
        
        Returns:
            Sheet name or None
        """
        if not self._workbook:
            return None
        
        try:
            return self._workbook.ActiveSheet.Name
        except:
            return None
    
    def get_active_workbook_info(self) -> Optional[Dict[str, str]]:
        """
        Get info about the currently active workbook.
        
        Returns:
            Dict with 'name', 'sheet', 'path' or None
        """
        if not self.is_connected():
            return None
        
        try:
            wb = self.app.ActiveWorkbook
            if wb:
                return {
                    'name': wb.Name,
                    'sheet': wb.ActiveSheet.Name,
                    'path': wb.FullName
                }
        except:
            pass
        return None
    
    def auto_save(self) -> bool:
        """
        Save the current workbook.
        
        Returns:
            True if saved successfully
        """
        if not self._workbook:
            return False
        
        try:
            self._workbook.Save()
            self.connection.is_saved = True
            return True
        except:
            return False
    
    def read_selection(self) -> Tuple[List[NumberItem], List[str]]:
        """
        Read numbers from the current Excel selection.
        
        Respects filters - only reads visible cells.
        Uses bulk reading for performance.
        
        Returns:
            Tuple of (list of NumberItem, list of warning messages)
        """
        items = []
        warnings = []
        
        if not self.is_connected() or not self._workbook:
            return items, ["Not connected to Excel"]
        
        try:
            selection = self.app.Selection
            
            if selection is None:
                return items, ["No cells selected"]
            
            # Check for multiple columns
            columns_used = set()
            
            # Check if single cell selected - handle directly without SpecialCells
            try:
                if selection.Count == 1:
                    # Single cell - process directly
                    row = selection.Row
                    col = selection.Column
                    col_letter = self._column_letter(col)
                    value = selection.Value
                    
                    if value is not None:
                        if isinstance(value, str):
                            try:
                                value = float(value.replace(',', ''))
                            except ValueError:
                                return items, ["Selected cell is not a number"]
                        
                        try:
                            numeric_value = float(value)
                            item = NumberItem(
                                value=numeric_value,
                                index=0,
                                row=row,
                                column=col_letter,
                                source=ItemSource.EXCEL
                            )
                            return [item], []
                        except (TypeError, ValueError):
                            return items, ["Selected cell is not a number"]
                    else:
                        return items, ["Selected cell is empty"]
            except:
                pass  # Not a single cell or Count failed, continue with normal flow
            
            # Get visible cells only (respects filters)
            try:
                visible_cells = selection.SpecialCells(12)  # xlCellTypeVisible = 12
            except:
                visible_cells = selection
            
            index = 0
            
            # Collect cell info - process each area
            for area in visible_cells.Areas:
                try:
                    # Get area dimensions
                    start_row = area.Row
                    start_col = area.Column
                    num_rows = area.Rows.Count
                    num_cols = area.Columns.Count
                    
                    # Get all values in one call
                    raw_values = area.Value
                    
                    # Convert to consistent 2D list format
                    if raw_values is None:
                        continue
                    elif num_rows == 1 and num_cols == 1:
                        # Single cell - just a value
                        values = [[raw_values]]
                    elif num_rows == 1:
                        # Single row - tuple of values
                        if isinstance(raw_values, tuple):
                            values = [list(raw_values)]
                        else:
                            values = [[raw_values]]
                    elif num_cols == 1:
                        # Single column - tuple of single-value tuples
                        if isinstance(raw_values, tuple):
                            values = [[v] if not isinstance(v, tuple) else [v[0]] for v in raw_values]
                        else:
                            values = [[raw_values]]
                    else:
                        # Multiple rows and columns - tuple of tuples
                        if isinstance(raw_values, tuple):
                            values = [list(row) if isinstance(row, tuple) else [row] for row in raw_values]
                        else:
                            values = [[raw_values]]
                    
                    # Process the values
                    for r_idx, row_values in enumerate(values):
                        if row_values is None:
                            continue
                        for c_idx, value in enumerate(row_values):
                            actual_row = start_row + r_idx
                            actual_col = start_col + c_idx
                            col_letter = self._column_letter(actual_col)
                            columns_used.add(actual_col)
                            
                            # Skip empty values
                            if value is None:
                                continue
                            
                            # Try to convert to number
                            if isinstance(value, str):
                                try:
                                    value = float(value.replace(',', ''))
                                except ValueError:
                                    continue
                            
                            try:
                                numeric_value = float(value)
                                
                                item = NumberItem(
                                    value=numeric_value,
                                    index=index,
                                    row=actual_row,
                                    column=col_letter,
                                    source=ItemSource.EXCEL
                                )
                                items.append(item)
                                index += 1
                                
                            except (TypeError, ValueError):
                                continue
                                
                except:
                    # If bulk read fails for this area, fall back to cell-by-cell
                    try:
                        for cell in area:
                            row = cell.Row
                            col = cell.Column
                            col_letter = self._column_letter(col)
                            columns_used.add(col)
                            
                            value = cell.Value
                            if value is None:
                                continue
                            
                            if isinstance(value, str):
                                try:
                                    value = float(value.replace(',', ''))
                                except ValueError:
                                    continue
                            
                            try:
                                numeric_value = float(value)
                                item = NumberItem(
                                    value=numeric_value,
                                    index=index,
                                    row=row,
                                    column=col_letter,
                                    source=ItemSource.EXCEL
                                )
                                items.append(item)
                                index += 1
                            except (TypeError, ValueError):
                                continue
                    except:
                        continue
            
            # Warn if multiple columns
            if len(columns_used) > 1:
                cols = sorted(columns_used)
                col_names = [self._column_letter(c) for c in cols]
                warnings.append(f"Selection spans multiple columns: {', '.join(col_names)}")
            
            return items, warnings
            
        except Exception as e:
            return items, [f"Error reading selection: {str(e)}"]
            
    def highlight_cell(
        self,
        row: int,
        column: str,
        color: Tuple[int, int, int]
    ) -> bool:
        """
        Highlight a specific cell with a color.
        
        Args:
            row: Row number
            column: Column letter (e.g., 'A', 'B', 'AA')
            color: RGB tuple (r, g, b) where each is 0-255
            
        Returns:
            True if highlighted successfully
        """
        if not self._sheet:
            return False
        
        try:
            cell_address = f"{column}{row}"
            cell = self._sheet.Range(cell_address)
            
            # Excel uses BGR format for colors
            r, g, b = color
            excel_color = r + (g * 256) + (b * 256 * 256)
            
            cell.Interior.Color = excel_color
            return True
            
        except:
            return False
    
    def highlight_items(
        self,
        items: List[NumberItem],
        color: Tuple[int, int, int]
    ) -> int:
        """
        Highlight multiple items in Excel.
        
        Args:
            items: List of NumberItem objects to highlight
            color: RGB color tuple
            
        Returns:
            Number of cells successfully highlighted
        """
        count = 0
        for item in items:
            if item.source == ItemSource.EXCEL and item.row and item.column:
                if self.highlight_cell(item.row, item.column, color):
                    count += 1
        return count
    
    def clear_highlight(self, row: int, column: str) -> bool:
        """
        Remove highlight from a cell (set to no fill).
        
        Args:
            row: Row number
            column: Column letter
            
        Returns:
            True if cleared successfully
        """
        if not self._sheet:
            return False
        
        try:
            cell_address = f"{column}{row}"
            cell = self._sheet.Range(cell_address)
            cell.Interior.ColorIndex = 0  # No fill
            return True
        except:
            return False
    
    def save_workbook(self) -> bool:
        """
        Save the current workbook.
        
        Returns:
            True if saved successfully
        """
        if not self._workbook:
            return False
        
        try:
            self._workbook.Save()
            return True
        except:
            return False
    
    @staticmethod
    def _column_letter(col_num: int) -> str:
        """
        Convert column number to letter(s).
        
        Args:
            col_num: Column number (1-based)
            
        Returns:
            Column letter(s) like 'A', 'B', 'AA', etc.
        """
        result = ""
        while col_num > 0:
            col_num, remainder = divmod(col_num - 1, 26)
            result = chr(65 + remainder) + result
        return result
    
    @staticmethod
    def _column_number(col_letter: str) -> int:
        """
        Convert column letter(s) to number.
        
        Args:
            col_letter: Column letter(s) like 'A', 'B', 'AA'
            
        Returns:
            Column number (1-based)
        """
        result = 0
        for char in col_letter.upper():
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result


# Singleton instance for the application
_excel_handler: Optional[ExcelHandler] = None


def get_excel_handler() -> ExcelHandler:
    """Get the singleton Excel handler instance."""
    global _excel_handler
    if _excel_handler is None:
        _excel_handler = ExcelHandler()
    return _excel_handler
