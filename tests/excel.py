"""Excel COM helpers. DispatchEx starts a private Excel process, so a workbook the user has open is never touched."""
import win32com.client as w


def open_excel():
    xl = w.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    return xl


def sheet_errors(book):
    """(sheet, row, col) of every cell holding an Excel error value."""
    errs = []
    for ws in book.Worksheets:
        ur = ws.UsedRange
        vals = ur.Value
        if vals is None:
            continue
        if not isinstance(vals, tuple):
            vals = ((vals,),)
        for i, row in enumerate(vals):
            for j, v in enumerate(row):
                if isinstance(v, int) and v < -2146826000:
                    errs.append((ws.Name, ur.Row + i, ur.Column + j))
    return errs
