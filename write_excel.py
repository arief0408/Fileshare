import sys
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
import win32clipboard

def get_text_from_clipboard():
    try:
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_TEXT):
            clipboard_data = win32clipboard.GetClipboardData(win32clipboard.CF_TEXT)
            return clipboard_data.decode('utf-8')
        else:
            raise ValueError("No text found in clipboard.")
    finally:
        win32clipboard.CloseClipboard()

def write_text_to_excel(excel_file, sheet_name, row, col, text):
    if os.path.exists(excel_file):
        wb = load_workbook(excel_file)
        if sheet_name not in wb.sheetnames:
            wb.create_sheet(sheet_name)
        sheet = wb[sheet_name]
    else:
        wb = Workbook()
        sheet = wb.active
        sheet.title = sheet_name

    cell = sheet[f"{col}{row}"]
    cell.value = text
    sheet.column_dimensions[col].width = 80
    cell.font = Font(name='Courier New', size=12)
    cell.alignment = Alignment(wrap_text=True)
    wb.save(excel_file)

if __name__ == "__main__":
    if len(sys.argv) < 6:
        print("Usage: python screenshot.py <sheet_name> <row> <col> <excel_file> [<text> | clipboard]")
        sys.exit(1)

    sheet_name = sys.argv[1]
    row = sys.argv[2]
    col = sys.argv[3]
    excel_file = sys.argv[4]

    if sys.argv[5].lower() == "clipboard":
        try:
            clipboard_text = get_text_from_clipboard()
            write_text_to_excel(excel_file, sheet_name, row, col, clipboard_text)
            print(f"Clipboard text written into {sheet_name} at {col}{row}, with Courier New font, column width set to 80, and text wrapping enabled.")
        except ValueError as e:
            print(e)
            sys.exit(1)
    else:
        text = " ".join(sys.argv[5:])
        write_text_to_excel(excel_file, sheet_name, row, col, text)
        print(f"Text written into {sheet_name} at {col}{row}, with Courier New font, column width set to 80, and text wrapping enabled.")
