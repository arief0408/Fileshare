import sys
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment
import ctypes
import ctypes.wintypes

def get_text_from_clipboard():
    CF_UNICODETEXT = 13  # Constant for Unicode text format
    kernel32 = ctypes.windll.kernel32
    user32 = ctypes.windll.user32

    # Open the clipboard
    if not user32.OpenClipboard(0):
        raise RuntimeError("Failed to open clipboard")

    try:
        # Check if the clipboard contains Unicode text
        if not user32.IsClipboardFormatAvailable(CF_UNICODETEXT):
            raise ValueError("No text found in clipboard.")
        
        # Get the handle to the clipboard data in CF_UNICODETEXT format
        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            raise RuntimeError("Failed to get clipboard data")

        # Lock the clipboard data to retrieve the text
        data_locked = kernel32.GlobalLock(handle)
        if not data_locked:
            raise RuntimeError("Failed to lock clipboard data")

        # Extract the text
        text = ctypes.wstring_at(data_locked)

        # Unlock the clipboard data
        kernel32.GlobalUnlock(handle)

        return text
    finally:
        # Ensure the clipboard is closed
        user32.CloseClipboard()

def write_text_to_excel(excel_file, sheet_name, row, col, text):
    # Check if the Excel file exists
    if os.path.exists(excel_file):
        # Load the existing workbook
        wb = load_workbook(excel_file)
        # If the sheet doesn't exist, create it
        if sheet_name not in wb.sheetnames:
            wb.create_sheet(sheet_name)
        sheet = wb[sheet_name]
    else:
        # Create a new workbook and add the specified sheet
        wb = Workbook()
        sheet = wb.active
        sheet.title = sheet_name

    # Write the text into the specified cell
    cell = sheet[f"{col}{row}"]
    cell.value = text

    # Set the column width to 80
    sheet.column_dimensions[col].width = 80

    # Set the font to Courier New
    cell.font = Font(name='Courier New', size=12)

    # Set text to wrap automatically
    cell.alignment = Alignment(wrap_text=True)

    # Save the updated or new Excel file
    wb.save(excel_file)

if __name__ == "__main__":
    # Check the number of arguments
    if len(sys.argv) < 6:
        print("Usage: python write_excel.py <sheet_name> <row> <col> <excel_file> [<text> | clipboard]")
        sys.exit(1)

    sheet_name = sys.argv[1]  # Sheet name
    row = sys.argv[2]  # Row number
    col = sys.argv[3]  # Column letter
    excel_file = sys.argv[4]  # Excel File
    if sys.argv[5].lower() == "clipboard":
        # Get text from clipboard
        try:
            clipboard_text = get_text_from_clipboard()
            write_text_to_excel(excel_file, sheet_name, row, col, clipboard_text)
            print(f"Clipboard text written into {sheet_name} at {col}{row}, with Courier New font, column width set to 80, and text wrapping enabled.")
        except ValueError as e:
            print(e)
            sys.exit(1)
        except RuntimeError as e:
            print(f"Error accessing clipboard: {e}")
            sys.exit(1)
    else:
        # Write the provided text directly
        text = " ".join(sys.argv[5:])  # Join the rest of the arguments as text
        write_text_to_excel(excel_file, sheet_name, row, col, text)
        print(f"Text written into {sheet_name} at {col}{row}, with Courier New font, column width set to 80, and text wrapping enabled.")
