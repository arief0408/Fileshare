import xlwings as xw

# Open the workbook
file_path = "your_excel_file.xlsx"
app = xw.App(visible=False)  # Open Excel in the background
wb = xw.Book(file_path)
sheet = wb.sheets["Sheet1"]

for i in range(10):
    # Input data into specific cells
    sheet.range("A2").value = i
    sheet.range("B2").value = i

    # Force recalculation
    wb.app.calculate()

    # Read the formula result
    result = sheet.range("C2").value
    print(f"Result from formula in C1: {result}")

# Save and close the workbook
wb.save(file_path)
wb.close()
app.quit()
