import xlwings as xw
import pandas as pd

# Input file path
instruction_file_path = 'instructions.xlsx'

# Read instructions into a DataFrame
instructions = pd.read_excel(instruction_file_path)

# Iterate through each row of instructions
for index, row in instructions.iterrows():
    try:
        # Open the specified Excel file
        excelname = row["excelname"]
        app = xw.App(visible=False)  # Open Excel in the background
        wb = xw.Book(excelname)

        # Iterate through all 13 sets of instructions
        for i in range(1, 14):
            sheet_key = f"sheet_{i}"
            row_key = f"row_{i}"
            cell_key = f"cell_{i}"
            value_key = f"value_{i}"

            # Check if the current instruction set exists
            if pd.notna(row.get(sheet_key)) and pd.notna(row.get(row_key)) and pd.notna(row.get(cell_key)):
                sheet = row[sheet_key]
                row_num = int(row[row_key])
                cell = row[cell_key]
                value = row[value_key]

                # Update the cell with the value
                wb.sheets[sheet].range((row_num, cell)).value = value

        # Force recalculation
        wb.app.calculate()

        # Retrieve the result from the output sheet
        sheet_output = row["sheet_output"]
        row_output = int(row["row_output"])
        cell_output = row["cell_output"]
        result = wb.sheets[sheet_output].range((row_output, cell_output)).value

        # Store the result in the DataFrame
        instructions.loc[index, "val_output"] = result

        # Save and close the workbook
        wb.save()
        wb.close()

    except Exception as e:
        print(f"Error processing row {index}: {e}")

    finally:
        # Quit the Excel app
        app.quit()

# Save updated instructions with results
output_file_path = 'instructions.xlsx'
instructions.to_excel(output_file_path, index=False)
print(f"Updated instructions saved to {output_file_path}")
