import pandas as pd
import numpy as np

# Load the Excel file into a DataFrame
excel_file = 'Macro_Table.xlsx'  # Replace with your actual file name
excel_data = pd.ExcelFile(excel_file)


# Read the .txt file containing the template
with open('mac_template.txt', 'r', encoding='utf-8') as template_file:
    mac_template = template_file.read()

# Iterate through all sheets
for sheet_index,sheet_name in enumerate(excel_data.sheet_names, start=1):
    df = excel_data.parse(sheet_name)
    output_string = ''

    for index, row in df.iterrows():
        item_value = row.get('Premium', '')  # Replace 'Premium' with your actual column name
        Row = row.get('Row', '')            # Replace 'Row' with your actual column name
        Col = row.get('Col', '')            # Replace 'Col' with your actual column name
        Col_select=Col+5

        if pd.notna(item_value) and str(item_value).strip():
            if item_value == "DELETE":
                # Add actions for DELETE condition
                output_string += f"""
                    <pause value="200"/>
                    <boxselection type="SELECT" srow="{Row}" scol="{Col}" erow="{Row}" ecol="{Col_select}" />
                    <pause value="200"/>
                    <input value="&apos;[cut]&apos;" row="{Row}" col="{Col}" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <pause value="200"/>
                """
            else:
                # Add actions for non-DELETE condition
                output_string += f"""
                    <pause value="200"/>
                    <boxselection type="SELECT" srow="{Row}" scol="{Col}" erow="{Row}" ecol="{Col_select}" />
                    <pause value="200"/>
                    <input value="&apos;[cut]&apos;" row="{Row}" col="{Col}" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <pause value="200"/>
                    <input value="&apos;{item_value}&apos;" row="{Row}" col="{Col}" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <pause value="200"/>
                """

    # Format the .mac content for the sheet, including item_value
    formatted_content = mac_template.format(
        item_value=sheet_name,  # Use sheet name as the item_value or customize as needed
        output_string=output_string,
        TC_Next=index + 2,
        sheet_index=sheet_index,
        sheet_index_next=sheet_index+1
    )

    # Save to a .mac file named after the sheet
    mac_file_name = f'bospo_{sheet_index}.mac'
    with open(mac_file_name, 'w', encoding='utf-8') as file:
        file.write(formatted_content)

