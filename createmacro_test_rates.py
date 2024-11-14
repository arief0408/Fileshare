import pandas as pd
import sys
import os

def generate_xml_script_header(file_path):
    xml_content = f'''<HAScript name="create_proposal" description="" timeout="60000" pausetime="300" promptall="true" blockinput="true" author="" creationdate="" supressclearevents="false" usevars="true" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true" continueontimeout="false">
'''
    with open(file_path, "w") as file:
        file.write(xml_content)

def generate_xml_script_footer(file_path):
    xml_content = f'''</HAScript>
    '''
    with open(file_path, "a") as file:
        file.write(xml_content)

def generate_xml_script(file_path, table_value, item_value, item_value_next, exit_screen, entry_screen,output_string):
    xml_content = f'''<screen name="Table_Code_Maintenance_submenu_{item_value}" entryscreen="{entry_screen}" exitscreen="false" transient="false">
    <description >
        <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
    </description>
    <actions>
        <input value="&apos;{table_value}&apos;" row="17" col="35" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;{item_value}&apos;" row="18" col="35" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;E&apos;" row="22" col="35" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;[enter]&apos;" row="22" col="35" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <pause value="200"/>
    </actions>
    <nextscreens timeout="0" >
        <nextscreen name="Input_Work_Unit_{item_value}" />
    </nextscreens>
    <recolimit value="10000" />
</screen>

<screen name="Input_Work_Unit_{item_value}" entryscreen="false" exitscreen="false" transient="false">
    <description >
        <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
    </description>
    <actions>
        <input value="&apos;1&apos;" row="10" col="23" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;[enter]&apos;" row="10" col="23" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <pause value="100"/>
    </actions>
    <nextscreens timeout="0" >
        <nextscreen name="Table_Item_Maintenance_{item_value}" />
    </nextscreens>
    <recolimit value="10000" />
</screen>

<screen name="Table_Item_Maintenance_{item_value}" entryscreen="false" exitscreen="false" transient="false">
    <description >
        <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
    </description>
    <actions>
        {output_string}
        <input value="&apos;[enter]&apos;" row="16" col="31" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <pause value="400"/>
        <input value="&apos;[enter]&apos;" row="16" col="31" movecursor="true" xlatehostkeys="true" encrypted="false" />
    </actions>
    <nextscreens timeout="0" >
        <nextscreen name="Table_Code_Maintenance_submenu_{item_value_next}" />
    </nextscreens>
    <recolimit value="10000" />
</screen>

'''
    with open(file_path, "a") as file:
        file.write(xml_content)

import pandas as pd
import sys
import os

def read_excel_and_generate_xml(file_path, output_path):
    # Read the Excel file
    df = pd.read_excel(file_path)
    entry_screen = 'false'
    exit_screen = 'false'
    generate_xml_script_header(output_path)

    # Iterate over each row
    for index, row in df.iterrows():
        table_value = row['Table']
        item_value = row['Item']
        
        # Dictionary to hold year and flat values dynamically
        values = {
            "year_1_value": row['To_Year_1'],
            "flat_1_value": row['Flat_1'],
            "year_2_value": row['To_Year_2'],
            "flat_2_value": row['Flat_2'],
            "year_3_value": row['To_Year_3'],
            "flat_3_value": row['Flat_3'],
            "year_4_value": row['To_Year_4'],
            "flat_4_value": row['Flat_4'],
            "year_5_value": row['To_Year_5'],
            "flat_5_value": row['Flat_5'],
            "year_6_value": row['To_Year_6'],
            "flat_6_value": row['Flat_6'],
            "year_7_value": row['To_Year_7'],
            "flat_7_value": row['Flat_7'],
            "year_8_value": row['To_Year_8'],
            "flat_8_value": row['Flat_8'],
            "year_9_value": row['To_Year_9'],
            "flat_9_value": row['Flat_9'],
            "year_10_value": row['To_Year_10'],
            "flat_10_value": row['Flat_10']
        }

        output_string = ""
        row_value = 11
        
        # Loop through 1 to 10 for year and flat pairs
        for i in range(1, 11):
            year_value = values.get(f"year_{i}_value")
            flat_value = values.get(f"flat_{i}_value")
            # Skip the iteration if either year_value or flat_value is NaN or None
            if pd.isna(year_value) or pd.isna(flat_value):
                continue  # Skips the current iteration if either value is NaN

            # Convert values to strings and remove '.0' if it exists
            year_value = str(year_value).rstrip('.0') if year_value is not None else ''
            flat_value = str(flat_value).rstrip('.0') if flat_value is not None else ''
            # Only proceed if both values are non-NaN
            if pd.notna(year_value) and pd.notna(flat_value):
                output_string += f"""
                    <input value="&apos;{year_value}&apos;" row="{row_value}" col="04" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="&apos;{flat_value}&apos;" row="{row_value}" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="&apos;Y&apos;" row="{row_value}" col="60" movecursor="true" xlatehostkeys="true" encrypted="false" />
                """
                row_value += 1

        # Print for debugging purposes
        print(f"Row {index}:\n{output_string}\n")

        # Check if there's a next row for the item_value_next
        if index + 1 < len(df):
            item_value_next = df.iloc[index + 1]['Item']
            if index == 0:
                entry_screen = 'true'
        else:
            item_value_next = "End"  # Default value if there's no next row
        if index == len(df) - 1:
            exit_screen = 'true'
        
        generate_xml_script(output_path, table_value, item_value, item_value_next, exit_screen, entry_screen, output_string)

    generate_xml_script_footer(output_path)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py <excel_file_path>")
    else:
        excel_file_path = sys.argv[1]

        # Define the output file path
        output_file_path = f'generated_xml_script.mac'

        # Generate the XML script
        read_excel_and_generate_xml(excel_file_path, output_file_path)

        print(f"XML script generated and saved to {output_file_path}")
