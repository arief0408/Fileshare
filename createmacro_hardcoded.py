import pandas as pd
import sys

def generate_xml_script_header():
    xml_content = f'''<HAScript name="create_proposal" description="" timeout="60000" pausetime="300" promptall="true" blockinput="true" author="" creationdate="" supressclearevents="false" usevars="true" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true" continueontimeout="false">
'''
    with open(f'generated_xml_script.mac', "w") as file:
        file.write(xml_content)

def generate_xml_script_footer():
    xml_content = f'''</HAScript>
    '''
    with open(f'generated_xml_script.mac', "a") as file:
        file.write(xml_content)

def generate_xml_script(table_value, item_value, long_desc_value, item_value_next,short_desc_value,valid_flag_value,work_unit_value,desc_1_xml,desc_2_xml,exit_screen,entry_screen):
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
        <input value="&apos;{work_unit_value}&apos;" row="11" col="41" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;[enter]&apos;" row="16" col="60" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <pause value="100"/>
        <input value="&apos;X&apos;" row="16" col="60" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;[enter]&apos;" row="16" col="60" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <pause value="200"/>
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
        <input value="&apos;{long_desc_value}&apos;" row="12" col="31" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;{short_desc_value}&apos;" row="14" col="31" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;{valid_flag_value}&apos;" row="16" col="31" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;[enter]&apos;" row="16" col="31" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <pause value="100"/>
    </actions>
    <nextscreens timeout="0" >
        <nextscreen name="Produce_Description_Edit_{item_value}" />
    </nextscreens>
    <recolimit value="10000" />
</screen>

<screen name="Produce_Description_Edit_{item_value}" entryscreen="false" exitscreen="{exit_screen}" transient="false">
    <description >
        <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
    </description>
    <actions>
        {desc_1_xml}
        {desc_2_xml}
        <input value="&apos;[enter]&apos;" row="13" col="08" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <pause value="200"/>
    </actions>
    <nextscreens timeout="0" >
        <nextscreen name="Table_Code_Maintenance_submenu_{item_value_next}"/>
    </nextscreens>
    <recolimit value="10000" />
</screen>
'''
    with open(f'generated_xml_script.mac', "a") as file:
        file.write(xml_content)

def read_excel_and_generate_xml(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)
    entry_screen = 'false'
    exit_screen = 'false'
    generate_xml_script_header()
    # Iterate over each row
    for index, row in df.iterrows():
        table_value = row['Table']
        item_value = row['Item']
        long_desc_value = row['Long_Desc']
        short_desc_value = row['Short_Desc']
        valid_flag_value = row['Valid_Flag']
        work_unit_value = row['Work_Unit']
        desc_1_value = row['Product_Desc_Line_1']
        desc_2_value = row['Product_Desc_Line_2']
        if desc_1_value =='[UNTOUCH]':
            desc_1_xml = ''
        else:
            desc_1_xml = f'<input value="&apos;{desc_1_value}&apos;" row="11" col="08" movecursor="true" xlatehostkeys="true" encrypted="false" />'
        if desc_2_value =='[UNTOUCH]':
            desc_2_xml = ''
        else:
            desc_2_xml = f'<input value="&apos;{desc_2_value}&apos;" row="13" col="08" movecursor="true" xlatehostkeys="true" encrypted="false" />'
        
        # Check if there's a next row for the item_value_next
        if index + 1 < len(df):
            item_value_next = df.iloc[index + 1]['Item']
            if index == 0:
                entry_screen = 'true'
        else:
            item_value_next = "End"  # Default value if there's no next row
        if index == len(df) - 1:
            exit_screen = 'true'
        
        
        generate_xml_script(table_value, item_value, long_desc_value, item_value_next,short_desc_value,valid_flag_value,work_unit_value,desc_1_xml,desc_2_xml,exit_screen,entry_screen)
    generate_xml_script_footer()

if __name__ == "__main__":
    # if len(sys.argv) != 2:
    #     print("Usage: python script.py <excel_file_path>")
    # else:
    #     excel_file_path = sys.argv[1]
    excel_file_path = 'Macro_Table.xlsx'
    read_excel_and_generate_xml(excel_file_path)
