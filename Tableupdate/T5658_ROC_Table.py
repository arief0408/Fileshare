import pandas as pd
import numpy as np

# Load the Excel file into a DataFrame
excel_file = 'Macro_Table.xlsx'  # Replace with your actual file name
excel_data = pd.ExcelFile(excel_file)

# Template for the .mac content
mac_template = """<HAScript name="create_proposal" description="" timeout="60000" pausetime="300" promptall="true" blockinput="true" author="" creationdate="" supressclearevents="false" usevars="true" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true" continueontimeout="false">

        <vars>
        
    <create name="$Check$" type="string" value="&apos;&apos;" />
  

    </vars>


    <screen name="Screen1" entryscreen="true" exitscreen="false" transient="false">
        <description >
        <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
    </description>
    <actions>
        <input value="&apos;T5658&apos;" row="17" col="35" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;{item_value}&apos;" row="18" col="35" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;E&apos;" row="22" col="35" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;[enter]&apos;" row="22" col="35" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <pause value="200"/>
    </actions>
    <nextscreens timeout="0" >
        <nextscreen name="Input_Work_Unit_{TC_Next}" />
    </nextscreens>
    <recolimit value="10000" />
</screen>



<screen name="Input_Work_Unit_{TC_Next}" entryscreen="false" exitscreen="false" transient="false">
    <description >
        <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
    </description>
    <actions>
        <input value="&apos;L01624&apos;" row="11" col="41" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;[enter]&apos;" row="16" col="60" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <pause value="200"/>
        <input value="&apos;X&apos;" row="16" col="60" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <input value="&apos;[enter]&apos;" row="16" col="60" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <pause value="200"/>
        <input value="&apos;[enter]&apos;" row="16" col="60" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <pause value="300"/>
        <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="10" scol="36" erow="10" ecol="45" unwrap="false" continuous="false" assigntovar="$Check$"/>
        
        <if condition="($Check$ !=&apos; &apos;)">
            <input value="&apos;1&apos;" row="10" col="23" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </if>
        <if condition="($Check$ ==&apos; &apos;)">
            <input value="&apos;1&apos;" row="11" col="23" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </if>
        <input value="&apos;[enter]&apos;" row="11" col="23" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <pause value="100"/>
    </actions>
    <nextscreens timeout="0" >
        <nextscreen name="Table_Item_Maintenance_{TC_Next}" />
    </nextscreens>
    <recolimit value="10000" />
</screen>

<screen name="Table_Item_Maintenance_{TC_Next}" entryscreen="false" exitscreen="true" transient="false">
    <description >
        <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
    </description>
    <actions>
        {output_string}
        <input value="&apos;[enter]&apos;" row="16" col="31" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <pause value="400"/>
        <input value="&apos;[enter]&apos;" row="16" col="31" movecursor="true" xlatehostkeys="true" encrypted="false" />
        <playmacro name="bospo_{sheet_index_next}.mac" startScreen="*DEFAULT*" transferVars="Transfer" />

    </actions>
    <nextscreens timeout="0" >
    </nextscreens>
    <recolimit value="10000" />
</screen>


    
</HAScript>

"""
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

