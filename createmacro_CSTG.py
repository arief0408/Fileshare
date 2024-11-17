
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
    <vars>
      <create name="$Check_Null$" row="" type="string" value="&apos;&apos;" />
      <create name="$Temp_Rcpt_Number$" row="" type="string" value="&apos;&apos;" />
    </vars>
    '''
    with open(file_path, "a") as file:
        file.write(xml_content)

def generate_xml_script(file_path, item_value, item_value_next, entry_screen,polnum):
    xml_content = f'''<screen name="Screen7_{item_value}" entryscreen="{entry_screen}" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;{polnum}&apos;" row="18" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="50" />   
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="18" scol="44" erow="18" ecol="51" unwrap="false" continuous="false" assigntovar="$Temp_Rcpt_Number$"/>
            <pause value="50" />   
            <input value="&apos;F&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
           
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen8IF_{item_value}"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen8IF_{item_value}" entryscreen="false" exitscreen="true" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <runprogram exe=
            "'C:\\Script\\screenshot.exe C:\\Script\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 2 A SCREENSHOT'"
            param="" wait="true"
            assignexitvalue="" />
            <runprogram exe=
            "'C:\\Script\\screenshot.exe C:\\Script\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 1 A '+$Temp_Rcpt_Number$"
            param="" wait="true"
            assignexitvalue="" />

 
            <input value="&apos;X&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="50" />   
            <runprogram exe=
            "'C:\\Script\\screenshot.exe C:\\Script\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 2 B SCREENSHOT'"
            param="" wait="true"
            assignexitvalue="" />

            <input value="&apos;[pagedn]&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="50" />     
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="21" scol="67" erow="21" ecol="72" unwrap="false" continuous="false" assigntovar="$Check_Null$"/>
            <pause value="50" />     

            <if condition="($Check_Null$ !=&apos;&apos;)">
                <pause value="50" />   
                <runprogram exe=
                "'C:\\Script\\screenshot.exe C:\\Script\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 2 C SCREENSHOT'"
                param="" wait="true"
                assignexitvalue="" />  
            </if>

            <input value="&apos;[pagedn]&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="50" />     
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="21" scol="67" erow="21" ecol="72" unwrap="false" continuous="false" assigntovar="$Check_Null$"/>
            <pause value="50" />     

            <if condition="($Check_Null$ !=&apos;&apos;)">
                <pause value="50" />   
                <runprogram exe=
                "'C:\\Script\\screenshot.exe C:\\Script\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 2 D SCREENSHOT'"
                param="" wait="true"
                assignexitvalue="" />  
            </if>

            <input value="&apos;[pagedn]&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="50" />     
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="21" scol="67" erow="21" ecol="72" unwrap="false" continuous="false" assigntovar="$Check_Null$"/>
            <pause value="50" />     

            <if condition="($Check_Null$ !=&apos;&apos;)">
                <pause value="50" />   
                <runprogram exe=
                "'C:\\Script\\screenshot.exe C:\\Script\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 2 E SCREENSHOT'"
                param="" wait="true"
                assignexitvalue="" />  
            </if>

            <input value="&apos;[pagedn]&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="50" />     
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="21" scol="67" erow="21" ecol="72" unwrap="false" continuous="false" assigntovar="$Check_Null$"/>
            <pause value="50" />     

            <if condition="($Check_Null$ !=&apos;&apos;)">
                <pause value="50" />   
                <runprogram exe=
                "'C:\\Script\\screenshot.exe C:\\Script\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 2 F SCREENSHOT'"
                param="" wait="true"
                assignexitvalue="" />  
            </if>

            <input value="&apos;[pagedn]&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="50" />     
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="21" scol="67" erow="21" ecol="72" unwrap="false" continuous="false" assigntovar="$Check_Null$"/>
            <pause value="50" />     

            <if condition="($Check_Null$ !=&apos;&apos;)">
                <pause value="50" />   
                <runprogram exe=
                "'C:\\Script\\screenshot.exe C:\\Script\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 2 G SCREENSHOT'"
                param="" wait="true"
                assignexitvalue="" />  
            </if>
            <pause value="50" />     
            <input value="&apos;[pf3]&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen9IF_{item_value_next}"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

'''
    with open(file_path, "a") as file:
        file.write(xml_content)


def read_excel_and_generate_xml(file_path, output_path):
    # Read the Excel file
    df = pd.read_excel(file_path)
    entry_screen = 'false'
    exit_screen = 'false'
    generate_xml_script_header(output_path)

    # Iterate over each row
    for index, row in df.iterrows():
        item_value = row['Item']
        polnum = row['Polnum']


        # Check if there's a next row for the item_value_next
        if index + 1 < len(df):
            item_value_next = df.iloc[index + 1]['Item']
            if index == 0:
                entry_screen = 'true'
        else:
            item_value_next = "End"  # Default value if there's no next row
        if index == len(df) - 1:
            exit_screen = 'true'
        
        generate_xml_script(output_path, item_value, item_value_next, entry_screen, polnum)


    generate_xml_script_footer(output_path)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py <excel_file_path>")
    else:
        excel_file_path = sys.argv[1]

        # Define the output file path
        output_file_path = f'CSTG.mac'

        # Generate the XML script
        read_excel_and_generate_xml(excel_file_path, output_file_path)

        print(f"XML script generated and saved to {output_file_path}")
