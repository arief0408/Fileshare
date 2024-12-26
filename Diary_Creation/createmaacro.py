import pandas as pd

# Load the Excel file into a DataFrame
input_file = 'data.xlsx'  # Replace with your Excel file name
df = pd.read_excel(input_file)

# Template for the .mac content
mac_template = """<HAScript name="create_proposal" description="" timeout="60000" pausetime="300" promptall="true" blockinput="true" author="" creationdate="" supressclearevents="false" usevars="true" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true" continueontimeout="false">

    <vars>
      <create name="$Surname$" type="string" value="&apos;Aref H11 A13&apos;" />
      <create name="$Product$" type="string" value="&apos;H13&apos;" />
      <create name="$Salutation$" type="string" value="&apos;BAPAK&apos;" />
      <create name="$Gender$" type="string" value="&apos;M&apos;" />
      <create name="$Married$" type="string" value="&apos;S&apos;" />
      <create name="$Street_Name$" type="string" value="&apos;Jalan Taman Pinus&apos;" />
      <create name="$State$" type="string" value="&apos;DKI Jakarta&apos;" />
      <create name="$Postal_Code$" type="string" value="&apos;13950&apos;" />
      <create name="$Phonenum$" type="string" value="&apos;6289672300432&apos;" />
      <create name="$ID_Num$" type="string" value="&apos;55464827&apos;" />
      <create name="$DOB$" type="string" value="&apos;01/01/1998&apos;" />
      <create name="$Contract_Owner$" type="string" value="&apos;55464827&apos;" />
      <create name="$Billing_Frequency$" type="string" value="&apos;12&apos;" />
      <!-- <create name="$Billing_Frequency$" type="string" value="&apos;12&apos;" /> -->
      <create name="$Payment_Method$" type="string" value="&apos;C&apos;" />
      <!-- <create name="$Payment_Method$" type="string" value="&apos;D&apos;" /> -->
      <create name="$Document_ID$" type="string" value="&apos;DEFAULT&apos;" />
      <create name="$Agency$" type="string" value="&apos;00001148&apos;" />
      <create name="$Prop_Date$" type="string" value="&apos;25/09/2024&apos;" />
      <create name="$Risk_Date$" type="string" value="&apos;27/11/2024&apos;" />
      <create name="$Temp_Rcpt_Number$" type="string" value="&apos;{Policy}&apos;" />
      <create name="$Existing_Pol_Stmt_1$" type="string" value="&apos;N&apos;" />
      <create name="$Existing_Pol_Stmt_2$" type="string" value="&apos;N&apos;" />
      <create name="$EDD_Details$" row="" type="string" value="&apos;X&apos;" />
      <create name="$PWD_Details$" row="" type="string" value="&apos;Y&apos;" />
      <create name="$Agent_Number$" type="string" value="&apos;00001148&apos;" />
      <!-- <create name="$Agent_Number$" type="string" value="&apos;00001148&apos;" /> -->
      <create name="$Refferal_Name$" type="string" value="&apos;55452907&apos;" />
      <create name="$Branch$" row="" type="string" value="&apos;BBCAB&apos;" />
      <create name="$CRS_Confirmation$" row="" type="string" value="&apos;Y&apos;" />
      <create name="$Currency$" row="" type="string" value="&apos;IDR&apos;" />
      <create name="$Curr_Date$" row="" type="string" value="&apos;26/09/2024&apos;" />
      <create name="$Premi_Cost$" row="" type="string" value="&apos;45000000&apos;" />
      <create name="$Check_Null$" row="" type="string" value="&apos;0&apos;" />
      <create name="$Referral$" row="" type="string" value="&apos;REBS&apos;" />
      <create name="$Term$" row="" type="string" value="&apos;L2J10&apos;" />
      <create name="$Pol_Num$" row="" type="string" value="&apos;&apos;" />
      <create name="$Pol_Price$" row="" type="string" value="&apos;{Premium}&apos;" />
      <create name="$Rand_Num$" row="" type="string" value="&apos;&apos;" />


      <create name="$Bill_Day$" row="" type="string" value="&apos;{Date}&apos;" />
      <create name="$Bill_Month$" row="" type="string" value="&apos;{Month}&apos;" />
      <create name="$Bill_Year$" row="" type="string" value="&apos;{Year}&apos;" />

      <create name="$Line_Capture$" row="" type="string" value="&apos;&apos;" />
      <create name="$Commission$" row="" type="string" value="&apos;&apos;" />

      <create name="$Product_Pol$" row="" type="string" value="&apos;&apos;" />
      <create name="$Check_Hold$" row="" type="string" value="&apos;&apos;" />



    </vars>


    <screen name="Screen1" entryscreen="true" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="$Temp_Rcpt_Number$" row="18" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;F&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />


        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen2" />
        </nextscreens>
        <recolimit value="10000" />
    </screen>


    <screen name="Screen2" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="2" scol="24" erow="2" ecol="26" unwrap="false" continuous="false" assigntovar="$Product_Pol$"/>
            <input value="&apos;[pf3]&apos;" row="9" col="10" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <input value="&apos;[pf3]&apos;" row="5" col="13" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen3"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen3" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;[pf3]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen3DRY"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen3DRY" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;[enter]&apos;" row="10" col="10" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen2A"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    
    <screen name="Screen2A" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;[enter]&apos;" row="9" col="13" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen2B"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen2B" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;01&apos;" row="15" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;E&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen2C"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen2C" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;1&apos;" row="5" col="9" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Temp_Rcpt_Number$" row="12" col="27" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Pol_Price$" row="5" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Pol_Price$" row="12" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <input value="&apos;[pf5]&apos;" row="12" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <if condition="($Pol_Price$ !=&apos;&apos;)">
                <input value="&apos;[enter]&apos;" row="12" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            </if>
            <if condition="($Pol_Price$ ==&apos;&apos;)">
                <input value="&apos;[pf3]&apos;" row="12" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            </if>
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen3AB"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen3AB" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <pause value="300" />
            <input value="&apos;[pf3]&apos;" row="15" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <input value="&apos;[pf3]&apos;" row="15" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <input value="&apos;[pf3]&apos;" row="15" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen13IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    
    <screen name="Screen13IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>

            <input value="&apos;chgcurlib arifch&apos;" row="18" col="07" movecursor="true" xlatehostkeys="true" encrypted="false" />           
            <input value="&apos;[enter]&apos;" row="18" col="07" movecursor="true" xlatehostkeys="true" encrypted="false" />        
            <pause value="100" />   
            <input value="&apos;upddta arifch/busdpf&apos;" row="18" col="07" movecursor="true" xlatehostkeys="true" encrypted="false" />           
            <input value="&apos;[enter]&apos;" row="18" col="07" movecursor="true" xlatehostkeys="true" encrypted="false" />    
            <pause value="400" />   
            <input value="&apos;[pagedn]&apos;" row="04" col="16" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="400" />       
            <input value="$Bill_Year$" row="06" col="16" movecursor="true" xlatehostkeys="true" encrypted="false" />    
            <input value="$Bill_Month$" row="06" col="20" movecursor="true" xlatehostkeys="true" encrypted="false" />    
            <input value="$Bill_Day$" row="06" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />    
            <input value="&apos;[pageup]&apos;" row="04" col="16" movecursor="true" xlatehostkeys="true" encrypted="false" />    
            <input value="&apos;[enter]&apos;" row="06" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />    
            <pause value="100" />   
            <input value="&apos;[pf3]&apos;" row="18" col="07" movecursor="true" xlatehostkeys="true" encrypted="false" />    
            <pause value="100" />   
            <input value="&apos;[enter]&apos;" row="18" col="07" movecursor="true" xlatehostkeys="true" encrypted="false" />    
            <pause value="100" />
            <input value="&apos;[enter]&apos;" row="18" col="07" movecursor="true" xlatehostkeys="true" encrypted="false" />    
            <pause value="1000" />   

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen14IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    
    <screen name="Screen14IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>

            <input value="&apos;d&apos;" row="18" col="07" movecursor="true" xlatehostkeys="true" encrypted="false" />           
            <input value="&apos;[enter]&apos;" row="18" col="07" movecursor="true" xlatehostkeys="true" encrypted="false" />           

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen15IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    
    
    <screen name="Screen15IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>

            <input value="&apos;[enter]&apos;" row="16" col="10" movecursor="true" xlatehostkeys="true" encrypted="false" />

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen23IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen23IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
                <input value="&apos;[enter]&apos;" row="9" col="13" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen23GF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen23GF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
                <input value="&apos;CH&apos;" row="14" col="41" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <input value="&apos;A&apos;" row="18" col="41" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <input value="$Temp_Rcpt_Number$" row="15" col="41" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <input value="&apos;[enter]&apos;" row="18" col="41" movecursor="true" xlatehostkeys="true" encrypted="false" />

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen23GF1"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen23GF1" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
                <input value="&apos;6&apos;" row="13" col="3" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <input value="&apos;[enter]&apos;" row="13" col="3" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="200" />   
                <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="24" scol="02" erow="24" ecol="07" unwrap="false" continuous="false" assigntovar="$Check_Hold$"/>
                <if condition="($Check_Hold$ ==&apos;Entity&apos;)">
                    <input value="&apos;[pf3]&apos;" row="13" col="3" movecursor="true" xlatehostkeys="true" encrypted="false" />
                </if>
                <if condition="($Check_Hold$ !=&apos;Entity&apos;)">
                    <input value="&apos;[enter]&apos;" row="13" col="3" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    
                </if>
                <pause value="500" />   
                <input value="&apos;[pf3]&apos;" row="13" col="3" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="500" />   

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen23IFF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen23IFF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
                <input value="&apos;[enter]&apos;" row="9" col="40" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen24IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen24IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
                <input value="&apos;[enter]&apos;" row="16" col="41" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen25IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen25IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
                <input value="&apos;CH&apos;" row="9" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Temp_Rcpt_Number$" row="10" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Temp_Rcpt_Number$" row="10" col="50" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Bill_Day$" row="11" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Bill_Month$" row="11" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Bill_Year$" row="11" col="26" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;N&apos;" row="12" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[pf5]&apos;" row="12" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <runprogram exe=
            "'C:\\\\Script\\\\screenshot.exe C:\\\\Script\\\\test_scen.xlsx '+$Temp_Rcpt_Number$+' {Row} A SCREENSHOT'"
            param="" wait="true"
            assignexitvalue="" />
            <pause value="200" />
            <input value="&apos;[enter]&apos;" row="12" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen26IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen26IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
                <input value="&apos;[pf3]&apos;" row="9" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen27IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen27IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
                <input value="&apos;[pf3]&apos;" row="9" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen28IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen28IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
                <input value="&apos;[enter]&apos;" row="11" col="10" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen29IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen29IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
                <input value="&apos;[enter]&apos;" row="8" col="13" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen30IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen30IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
                <input value="$Temp_Rcpt_Number$" row="18" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;F&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
    </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen31IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen31IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
                <pause value="500" />
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="15" scol="19" erow="15" ecol="20" unwrap="false" continuous="false" assigntovar="$Bill_Day$"/>
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="15" scol="22" erow="15" ecol="23" unwrap="false" continuous="false" assigntovar="$Bill_Month$"/>
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="15" scol="25" erow="15" ecol="28" unwrap="false" continuous="false" assigntovar="$Bill_Year$"/>
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="15" scol="35" erow="15" ecol="44" unwrap="false" continuous="false" assigntovar="$Pol_Price$"/>
            <runprogram exe=
            "'C:\\\\Script\\\\screenshot.exe C:\\\\Script\\\\test_scen.xlsx '+$Temp_Rcpt_Number$+' {Row} B SCREENSHOT'"
            param="" wait="true"
            assignexitvalue="" />
            <input value="&apos;X&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <input value="&apos;X&apos;" row="9" col="3" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="9" col="3" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <runprogram exe=
            "'C:\\\\Script\\\\screenshot.exe C:\\\\Script\\\\test_scen.xlsx '+$Temp_Rcpt_Number$+' {Row} C SCREENSHOT'"
            param="" wait="true"
            assignexitvalue="" />
            <!-- <runprogram exe=
            "'C:\\\\Script\\\\screenshot.exe C:\\\\Script\\\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 1 D YEAR 2'"
            param="" wait="true"
            assignexitvalue="" /> -->
            <!-- <if condition="($Product_Pol$ !=&apos;T2U&apos;)">

            <input value="&apos;[pagedn]&apos;" row="9" col="3" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <runprogram exe=
            "'C:\\\\Script\\\\screenshot.exe C:\\\\Script\\\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 2 E SCREENSHOT'"
            param="" wait="true"
            assignexitvalue="" />
            </if> -->
            <!-- <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" /> -->
           
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen10IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>


    <screen name="Screen10IF" entryscreen="false" exitscreen="true" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>

            <input value="&apos;[pf3]&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />          
            <playmacro name="bospo_{TC_Next}.mac" startScreen="*DEFAULT*" transferVars="Transfer" />

        </actions>
        <nextscreens timeout="0" >
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    

   
    
</HAScript>
 
"""

# Generate .mac files for each row in the DataFrame
for index, row in df.iterrows():
    formatted_content = mac_template.format(
        Policy=row.get('Policy', ''),
        Premium=row.get('Premium', ''),
        Bill_Freq=row.get('Bill_Freq', ''),
        Date=row.get('Date', ''),
        Month=row.get('Month', ''),
        Year=row.get('Year', ''),
        Row=row.get('Row', ''),
        TC_Next=row.get('TC_Next', '')
        
    )
    
    # Save to a .mac file
    with open(f'bospo_{index + 1}.mac', 'w', encoding='utf-8') as file:
        file.write(formatted_content)

print("MAC files generated successfully!")
