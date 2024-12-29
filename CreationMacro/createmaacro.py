import pandas as pd

# Load the Excel file into a DataFrame
input_file = 'data.xlsx'  # Replace with your Excel file name
df = pd.read_excel(input_file)

# Template for the .mac content
mac_template = """<HAScript name="create_proposal" description="" timeout="60000" pausetime="300" promptall="true" blockinput="true" author="" creationdate="" supressclearevents="false" usevars="true" ignorepauseforenhancedtn="true" delayifnotenhancedtn="0" ignorepausetimeforenhancedtn="true" continueontimeout="false">

    <vars>
        <create name="$Salutation$" type="string" value="&apos;BAPAK&apos;" />
        <create name="$Married$" type="string" value="&apos;S&apos;" />
        <create name="$Street_Name$" type="string" value="&apos;Jalan Taman Pinus&apos;" />
        <create name="$State$" type="string" value="&apos;DKI Jakarta&apos;" />
        <create name="$Postal_Code$" type="string" value="&apos;13950&apos;" />
        <create name="$Phonenum$" type="string" value="&apos;6289672300432&apos;" />
        <create name="$ID_Num$" type="string" value="&apos;55464827&apos;" />
        <create name="$Contract_Owner$" type="string" value="&apos;55464827&apos;" />
        <!-- <create name="$Billing_Frequency$" type="string" value="&apos;12&apos;" /> -->
        <!-- <create name="$Payment_Method$" type="string" value="&apos;D&apos;" /> -->
        <create name="$Document_ID$" type="string" value="&apos;DEFAULT&apos;" />
        <create name="$Agency$" type="string" value="&apos;00000402&apos;" />
        <!-- <create name="$Agency$" type="string" value="&apos;00001148&apos;" /> -->
        <!-- <create name="$Agency$" type="string" value="&apos;00001418&apos;" /> -->
        <create name="$Prop_Date$" type="string" value="&apos;25/09/2024&apos;" />
        <create name="$Risk_Date$" type="string" value="&apos;27/11/2024&apos;" />
        <create name="$Temp_Rcpt_Number$" type="string" value="&apos;&apos;" />
        <create name="$Existing_Pol_Stmt_1$" type="string" value="&apos;N&apos;" />
        <create name="$Existing_Pol_Stmt_2$" type="string" value="&apos;N&apos;" />
        <create name="$EDD_Details$" row="" type="string" value="&apos;X&apos;" />
        <create name="$PWD_Details$" row="" type="string" value="&apos;Y&apos;" />
        <create name="$Agent_Number$" type="string" value="&apos;00000402&apos;" />
        <!-- <create name="$Agent_Number$" type="string" value="&apos;00001148&apos;" /> -->
        <!-- <create name="$Agent_Number$" type="string" value="&apos;00000401&apos;" /> -->
        <!-- <create name="$Agent_Number$" type="string" value="&apos;00001418&apos;" /> -->
        <create name="$Refferal_Name$" type="string" value="&apos;55452907&apos;" />
        <create name="$Branch$" row="" type="string" value="&apos;BBCAB&apos;" />
        <!-- <create name="$Branch$" row="" type="string" value="&apos;HODM&apos;" /> -->
        <create name="$CRS_Confirmation$" row="" type="string" value="&apos;Y&apos;" />
        <create name="$Currency$" row="" type="string" value="&apos;IDR&apos;" />
        <create name="$Curr_Date$" row="" type="string" value="&apos;26/09/2024&apos;" />
        <create name="$Premi_Cost$" row="" type="string" value="&apos;45000000&apos;" />
        <create name="$Check_Null$" row="" type="string" value="&apos;0&apos;" />
        <create name="$Referral$" row="" type="string" value="&apos;REBS&apos;" />
        <create name="$Pol_Num$" row="" type="string" value="&apos;&apos;" />
        <create name="$Pol_Price$" row="" type="string" value="&apos;&apos;" />
        <create name="$Rand_Num$" row="" type="string" value="&apos;&apos;" />
        <create name="$Email$" row="" type="string" value="&apos;ariefchaerudin@gmail.com&apos;" />
  
  
        <create name="$Bill_Day$" row="" type="string" value="&apos;&apos;" />
        <create name="$Bill_Month$" row="" type="string" value="&apos;&apos;" />
        <create name="$Bill_Year$" row="" type="string" value="&apos;&apos;" />
  
        <create name="$Line_Capture$" row="" type="string" value="&apos;&apos;" />
        <create name="$Commission$" row="" type="string" value="&apos;&apos;" />
        <create name="$E_Sub$" row="" type="string" value="&apos;Y&apos;" />
        <create name="$E_Pol$" row="" type="string" value="&apos;Y&apos;" />
        <create name="$Mail$" row="" type="string" value="&apos;P&apos;" />
  
  
        <create name="$Surname$" type="string" value="&apos;{Surname}&apos;" />
        <create name="$Product$" type="string" value="&apos;{Product}&apos;" />
        <create name="$Payment_Method$" type="string" value="&apos;{Payment_Method}&apos;" />
        <create name="$Prem_Term$" row="" type="string" value="&apos;{Prem_Term}&apos;" />
        <create name="$Term$" row="" type="string" value="&apos;{Term}&apos;" />
        <create name="$Billing_Frequency$" type="string" value="&apos;{Billing_Frequency}&apos;" />
        <create name="$Sum_Assured$" row="" type="string" value="&apos;{Sum_Assured}&apos;" />
        <create name="$Gender$" type="string" value="&apos;{Gender}&apos;" />
        <create name="$Premium$" row="" type="string" value="&apos;{Premium}&apos;" />  
        <create name="$Smoking$" row="" type="string" value="&apos;{Smoking}&apos;" />
        <create name="$Premill$" row="" type="string" value="&apos;{Premill}&apos;" />
        <create name="$Death$" row="" type="string" value="&apos;{Death}&apos;" />
        <create name="$CI$" row="" type="string" value="&apos;{CI}&apos;" />
        <create name="$DOB$" type="string" value="&apos;{DOB}&apos;" />
  

    </vars>


    <screen name="Screen1" entryscreen="true" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
           

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
            <input value="&apos;[enter]&apos;" row="9" col="10" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="100" />
            <input value="&apos;[enter]&apos;" row="5" col="13" movecursor="true" xlatehostkeys="true" encrypted="false" />
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
            <input value="&apos;A&apos;" row="21" col="37" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen4"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    
    <screen name="Screen4" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="$Surname$" row="5" col="17" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Salutation$" row="7" col="17" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Gender$" row="7" col="57" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Married$" row="7" col="74" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Street_Name$" row="8" col="17" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$State$" row="11" col="17" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Postal_Code$" row="12" col="17" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <!-- <input value="$State$" row="12" col="37" movecursor="true" xlatehostkeys="true" encrypted="false" /> -->
            <input value="&apos;RI&apos;" row="13" col="17" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;RI&apos;" row="14" col="17" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Phonenum$" row="14" col="21" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;R&apos;" row="13" col="35" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="4" scol="17" erow="4" ecol="24" unwrap="false" continuous="false" assigntovar="$ID_Num$"/>
            <input value="$ID_Num$" row="13" col="52" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;ADMO&apos;" row="16" col="17" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$DOB$" row="17" col="17" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$State$" row="17" col="43" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;DIJKT&apos;" row="12" col="28" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;RI&apos;" row="15" col="53" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[pf5]&apos;" row="15" col="53" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;X&apos;" row="19" col="36" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;X&apos;" row="19" col="52" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;X&apos;" row="19" col="63" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;X&apos;" row="19" col="77" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;X&apos;" row="20" col="52" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;X&apos;" row="20" col="63" movecursor="true" xlatehostkeys="true" encrypted="false" />
            
            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="1000"/>
            <input value="$Phonenum$" row="10" col="28" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Email$" row="14" col="28" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;RI&apos;" row="10" col="47" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;RI&apos;" row="16" col="68" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;N&apos;" row="16" col="28" movecursor="true" xlatehostkeys="true" encrypted="false" />

            <input value="&apos;[enter]&apos;" row="16" col="28" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="800"/>
            <input value="&apos;H&apos;" row="5" col="42" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;G&apos;" row="6" col="42" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[pf5]&apos;" row="6" col="42" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="400"/>
            <input value="&apos;G&apos;" row="12" col="7" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;F&apos;" row="12" col="20" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;50000000&apos;" row="12" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;50000000&apos;" row="12" col="57" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;G&apos;" row="12" col="42" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;F&apos;" row="12" col="55" movecursor="true" xlatehostkeys="true" encrypted="false" />

            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="500"/>
            <input value="&apos;A&apos;" row="4" col="43" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;A&apos;" row="5" col="43" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;E&apos;" row="6" col="43" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[pf5]&apos;" row="6" col="43" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="400"/>

            <input value="&apos;[enter]&apos;" row="6" col="43" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="800"/>
            <input value="&apos;NOTAX&apos;" row="8" col="4" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;R&apos;" row="8" col="54" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[pf5]&apos;" row="8" col="54" movecursor="true" xlatehostkeys="true" encrypted="false" />
           
            <input value="&apos;[enter]&apos;" row="8" col="54" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="500"/>
            <input value="&apos;N&apos;" row="9" col="48" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="9" col="48" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="500"/>
            <input value="&apos;[enter]&apos;" row="9" col="48" movecursor="true" xlatehostkeys="true" encrypted="false" />
           
            </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen5A"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen5A" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <pause value="100" />
            <input value="&apos;[pf3]&apos;" row="11" col="11" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="100" />
            <input value="&apos;[pf3]&apos;" row="11" col="11" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="100" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen2SS" />
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen2SS" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <pause value="50" />
            <input value="&apos;[enter]&apos;" row="11" col="11" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="50" />
            <input value="&apos;[enter]&apos;" row="8" col="13" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="50" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen3A"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen3A" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <if condition="($Product$ ==&apos;E2R1&apos; || $Product$ ==&apos;E2R3&apos; || $Product$ ==&apos;E2R5&apos;)">
                <input value="&apos;E2R&apos;" row="19" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            </if>
            
            <if condition="($Product$ !=&apos;E2R1&apos; || $Product$ !=&apos;E2R3&apos; || $Product$ !=&apos;E2R5&apos;)">
                <input value="$Product$" row="19" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            </if>
            <pause value="100" />
            <input value="&apos;A&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="500" />
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="1" scol="72" erow="1" ecol="76" unwrap="   false" continuous="false" assigntovar="$Check_Null$"/>
        <if condition="($Check_Null$ ==&apos;S5002&apos;)">
            <!-- <if condition="($condition1$ !=&apos;&apos;)" > -->

            <boxselection type="SELECT" srow="18" scol="44" erow="18" ecol="53" />
            <pause value="500" />
            <input value="&apos;[cut]&apos;" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="500" />
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="1" scol="72" erow="1" ecol="76" unwrap="   false" continuous="false" assigntovar="$Check_Null$"/>

        </if>
        <if condition="($Check_Null$ ==&apos;S5002&apos;)">
            <!-- <if condition="($condition1$ !=&apos;&apos;)" > -->

            <boxselection type="SELECT" srow="18" scol="44" erow="18" ecol="53" />
            <pause value="500" />
            <input value="&apos;[cut]&apos;" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="500" />
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="1" scol="72" erow="1" ecol="76" unwrap="   false" continuous="false" assigntovar="$Check_Null$"/>

        </if>
        <if condition="($Check_Null$ ==&apos;S5002&apos;)">
            <!-- <if condition="($condition1$ !=&apos;&apos;)" > -->

            <boxselection type="SELECT" srow="18" scol="44" erow="18" ecol="53" />
            <pause value="500" />
            <input value="&apos;[cut]&apos;" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="500" />
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="1" scol="72" erow="1" ecol="76" unwrap="   false" continuous="false" assigntovar="$Check_Null$"/>

        </if>
        <pause value="500" />
        <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="1" scol="72" erow="1" ecol="76" unwrap="   false" continuous="false" assigntovar="$Check_Null$"/>
        <if condition="($Check_Null$ ==&apos;S5002&apos;)">
            <!-- <if condition="($condition1$ !=&apos;&apos;)" > -->

            <boxselection type="SELECT" srow="18" scol="44" erow="18" ecol="53" />
            <pause value="500" />
            <input value="&apos;[cut]&apos;" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </if>
            </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen4A"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    
    <screen name="Screen4A" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <pause value="500" />
            <input value="&apos;[pf2]&apos;" row="5" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="500" />            
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="9" scol="36" erow="9" ecol="45" unwrap="false" continuous="false" assigntovar="$Curr_Date$"/>
            <pause value="200" />
            <input value="&apos;[pf3]&apos;" row="5" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="200" />
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="3" scol="22" erow="3" ecol="29" unwrap="false" continuous="false" assigntovar="$Temp_Rcpt_Number$"/>
            <pause value="200" />
            <input value="$ID_Num$" row="5" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Billing_Frequency$" row="7" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Payment_Method$" row="8" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Document_ID$" row="13" col="15" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Agency$" row="16" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <!-- <input value="&apos;Y&apos;" row="07" col="52" movecursor="true" xlatehostkeys="true" encrypted="false" /> -->
            <input value="$Curr_Date$" row="6" col="71" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <!-- <input value="$Risk_Date$" row="6" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" /> -->
            <!-- <input value="$Curr_Date$" row="6" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" /> -->
            <input value="&apos;[pf5]&apos;" row="5" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[pf5]&apos;" row="5" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="200" />
            <input value="$Temp_Rcpt_Number$" row="13" col="61" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="200" />
            <input value="$Existing_Pol_Stmt_1$" row="18" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="200" />
            <input value="$E_Pol$" row="19" col="48" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="200" />
            <input value="$E_Sub$" row="19" col="56" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="200" />
            <input value="$Mail$" row="23" col="48" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="200" />
            <input value="$Existing_Pol_Stmt_2$" row="18" col="46" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="200" />
            <input value="$EDD_Details$" row="18" col="63" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$PWD_Details$" row="18" col="77" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$CRS_Confirmation$" row="22" col="56" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Currency$" row="10" col="22" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Currency$" row="10" col="52" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="13" col="61" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="500"/>
            <if condition="($Payment_Method$ ==&apos;K&apos;) || ($Payment_Method$ ==&apos;D&apos;)">
                <input value="&apos;[pf4]&apos;" row="15" col="28" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="500"/>
                <input value="&apos;[pf4]&apos;" row="17" col="25" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="250"/>
                <input value="&apos;[pf7]&apos;" row="17" col="25" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="500"/>
                <input value="&apos;BC&apos;" row="7" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <input value="&apos;BCA&apos;" row="9" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <!-- <input value="&apos;12451234521&apos;" row="13" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" /> -->
    
                <!-- <input value="&apos;[pf2]&apos;" row="13" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="11" scol="42" erow="11" ecol="43" unwrap="   false" continuous="false" assigntovar="$Rand_Num$"/>
                <input value="&apos;[enter]&apos;" row="13" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <input value="$Rand_Num$" row="13" col="25" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="100" /> -->
    
                <!-- <input value="&apos;[pf2]&apos;" row="13" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="11" scol="42" erow="11" ecol="43" unwrap="   false" continuous="false" assigntovar="$Rand_Num$"/>
                <input value="&apos;[enter]&apos;" row="13" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" /> -->
                <input value="$Temp_Rcpt_Number$" row="13" col="30" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="100" />
    
                <!-- <input value="&apos;[pf2]&apos;" row="13" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="11" scol="42" erow="11" ecol="43" unwrap="   false" continuous="false" assigntovar="$Rand_Num$"/>
                <input value="&apos;[enter]&apos;" row="13" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <input value="$Rand_Num$" row="13" col="31" movecursor="true" xlatehostkeys="true" encrypted="false" /> -->
                <input value="$Temp_Rcpt_Number$" row="13" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="100" />
    
                <!-- <input value="&apos;[pf2]&apos;" row="13" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="11" scol="42" erow="11" ecol="43" unwrap="   false" continuous="false" assigntovar="$Rand_Num$"/>
                <input value="&apos;[enter]&apos;" row="13" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <input value="$Rand_Num$" row="13" col="34" movecursor="true" xlatehostkeys="true" encrypted="false" /> -->
                <pause value="100" />
    
                <input value="&apos;BCA Dummy&apos;" row="14" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <input value="&apos;[pf5]&apos;" row="14" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <input value="&apos;[enter]&apos;" row="15" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="100" />
                <input value="&apos;[enter]&apos;" row="15" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="100" />
                <input value="&apos;[enter]&apos;" row="15" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="100" />
                <input value="&apos;[enter]&apos;" row="15" col="24" movecursor="true" xlatehostkeys="true" encrypted="false" />

            </if>
            <input value="$Referral$" row="10" col="19" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Agent_Number$" row="12" col="19" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Refferal_Name$" row="13" col="19" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Branch$" row="14" col="19" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="18" col="63" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="500"/>
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen5B"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    
    <screen name="Screen5B" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;Y&apos;" row="10" col="8" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="13" col="61" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen6A"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen6A" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;X&apos;" row="19" col="41" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;Y&apos;" row="19" col="72" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="19" col="72" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen7A"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen7A" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="$Curr_Date$" row="8" col="29" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;L&apos;" row="10" col="29" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;N&apos;" row="12" col="29" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;G&apos;" row="15" col="35" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="13" col="61" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="13" col="61" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="13" col="61" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen8A"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen8A" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;[pf5]&apos;" row="10" col="27" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;NIL&apos;" row="14" col="27" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="$Smoking$" row="16" col="27" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;X&apos;" row="21" col="27" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="13" col="61" movecursor="true" xlatehostkeys="true" encrypted="false" />

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen9A"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen9A" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;170&apos;" row="7" col="11" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;77&apos;" row="7" col="27" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;OW&apos;" row="7" col="77" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;BC&apos;" row="8" col="15" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="13" col="61" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="13" col="61" movecursor="true" xlatehostkeys="true" encrypted="false" />

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen10"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen10" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <if condition="($Product$ !=&apos;E2R3&apos; &amp;&amp; $Product$ !=&apos;E2R5&apos;)">
                <input value="&apos;X&apos;" row="8" col="9" movecursor="true" xlatehostkeys="true" encrypted="false" />
            </if>
            <if condition="($Product$ ==&apos;E2R3&apos;)">
                <input value="&apos;X&apos;" row="9" col="9" movecursor="true" xlatehostkeys="true" encrypted="false" />
            </if>
            <if condition="($Product$ ==&apos;E2R5&apos;)">
                <input value="&apos;X&apos;" row="10" col="9" movecursor="true" xlatehostkeys="true" encrypted="false" />
            </if>
            <input value="&apos;[enter]&apos;" row="8" col="9" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen11"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen11" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>

            <if condition="($Product$ ==&apos;E2R1&apos;)">
                    <input value="$Sum_Assured$" row="8" col="17" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="$Premium$" row="17" col="58" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <!-- <input value="&apos;10&apos;" row="10" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" /> -->
                    <input value="&apos;[pf5]&apos;" row="10" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <pause value="200" />
            </if> 

            

            <if condition="($Product$ ==&apos;E2R3&apos;)">
                    <input value="$Sum_Assured$" row="8" col="17" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="$Premium$" row="17" col="58" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="$Term$" row="9" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="$Prem_Term$" row="10" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="&apos;[pf5]&apos;" row="10" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <pause value="200" />
            </if> 

            <if condition="($Product$ ==&apos;E2R5&apos;)">
                    <input value="$Sum_Assured$" row="8" col="17" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="$Premium$" row="17" col="58" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="$Term$" row="9" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="$Prem_Term$" row="10" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="&apos;[pf5]&apos;" row="10" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="&apos;[enter]&apos;" row="10" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <pause value="200" />
            </if> 

            <if condition="($Premill$ != &apos;&apos; || $Death$ != &apos;&apos; || $CI$ != &apos;&apos;)">
                    <input value="&apos;x&apos;" row="17" col="32" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="&apos;[enter]&apos;" row="17" col="32" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <pause value="400" />
                    <input value="$Death$" row="11" col="25" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="$CI$" row="11" col="34" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="$Premill$" row="11" col="59" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="&apos;oc&apos;" row="15" col="2" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="&apos;[pf5]&apos;" row="15" col="2" movecursor="true" xlatehostkeys="true" encrypted="false" />
                        <runprogram exe=
                        "'C:\\\\Script\\\\screenshot.exe C:\\\\Script\\\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 2 A SCREENSHOT'"
                        param="" wait="true"
                        assignexitvalue="" />
                    <pause value="200" />
                    <input value="&apos;[enter]&apos;" row="15" col="2" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <pause value="400" />
                    <boxselection type="SELECT" srow="17" scol="58" erow="17" ecol="79" />
                    <pause value="500" />
                    <input value="&apos;[cut]&apos;" row="0" col="0" movecursor="true" xlatehostkeys="true" encrypted="false" />
                    <input value="&apos;[pf5]&apos;" row="10" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
                        <runprogram exe=
                        "'C:\\\\Script\\\\screenshot.exe C:\\\\Script\\\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 2 B SCREENSHOT'"
                        param="" wait="true"
                        assignexitvalue="" />
                    <pause value="400" />
            </if> 
            <input value="&apos;[enter]&apos;" row="10" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />


        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen12"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen12" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;[pf5]&apos;" row="9" col="3" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="100" />
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="5" scol="20" erow="5" ecol="27" unwrap="false" continuous="false" assigntovar="$Pol_Num$"/>
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="10" scol="27" erow="10" ecol="37" unwrap="false" continuous="false" assigntovar="$Pol_Price$"/>
            <pause value="700" />
            <input value="&apos;[enter]&apos;" row="9" col="3" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="100" />
            <if condition="($Payment_Method$ ==&apos;C&apos;) &amp;&amp; $Product$ !=&apos;E2R1&apos; &amp;&amp; $Product$ !=&apos;E2R3&apos; &amp;&amp; $Product$ !=&apos;E2R5&apos;">
                <input value="&apos;1&apos;" row="9" col="3" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <input value="&apos;[enter]&apos;" row="9" col="3" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="500" />
                <input value="&apos;x&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <input value="&apos;[enter]&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="200" />
                <input value="&apos;R&apos;" row="12" col="15" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <input value="&apos;[enter]&apos;" row="15" col="15" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="100" />
                <input value="&apos;[enter]&apos;" row="15" col="15" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="100" />
                <input value="&apos;[enter]&apos;" row="15" col="15" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="100" />
                <input value="&apos;[pf5]&apos;" row="15" col="15" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="500" />
                <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="5" scol="20" erow="5" ecol="27" unwrap="false" continuous="false" assigntovar="$Pol_Num$"/>
                <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="10" scol="27" erow="10" ecol="37" unwrap="false" continuous="false" assigntovar="$Pol_Price$"/>
                <pause value="700" />
                <input value="&apos;[enter]&apos;" row="15" col="15" movecursor="true" xlatehostkeys="true" encrypted="false" />
                <pause value="700" />

    

            </if>
            <pause value="500" />
            <input value="&apos;[enter]&apos;" row="15" col="15" movecursor="true" xlatehostkeys="true" encrypted="false" />

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen2TA"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen2TA" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;[pf3]&apos;" row="10" col="10" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="400" />
            <input value="&apos;[pf3]&apos;" row="10" col="10" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="400" />
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
            <input value="&apos;[enter]&apos;" row="12" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
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

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen4AB"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    <screen name="Screen4AB" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;[enter]&apos;" row="11" col="10" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <input value="&apos;[enter]&apos;" row="8" col="13" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="500" />
            <input value="$Temp_Rcpt_Number$" row="18" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;C&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen5"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen5" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;[pf5]&apos;" row="11" col="10" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <input value="&apos;[enter]&apos;" row="8" col="13" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen6"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>
    
    
    <screen name="Screen6" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;[pf3]&apos;" row="11" col="10" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <input value="&apos;[enter]&apos;" row="8" col="40" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="500" />
            <input value="$Temp_Rcpt_Number$" row="17" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;A&apos;" row="19" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="19" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <input value="&apos;[pf5]&apos;" row="19" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <input value="&apos;[enter]&apos;" row="19" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <input value="&apos;[enter]&apos;" row="19" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />

        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen7"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen7" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;[pf3]&apos;" row="11" col="10" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <input value="&apos;[enter]&apos;" row="8" col="13" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="1200" />
            <input value="$Temp_Rcpt_Number$" row="18" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;F&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
           
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen8IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen8IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>

            <pause value="500" />
            <runprogram exe=
            "'C:\\\\Script\\\\screenshot.exe C:\\\\Script\\\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 2 C SCREENSHOT'"
            param="" wait="true"
            assignexitvalue="" />
            <runprogram exe=
            "'C:\\\\Script\\\\screenshot.exe C:\\\\Script\\\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 1 A '+$Temp_Rcpt_Number$"
            param="" wait="true"
            assignexitvalue="" />
             <pause value="500" />

            <input value="&apos;X&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="15" scol="19" erow="15" ecol="20" unwrap="false" continuous="false" assigntovar="$Bill_Day$"/>
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="15" scol="22" erow="15" ecol="23" unwrap="false" continuous="false" assigntovar="$Bill_Month$"/>
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="15" scol="25" erow="15" ecol="28" unwrap="false" continuous="false" assigntovar="$Bill_Year$"/>
            <extract name="&apos;Extract&apos;" planetype="TEXT_PLANE" srow="15" scol="35" erow="15" ecol="44" unwrap="false" continuous="false" assigntovar="$Pol_Price$"/>
            <input value="&apos;X&apos;" row="20" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="800" />
            <input value="&apos;1&apos;" row="11" col="2" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="9" col="3" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <runprogram exe=
            "'C:\\\\Script\\\\screenshot.exe C:\\\\Script\\\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 2 D SCREENSHOT'"
            param="" wait="true"
            assignexitvalue="" />
            <pause value="300" />
            <input value="&apos;X&apos;" row="20" col="50" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <input value="&apos;[enter]&apos;" row="20" col="5" movecursor="true" xlatehostkeys="true" encrypted="false" />
            <pause value="300" />
            <runprogram exe=
            "'C:\\\\Script\\\\screenshot.exe C:\\\\Script\\\\test_scen.xlsx '+$Temp_Rcpt_Number$+' 2 D SCREENSHOT'"
            param="" wait="true"
            assignexitvalue="" />
            <pause value="100" />
            <input value="&apos;[pf3]&apos;" row="21" col="44" movecursor="true" xlatehostkeys="true" encrypted="false" />
        </actions>
        <nextscreens timeout="0" >
            <nextscreen name="Screen9IF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    <screen name="Screen9IF" entryscreen="false" exitscreen="false" transient="false">
        <description >
            <oia status="NOTINHIBITED" optional="false" invertmatch="false" />
        </description>
        <actions>
            <input value="&apos;[pf3]&apos;" row="21" col="39" movecursor="true" xlatehostkeys="true" encrypted="false" /> 
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
            <nextscreen name="Screen10IFF"/>
        </nextscreens>
        <recolimit value="10000" />
    </screen>

    
</HAScript>

"""

# Generate .mac files for each row in the DataFrame
for index, row in df.iterrows():
    formatted_content = mac_template.format(
        Surname=row.get('Surname', ''),
        Product=row.get('Product', ''),
        Payment_Method=row.get('Payment_Method', ''),
        Prem_Term=row.get('Prem_Term', ''),
        Term=row.get('Term', ''),
        Billing_Frequency=row.get('Billing_Frequency', ''),
        Sum_Assured=row.get('Sum_Assured', ''),
        Gender=row.get('Gender', ''),
        Premium=row.get('Premium', ''),
        Smoking=row.get('Smoking', ''),
        Premill=row.get('Permill', ''),
        Death=row.get('Death', ''),
        CI=row.get('CI', ''),
        TC_Next=row.get('TC_Next', ''),
        DOB=row.get('DOB', '')
        
    )
    
    # Save to a .mac file
    with open(f'bospo_{index + 1}.mac', 'w', encoding='utf-8') as file:
        file.write(formatted_content)

print("MAC files generated successfully!")
