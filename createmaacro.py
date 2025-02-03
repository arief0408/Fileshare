import os
import pandas as pd
import numpy as np

# Load the Excel file into a DataFrame
input_file = 'data_cstg_prem_.xlsx'  # Replace with your Excel file name
df = pd.read_excel(input_file)

# Read the .txt file containing the template
with open('mac_template_cstg_prem_.txt', 'r', encoding='utf-8') as template_file:
    mac_template = template_file.read()

# Replace NaN values in the DataFrame with empty strings
df = pd.read_excel(input_file).replace(np.nan, '', regex=True)

# Generate .mac files for each row in the DataFrame
for index, row in df.iterrows():
    # Generate the full path for the .mac file
    formatted_Policy_1 = str(row.get('Policy_1', '')).zfill(8)
    formatted_Premium_1 = str(row.get('Premium_1', '')).zfill(8)
    formatted_Policy_2 = str(row.get('Policy_2', '')).zfill(8)
    formatted_Premium_2 = str(row.get('Premium_2', '')).zfill(8)
    formatted_Policy_3 = str(row.get('Policy_3', '')).zfill(8)
    formatted_Premium_3 = str(row.get('Premium_3', '')).zfill(8)
    formatted_Policy_4 = str(row.get('Policy_4', '')).zfill(8)
    formatted_Premium_4 = str(row.get('Premium_4', '')).zfill(8)
    formatted_Policy_5 = str(row.get('Policy_5', '')).zfill(8)
    formatted_Premium_5 = str(row.get('Premium_5', '')).zfill(8)
    formatted_Policy_6 = str(row.get('Policy_6', '')).zfill(8)
    formatted_Premium_6 = str(row.get('Premium_6', '')).zfill(8)
    formatted_Policy_7 = str(row.get('Policy_7', '')).zfill(8)
    formatted_Premium_7 = str(row.get('Premium_7', '')).zfill(8)
    formatted_Policy_8 = str(row.get('Policy_8', '')).zfill(8)
    formatted_Premium_8 = str(row.get('Premium_8', '')).zfill(8)
    formatted_Policy_9 = str(row.get('Policy_9', '')).zfill(8)
    formatted_Premium_9 = str(row.get('Premium_9', '')).zfill(8)
    formatted_Policy_10 = str(row.get('Policy_10', '')).zfill(8)
    formatted_Premium_10 = str(row.get('Premium_10', '')).zfill(8)

    # Get the IBM_PATH from the row
    ibm_path = row.get('IBM_Path', '')

    # Ensure the IBM_PATH exists and is valid
    if not os.path.exists(ibm_path):
        print(f"Error: IBM_PATH '{ibm_path}' does not exist.")
        continue  # Skip this iteration if the path is invalid

    # Format the content by replacing placeholders with values from the DataFrame
    formatted_content = mac_template.format(
        Bank_Code=row.get('Bank_Code', ''),
        Action=row.get('Action', ''),
        IBM_Path=ibm_path,
        Policy_1=formatted_Policy_1,
        Premium_1=formatted_Premium_1,
        Policy_2=formatted_Policy_2,
        Premium_2=formatted_Premium_2,
        Policy_3=formatted_Policy_3,
        Premium_3=formatted_Premium_3,
        Policy_4=formatted_Policy_4,
        Premium_4=formatted_Premium_4,
        Policy_5=formatted_Policy_5,
        Premium_5=formatted_Premium_5,
        Policy_6=formatted_Policy_6,
        Premium_6=formatted_Premium_6,
        Policy_7=formatted_Policy_7,
        Premium_7=formatted_Premium_7,
        Policy_8=formatted_Policy_8,
        Premium_8=formatted_Premium_8,
        Policy_9=formatted_Policy_9,
        Premium_9=formatted_Premium_9,
        Policy_10=formatted_Policy_10,
        Premium_10=formatted_Premium_10,
        
          # Set IBM_Path to the full path of the .mac file

        TC_Next=index + 2,
        cell=index + 2
    )

    # Create the file path for the .mac file
    mac_filename = f'cstg_prem_{index + 1}.mac'  # The file name (e.g., bospo_1.mac)
    mac_file_path = os.path.join(ibm_path, mac_filename)  # Combine IBM_PATH with the filename

    # Save the formatted content to the .mac file in the specified IBM_PATH
    with open(mac_file_path, 'w', encoding='utf-8') as file:
        file.write(formatted_content)

print("MAC files generated successfully!")
