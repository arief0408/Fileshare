import os
import pandas as pd
import numpy as np

# Load the Excel file into a DataFrame
input_file = 'data_diary.xlsx'  # Replace with your Excel file name
df = pd.read_excel(input_file)

# Read the .txt file containing the template
with open('mac_template_diary.txt', 'r', encoding='utf-8') as template_file:
    mac_template = template_file.read()

# Replace NaN values in the DataFrame with empty strings
df = df.replace(np.nan, '', regex=True)

# Generate .mac files for each row in the DataFrame
for index, row in df.iterrows():
    formatted_policy = str(row.get('Policy', '')).zfill(8)  
    formatted_month = str(row.get('Month', '')).zfill(2)  
    formatted_date = str(row.get('Date', '')).zfill(2)  
    ibm_path = row.get('IBM_Path', '')

    # Ensure the IBM_PATH exists and is valid
    if not os.path.exists(ibm_path):
        print(f"Error: IBM_PATH '{ibm_path}' does not exist.")
        continue  

    # Handle last record differently
    if index == len(df) - 1:
        macro_chain='1'


    # Format the content by replacing placeholders with values from the DataFrame
    formatted_content = mac_template.format(
        Policy=formatted_policy,
        Premium=row.get('Premium', ''),
        Month=formatted_month,
        Date=formatted_date,
        Year=row.get('Year', ''),
        IBM_Path=ibm_path,
        TC_Next=index + 2,
        macro_chain='0',
        cell=index + 2
    )

    # Create the file path for the .mac file
    mac_filename = f'diary_{index + 1}.mac'
    mac_file_path = os.path.join(ibm_path, mac_filename)

    # Save the formatted content to the .mac file in the specified IBM_PATH
    with open(mac_file_path, 'w', encoding='utf-8') as file:
        file.write(formatted_content)

print("MAC files generated successfully!")
