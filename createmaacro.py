import os
import pandas as pd
import numpy as np

# Load the Excel file into a DataFrame
input_file = 'data_l1j.xlsx'  # Replace with your Excel file name
df = pd.read_excel(input_file)

# Read the .txt file containing the template
with open('mac_template_l1j.txt', 'r', encoding='utf-8') as template_file:
    mac_template = template_file.read()

# Replace NaN values in the DataFrame with empty strings
df = pd.read_excel(input_file).replace(np.nan, '', regex=True)

# Generate .mac files for each row in the DataFrame
for index, row in df.iterrows():
    # Generate the full path for the .mac file
    formatted_agent = str(row.get('Agent', '')).zfill(8)  # Ensure it's 8 digits
    formatted_Reff_Partner = str(row.get('Reff_Partner', '')).zfill(8)  # Ensure it's 8 digits

    # Get the IBM_PATH from the row
    ibm_path = row.get('IBM_Path', '')  # Example: 'C:\\Fileshare_GIT_Pru\\CreationMacro'

    # Ensure the IBM_PATH exists and is valid
    if not os.path.exists(ibm_path):
        print(f"Error: IBM_PATH '{ibm_path}' does not exist.")
        continue  # Skip this iteration if the path is invalid

    # Format the content by replacing placeholders with values from the DataFrame
    formatted_content = mac_template.format(
        Surname=row.get('Surname', ''),
        Product=row.get('Product', ''),
        Payment_Method=row.get('Payment_Method', ''),
        Billing_Frequency=f"{row.get('Billing_Frequency', ''):02}",  # Ensures two digits
        Sum_Insured=row.get('Sum_Insured', ''),
        Deductible=row.get('Deductible', ''),
        Plan_Type=row.get('Plan_Type', ''),
        EPOL=row.get('EPOL', ''),
        ESUB=row.get('ESUB', ''),
        MAILFLAG=row.get('MAILFLAG', ''),
        Gender=row.get('Gender', ''),
        Until_Issue=row.get('Until_Issue', ''),
        Mortality=row.get('Mortality', ''),
        DOB=row.get('DOB', ''),
        Waiver_Code=row.get('Waiver_Code', ''),
        Risk_Cess_Age=row.get('Risk_Cess_Age', ''),
        Premi_Term=row.get('Premi_Term', ''),
        Agent=formatted_agent,  # Add the formatted agent number here
        Reff_Partner=formatted_Reff_Partner,  # Add the formatted agent number here
        Referral=row.get('Referral', ''),
        IBM_Path=ibm_path,  # Set IBM_Path to the full path of the .mac file
        TC_Next=index + 2,
        cell=index + 2
    )

    # Create the file path for the .mac file
    mac_filename = f'l1j_{index + 1}.mac'  # The file name (e.g., bospo_1.mac)
    mac_file_path = os.path.join(ibm_path, mac_filename)  # Combine IBM_PATH with the filename

    # Save the formatted content to the .mac file in the specified IBM_PATH
    with open(mac_file_path, 'w', encoding='utf-8') as file:
        file.write(formatted_content)

print("MAC files generated successfully!")
