import pandas as pd

# Load the Excel file
input_excel_path = 'DailyReport.xlsm'
output_excel_path = 'output.xlsx'

# Read the input Excel file into a dictionary of DataFrames
df_input = pd.read_excel(input_excel_path, sheet_name=None)

# Create an empty DataFrame to store the extracted data
df_output = pd.DataFrame(columns=['Spreadsheet'])

# Iterate through each sheet (section) in the input Excel file
for sheet_name, sheet_data in df_input.items():
    # Create a new row for each entry in the section
    for index, row in sheet_data.iterrows():
        entry_row = row.to_dict()
        entry_row['Spreadsheet'] = sheet_name
        df_output = pd.concat([df_output, pd.DataFrame([entry_row])], ignore_index=True)

        # df_output = df_output.append(entry_row, ignore_index=True)

# Save the extracted data into a new Excel file
df_output.to_excel(output_excel_path, index=False)

print("Data extraction and transformation complete.")
