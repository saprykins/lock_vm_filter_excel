import pandas as pd

# Read the Excel file
file_path = 'input.xlsx'
xl = pd.ExcelFile(file_path)

# List of sheet names you want to process
sheets_to_process = ['ME', 'NE', 'NA', 'SE', 'AT']

# Define columns to keep
columns_to_keep = ['Hostname', 'Target', 'Environment', 'DC']

# Initialize an empty list to store filtered dataframes
filtered_dataframes = []

# Loop through the specified sheets and filter rows based on criteria
for sheet_name in sheets_to_process:
    # Read data from the current sheet
    data = pd.read_excel(xl, sheet_name=sheet_name)
    
    # Apply filters
    filtered_rows = data[
        (data['Aging Since locking'] == 'Not Locked') & 
        (data['statusFactoryVMInTarget'] == 1) & 
        (data['Target'].isin(['MPI-Azure', 'MPI-AWS', 'POD'])) & 
        (data['VM state'] == 'stopped')
    ]
    
    # Check if there are any rows left after filtering
    if not filtered_rows.empty:
        # Add 'DC' column with the sheet name as value
        # filtered_rows['DC'] = sheet_name
        filtered_rows.loc[:, 'DC'] = sheet_name
        
        # Select the desired columns
        filtered_rows = filtered_rows[columns_to_keep]
        
        # Store the filtered dataframe in the list
        filtered_dataframes.append(filtered_rows)

# Concatenate the list of filtered dataframes into one dataframe
if filtered_dataframes:
    filtered_data = pd.concat(filtered_dataframes, ignore_index=True)
    filtered_data.to_csv('output.csv', index=False)
    print('Filtered data has been saved to output.csv')
else:
    print('No data found after applying the filter criteria.')
