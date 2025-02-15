import pandas as pd

# Load the Excel file
file_path = "./data/11_Supervisors_feedback_focus.xlsx"
xls = pd.ExcelFile(file_path)

# Read the first sheet
sheet2 = pd.read_excel(xls, sheet_name=xls.sheet_names[1])

# Clean column names by replacing non-breaking spaces and extra spaces
sheet2.columns = sheet2.columns.str.replace("\xa0", " ", regex=True).str.strip()
sheet2.columns = sheet2.columns.str.replace("\s+", " ", regex=True)

# Remove empty or improperly formatted categories
sheet2 = sheet2[sheet2['Feedback Focus'].str.strip() != ""]  # Removes blank categories
sheet2 = sheet2.dropna(subset=['Feedback Focus'])  # Drops NaN values in the category column

# Ensure 'Number of Feedback' column is numeric
sheet2['Number of Feedback'] = pd.to_numeric(sheet2['Number of Feedback'], errors='coerce').fillna(0)

# Aggregate data by 'Feedback Focus' (Category), summing 'Number of Feedback'
aggregated_data = sheet2.groupby('Feedback Focus')['Number of Feedback'].agg(['sum', 'mean', 'std']).reset_index()

# Rename columns to match the required format
aggregated_data.columns = ['Category', 'AF', 'M', 'SD']

# Compute total absolute frequency
af_total = aggregated_data['AF'].sum()

# Compute Relative Frequency (RF) as a percentage
aggregated_data['RF'] = (aggregated_data['AF'] / af_total) * 100

# Round values for better readability
aggregated_data = aggregated_data.round(2)

# Define column order explicitly
column_order = ['Category', 'AF', 'RF', 'M', 'SD']

# Reorder the DataFrame to maintain correct column order
aggregated_data = aggregated_data[column_order]

# Create the total row as a DataFrame with column order
total_row = pd.DataFrame([[  
    'Total',
    round(aggregated_data['AF'].sum(), 2),
    100.00,
    round(aggregated_data['M'].mean(), 2),
    round(aggregated_data['SD'].mean(), 2)
]], columns=column_order)

# Append total row and enforce column order
final_table = pd.concat([aggregated_data, total_row], ignore_index=True)[column_order]



print(final_table)

# Save to Excel
output_file = "./statistics_output.xlsx"
final_table.to_excel(output_file, index=False)

