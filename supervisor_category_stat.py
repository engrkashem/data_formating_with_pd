import pandas as pd

# Load the Excel file
file_path = "./data/Function_Feedback_15_2_25.xlsx"
xls = pd.ExcelFile(file_path)

# Read the first sheet
sheet1 = pd.read_excel(xls, sheet_name=xls.sheet_names[0])


# Clean column names by replacing non-breaking spaces and extra spaces
sheet1.columns = sheet1.columns.str.replace("\xa0", " ", regex=True).str.strip()
sheet1.columns = sheet1.columns.str.replace("\s+", " ", regex=True)

# Remove empty or improperly formatted categories
sheet1 = sheet1[sheet1['Supervisors'].str.strip() != ""]  # Removes blank categories
sheet1 = sheet1.dropna(subset=['Supervisors'])  # Drops NaN values in the category column

# Ensure 'Number of Feedback' column is numeric
sheet1['Total Feedback'] = pd.to_numeric(sheet1['Total Feedback'], errors='coerce').fillna(0)

# Aggregate data by 'Supervisors' and 'Feedback Focus' (Category), summing 'Number of Feedback'
aggregated_data = sheet1.groupby(['Supervisors', 'Feedback Function'])['Total Feedback'].agg(['sum', 'mean', 'std']).reset_index()

# Rename columns to match the required format
# aggregated_data.columns = ['Supervisors', 'AF', 'M', 'SD']
aggregated_data.columns = ['Supervisors', 'Category', 'AF', 'M', 'SD']

# Compute total absolute frequency
af_total = aggregated_data['AF'].sum()

# Compute Relative Frequency (RF) as a percentage
aggregated_data['RF'] = (aggregated_data['AF'] / af_total) * 100

# Round values for better readability
aggregated_data = aggregated_data.round(2)

# Define column order explicitly
column_order = ['Supervisors', 'Category', 'AF', 'RF', 'M', 'SD']

# Reorder the DataFrame to maintain correct column order
aggregated_data = aggregated_data[column_order]

# Sort Categories in natural order (S1, S2, ..., S10, S11, ...)
aggregated_data['Sort_Key_Supervisors'] = aggregated_data['Supervisors'].apply(lambda x: float('inf') if x == 'Total' else int(x[1:]) if x[1:].isdigit() else float('inf'))
aggregated_data = aggregated_data.sort_values(by=['Sort_Key_Supervisors', 'Category']).drop(columns=['Sort_Key_Supervisors'])


# Create the total row as a DataFrame with column order
total_row = pd.DataFrame([[  
    'Total',
    'All Categories',
    round(aggregated_data['AF'].sum(), 2),
    100.00,
    round(aggregated_data['M'].mean(), 2),
    round(aggregated_data['SD'].mean(), 2)
]], columns=column_order)

# Append total row and enforce column order
final_table = pd.concat([aggregated_data, total_row], ignore_index=True)[column_order]



print(final_table)

# Save to Excel
output_file = "./statistics_supervisors_category_output.xlsx"
final_table.to_excel(output_file, index=False)

