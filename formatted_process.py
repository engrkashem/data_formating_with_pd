import pandas as pd

# Load the Excel file
file_path = "./data/formatted.xlsx"
xls = pd.ExcelFile(file_path)

# Read the first sheet
sheet1 = pd.read_excel(xls, sheet_name=xls.sheet_names[0])

# Select relevant columns and forward-fill supervisor names
sheet1_cleaned = sheet1.iloc[:, :6]  # Taking only necessary columns
# Forward fill Supervisors to handle NaN values in merged rows
sheet1_cleaned['Supervisors'] = sheet1_cleaned['Supervisors'].fillna(method='ffill')

# Drop rows where Feedback Focus is 'TOTAL'
sheet1_cleaned = sheet1_cleaned[sheet1_cleaned['Feedback Focus'] != 'TOTAL']

# Reshape data to long format
formatted_sheet1 = sheet1_cleaned.melt(
    id_vars=['Supervisors', 'Feedback Focus'], 
    value_vars=['1st Draft', '2nd Draft', '3rd Draft'], 
    var_name='Drafts', 
    value_name='Number of Feedback'
)

# Cleaning up draft column names
formatted_sheet1['Drafts'] = formatted_sheet1['Drafts'].str.replace("\xa0", " ")  # Remove non-breaking spaces

# Define custom order for sorting
draft_order = ["1st Draft", "2nd Draft", "3rd Draft"]
formatted_sheet1["Drafts"] = pd.Categorical(formatted_sheet1["Drafts"], categories=draft_order, ordered=True)

# Sort by Supervisors and Drafts
# formatted_sheet1_sorted = formatted_sheet1.sort_values(by=["Supervisors", "Drafts"])

# Extract numeric part from Supervisors for proper sorting
formatted_sheet1["Supervisors_Numeric"] = (
    formatted_sheet1["Supervisors"].str.extract("(\d+)").astype(float)
)

# Sort by extracted numeric Supervisors and Drafts
formatted_sheet1_sorted = formatted_sheet1.sort_values(
    by=["Supervisors_Numeric", "Drafts"]
).drop(columns=["Supervisors_Numeric"])  # Drop helper column after sorting


# Save to Excel
output_file = "./formatted_output.xlsx"
formatted_sheet1_sorted.to_excel(output_file, index=False)

# formatted_sheet1.to_excel("formatted_output.xlsx", index=False)




"""
Get sheet names
sheet_names = xls.sheet_names

Read both sheets
sheet1 = pd.read_excel(xls, sheet_name=sheet_names[0])
sheet2 = pd.read_excel(xls, sheet_name=sheet_names[1])
"""
