import pandas as pd
import re
#Loading the excel file

file_path = 'D:/Wasif Data/Python Script/Worksheet-Requirements.xlsx'
df = pd.read_excel(file_path,sheet_name = 'Sheet1')

#Extracting the "Make & Type " columns to get unique models
unique_models = df['Make & Type'].unique()

print(unique_models)

# Create a function to clean the sheet names
def clean_sheet_name(name):
    # Replace invalid characters with an underscore or remove them
    return re.sub(r'[\/:*?"<>|]', '_', name)[:30]  # Excel sheet names cannot be longer than 31 characters

# Create a function to escape regex special characters
def escape_special_characters(string):
    return re.escape(string)

# Create a new Excel writer to save new sheets
with pd.ExcelWriter('split_models.xlsx', engine='openpyxl') as writer:
    # Loop over each unique model
    for model in df['Make & Type'].unique():
        # Clean the model name for use as a sheet name
        sheet_name = clean_sheet_name(model)
        
        # Escape any special characters in the model name
        escaped_model = escape_special_characters(model)

        # Filter rows that contain the current model using the escaped model
        filtered_df = df[df['Make & Type'].str.contains(escaped_model, case=False, na=False)]
        
        # Write filtered data to a new sheet with the cleaned model name
        filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print("Data has been split into new sheets in 'split_models.xlsx'")