# Excel Automation for Model-Based Data Separation

This project automates the task of separating rows in an Excel file based on unique values in the "Make & Type" column. The script processes a dataset containing various car makes and models, cleans up invalid characters in the model names (for Excel sheet names), and then creates separate sheets for each unique model. It uses **Pandas** for data manipulation and **OpenPyXL** to write Excel files.

## Key Features
- Automatically splits data into separate sheets by unique models.
- Handles invalid characters for Excel sheet names.
- Supports large datasets (e.g., over 1400 rows) efficiently.
- Includes a function to escape special characters in regex patterns.
- Easy to modify for different column names or file paths.

## Technologies Used
- Python
- Pandas
- OpenPyXL
- Regular Expressions (regex)

## How to Run
1. Place your dataset in the specified directory.
2. Run the script, and the output file will be saved as `split_models.xlsx` in the working directory.
