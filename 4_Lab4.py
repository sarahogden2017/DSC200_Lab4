# These lines import necessary tools to work with Excel files and CSV files
import openpyxl
import csv

# Open the Excel file named 'Lab4Data.xlsx'
workbook = openpyxl.load_workbook('data/Lab4Data.xlsx')
# Select the sheet named 'Table 9 ' from the Excel file
sheet = workbook['Table 9 ']

# This dictionary maps category names to their column numbers in the Excel sheet
subcategories = {
    'child_labor_total': 5,
    'child_labor_male': 7,
    'child_labor_female': 9,
    'child_marriage_by_15': 11,
    'child_marriage_by_18': 13,
    'birth_registrations_total': 15,
    'female_genital_mutilation_prevalence_women': 17,
    'female_genital_mutilation_prevalence_girls': 19,
    'female_genital_mutilation_attitudes_support': 21,
    'justification_wife_beating_male': 23,
    'justification_wife_beating_female': 25,
    'violent_discipline_total': 27,
    'violent_discipline_male': 29,
    'violent_discipline_female': 31
}

# Prepare an empty list to store our processed data
csv_rows = []

# Loop through each row in the Excel sheet from row 15 to 211
for row in sheet.iter_rows(min_row=15, max_row=211, min_col=1, max_col=31):
    # Get the country name from the second column (index 1)
    country_name = row[1].value
    
    # For each category in our subcategories dictionary
    for subcategory, col in subcategories.items():
        # Get the value from the corresponding column
        value = row[col - 1].value
        try:
            # Try to convert the value to a float (decimal number)
            # If successful, create a new row with country name, category, and value
            new_row = [country_name, subcategory, float(value)]
            # Add this new row to our list of processed data
            csv_rows.append(new_row)
        except:
            # If we can't convert to float (e.g., it's empty or not a number), skip this entry
            continue

# Set the name for our output CSV file
output_file = 'khatrir2.csv'

# Open the output CSV file for writing
with open(output_file, 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)

    # Write the header row to the CSV file
    header_row = ['Country', 'Subcategory', 'Value']
    writer.writerow(header_row)

    # Write all our processed data rows to the CSV file
    writer.writerows(csv_rows)

# Count how many rows of data we processed
num_rows = len(csv_rows)
# Print a message saying how many rows were written to the CSV file
print(f"CSV file '{output_file}' has been created with {num_rows} rows.")
