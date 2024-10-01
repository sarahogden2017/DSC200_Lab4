import openpyxl

workbook = openpyxl.load_workbook('data/Lab4Data.xlsx')
sheet = workbook['Table 9 ']

for row in sheet.iter_rows(min_row=5, max_row=7, min_col=1, max_col=31):
    row_json = {}
    for for idx, cell in enumerate(row):
        row_json[header[idx]] = cell.value


# 
data_header = {
    'child_labor_total': 7,
    'child_labor_male': 9,
}
