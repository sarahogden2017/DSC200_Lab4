import openpyxl

workbook = openpyxl.load_workbook('data/Lab4Data.xlsx')
sheet = workbook['Table 9 ']
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
csv_rows = []
for row in sheet.iter_rows(min_row=15, max_row=211, min_col=1, max_col=31):
    country_name = row[1].value
    for subcategory, col in subcategories.items():
        value = row[col - 1].value
        try:
            new_row = [country_name, subcategory, float(value)]
            csv_rows.append(new_row)
        except:
            continue
