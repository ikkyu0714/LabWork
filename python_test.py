import openpyxl

books = '../../研究/アンケートデータ/文化差データ/文化差データ_多数決1なし.xlsx'
book = openpyxl.load_workbook(books)
sheet = book['Sheet']

culture_inclusion_list = []
culture_exclusion_list = []
no_culture_list = []
exception_list = []
value_list = []
for lines in sheet.iter_rows(min_row=2, min_col=1):
    if lines[1].value == '文化差あり(包含関係)':
        culture_inclusion_list.append(lines[0].value)
    elif lines[1].value == '文化差あり(排他関係)':
        culture_exclusion_list.append(lines[0].value)
    elif lines[1].value == '文化差なし':
        no_culture_list.append(lines[0].value)
    elif lines[1].value == '対象外':
        exception_list.append(lines[0].value)

print('文化差あり(包含):{},(排他):{}, 文化差なし:{}, 対象外:{}'.format(len(culture_inclusion_list), len(culture_exclusion_list), len(no_culture_list), len(exception_list)))
print(len(culture_exclusion_list)+len(culture_inclusion_list)+len(no_culture_list)+len(exception_list))
