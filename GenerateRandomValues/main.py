from pathlib import Path

from ExcelClass import Excel

file_name = 'random.xlsx'
sheet_name = 'Sheet1'
# file = open(file_name, 'w')

start_row = 1
end_row = 97
start_column = 1
end_column = 56
min_value = -987654
max_value = 987654

path = Path()
print(path.name)
excel = Excel(path, file_name, sheet_name)
excel.generate_random_integers(start_row=start_row,
                               end_row=end_row,
                               start_column=start_column,
                               end_column=end_column,
                               min_value=min_value,
                               max_value=max_value)
excel.workbook_file.save(file_name)
