import openpyxl as xl

file_name = 'random.csv'
sheet_name = 'random'
file = open(file_name, 'w')
workbook_file = xl.load_workbook(file_name)
sheet = workbook_file[sheet_name]
