from ExcelClass import Excel
from pathlib import Path


def get_existing_files(path):
    files = []
    print('Existing Files:')
    for file in path.glob('*'):
        files.append(file.name)
        print(file.name)
    print('-------')
    return files


input_file_path = Path('input_files')
# excel_file_name = 'transactions.xlsx'
sheet_name = 'Sheet1'
input_files = get_existing_files(input_file_path)
for input_file in input_files:
    excel = Excel(input_file_path, input_file, sheet_name)
    excel.process_excel_file()
    excel.save_as(f'updated_{input_file}')
