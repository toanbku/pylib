import io
from openpyxl import load_workbook

def read_file(file_path):
  f = io.open(file_path, 'r', encoding='utf-8')
  return f.read()

def read_file_list(file_path):
  f = io.open(file_path, 'r', encoding='utf-8')
  l = f.read().split('\n')
  return l

def write_file(file_path, content):
  f = io.open(file_path, 'w', encoding='utf-8')
  f.write(content)
  f.close()

def write_file_list(file_path, content):
  f = io.open(file_path, 'w', encoding='utf-8')
  f.write('\n'.join(content))
  f.close()

def read_cell(file_path, sheet_name, cell_name):
  wb = load_workbook(file_path)
  sheet_ranges = wb[sheet_name]
  return sheet_ranges[cell_name].value

def update_cell(file_path, sheet_name, cell_name, content):
  wb = load_workbook(file_path)
  sheet_ranges = wb[sheet_name]
  sheet_ranges[cell_name].value = content
  wb.close()
  wb.save(file_path)


if __name__ == '__main__':
  file_path = 'res.xlsx'
  sheet_name = 'Sheet1'
  cell_name = 'A3'
  content = 'hihi'
  update_cell(file_path, sheet_name, 'A1', 'Tên')
  update_cell(file_path, sheet_name, 'B1', 'Tuổi')
  L1 = read_file_list('file1.txt')
  L2 = read_file_list('file2.txt')

  for i in range(0, len(L1)):
    cell_name = 'A%s'%(i+2)
    cell_age = 'B%s'%(i+2)
    update_cell(file_path, sheet_name, cell_name, L1[i])
    update_cell(file_path, sheet_name, cell_age, L2[i])
  print('done')