import json
from openpyxl import load_workbook
from os import path

def main():
  wb = load_workbook(path.join(path.dirname(__file__), 'data.xlsx'))
  sheet = wb['Sheet1']
  for index, row in enumerate(sheet.values):
    # 跳过表头
    if index == 0:
      continue
    print(row)
    jsonObj = json.loads(row[2])
    jsonObj['c']['aa'] = 'new'
    print(jsonObj)

  # 读取特定单元格的值
  # cell_value = sheet["A1"].value
  # print(f"A1 单元格的值: {cell_value}")

  # 关闭工作簿（可选）
  wb.close()

if __name__ == '__main__':
  main()