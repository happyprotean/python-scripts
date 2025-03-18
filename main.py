import json
import time
from openpyxl import load_workbook
from os import path

def main():
  wb = load_workbook(path.join(path.dirname(__file__), 'data.xlsx'))
  sheet = wb['Sheet1']
  # 遍历excel中每一行
  for index, row in enumerate(sheet.values):
    # 跳过表头
    if index == 0:
      continue
    print(row)
    jsonObj = json.loads(row[2])
    jsonObj['c']['aa'] = 'new'
    params = {
        "InstanceType": row[0],
        "ImageId": row[1],  # 选择合适的操作系统镜像ID
        "Placement": {"Zone": row[2]},  # 指定所在的物理机位置
        "SystemDisk": {"DiskSize": 75},
        "DataDisks": [{"DiskSize": 100,"DiskType": "CLOUD_SSD"}],
        "VirtualPrivateCloud":
            {"VpcId": "vpc-o6zsvhvr",
            "SubnetId": "subnet-n4g05gye"
             }  # 指定所属的VPC ID
    }
    print(jsonObj)
    # 睡眠3秒
    time.sleep(3)

  # 读取特定单元格的值
  # cell_value = sheet["A1"].value
  # print(f"A1 单元格的值: {cell_value}")

  # 关闭工作簿（可选）
  wb.close()

if __name__ == '__main__':
  main()