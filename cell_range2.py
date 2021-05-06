from openpyxl import Workbook
import random
from openpyxl.utils.cell import coordinate_from_string


wb = Workbook()
ws = wb.active

# 1줄씩 데이터 넣기
ws.append(["번호", "영어", "수학"])  # A, B, C
for i in range(1, 11):  # 10개 데이터 넣기
    ws.append([i, random.randint(0, 100), random.randint(0, 100)])

# 전체 rows
# print(tuple(ws.rows))
for row in tuple(ws.rows):
    print(row[1].value)

# 전체 columns
print(tuple(ws.columns))

wb.save("sample.xlsx")
