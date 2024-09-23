import openpyxl
from openpyxl.styles import PatternFill

# Excelファイルをロード
wb = openpyxl.load_workbook("files/製造マスタサンプル.xlsx")

# アクティブなシートを選択
ws = wb.active

# B列、E列、J列のすべてのセルの値をリストに保存
b_column_values = []
e_column_values = []
j_column_values = []

# B列の値を取得
for cell in ws["B"]:
    b_column_values.append(cell.value)

# E列の値を取得
for cell in ws["E"]:
    e_column_values.append(cell.value)

# J列の値を取得
for cell in ws["J"]:
    j_column_values.append(cell.value)

matched_list = []
for b_column in b_column_values:
    for j_column in j_column_values:
        if b_column == j_column:
            matched_list.append(b_column)


print(matched_list)
# print(e_column_values)
# print(j_column_values)
