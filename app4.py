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
# 2行目以降を指定

for cell in ws["B"][2:]:
    if cell.value is not None:  # Noneじゃない場合のみ追加
        b_column_values.append(cell.value)


# E列の値を取得
for cell in ws["E"][2:]:
    if cell.value is not None:
        e_column_values.append(cell.value)

# J列の値を取得
for cell in ws["J"][2:]:
    if cell.value is not None:
        j_column_values.append(cell.value)

matched_list = []
for b_column in b_column_values:
    for j_column in j_column_values:
        if b_column == j_column:
            matched_list.append(b_column)

fill_color = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

for row, cell in enumerate(ws["B"][1:], start=2):
    if cell.value in matched_list:
        ws[f"B{row}"].fill = fill_color

for row, cell in enumerate(ws["E"][1:], start=2):
    if cell.value in matched_list:
        ws[f"E{row}"].fill = fill_color


# 名前を付けて保存
wb.save("result001.xlsx")

# print(b_column_values)
# print(e_column_values)
# print(j_column_values)
# print(matched_list)
# print(e_column_values)
# print(j_column_values)
