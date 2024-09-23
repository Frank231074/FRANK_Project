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
