import openpyxl

wb = openpyxl.load_workbook('C:\\Users\\pytho\\OneDrive\\Desktop\\FRANK_Project\\files\\製造マスタサンプル.xlsx')

ws = wb.worksheets[0]

values = []
for cell in ws['A':'B']:
    values.append(cell.value)

print(values)
