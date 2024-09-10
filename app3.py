import openpyxl

wb = openpyxl.load_workbook(
    "C:\\Users\\pytho\\OneDrive\\Desktop\\FRANK_Project\\files\\製造マスタサンプル.xlsx"
)
sheet = wb.get_sheet_by_name('Sheet1')

for cell_obj in list(sheet.columns)[1]:
    print(cell_obj.value)
