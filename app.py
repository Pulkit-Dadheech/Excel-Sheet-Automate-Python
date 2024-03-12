import openpyxl as xl

wb = xl.load_workbook('transactions.xlsx')
sheet = wb['Sheet1']
cell = sheet.cell(row=1, column=1)
print(cell.value)
print(sheet.max_row)

for row in range(2, sheet.max_row + 1):
    try:
        third_cell_value = int(sheet.cell(row=row, column=3).value)
        discounted_value = third_cell_value * 0.9
        sheet.cell(row=row, column=5).value = discounted_value
    except (ValueError, TypeError):
        continue

wb.save('transactions.xlsx')