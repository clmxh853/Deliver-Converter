import openpyxl
import json

wb = openpyxl.load_workbook('Confo_Template_20230317.xlsx', data_only=False)
print('Sheets:', wb.sheetnames)
print()

# Check TBLT sheet
if 'TBLT' in wb.sheetnames:
    tblt = wb['TBLT']
    print('=== TBLT Sheet ===')
    print('Max row:', tblt.max_row)
    print('Max col:', tblt.max_column)
    print('Column N (14th column) sample values (first 10 rows):')
    for i in range(1, min(11, tblt.max_row+1)):
        val = tblt.cell(i, 14).value
        print(f'  Row {i}: {val}')
    print()
    # Show header row
    print('Header (row 1):')
    for i in range(1, min(20, tblt.max_column+1)):
        print(f'  Col {i}: {tblt.cell(1, i).value}')

# Check Template sheet
if 'Template' in wb.sheetnames:
    tmpl = wb['Template']
    print('\n=== Template Sheet ===')
    print('Max row:', tmpl.max_row)
    print('Max col:', tmpl.max_column)
    print('Sample cells with formulas in B1:I46:')
    for row in range(1, min(10, 47)):
        for col in range(2, 10):  # B to I
            cell = tmpl.cell(row, col)
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                print(f'  {cell.coordinate}: {cell.value}')
                break
