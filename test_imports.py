# Syntax check for all modules
import ast
import sys

files = [
    'excel_automation.py',
    'trade_processor.py',
    'Confo_Generator.py'
]

all_ok = True
for f in files:
    try:
        with open(f, 'r') as file:
            ast.parse(file.read())
        print(f"OK: {f}")
    except SyntaxError as e:
        print(f"ERROR: {f} - {e}")
        all_ok = False

if all_ok:
    print("\nAll files have valid syntax!")
else:
    print("\nSome files have syntax errors!")
    sys.exit(1)
