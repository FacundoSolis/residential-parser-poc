import openpyxl

excel_path = 'data/output/DE PAZ FRANCO QUINTILIANA_Checks.xlsx'
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

print("\n=== EXCEL GENERADO ===\n")
for i, row in enumerate(ws.iter_rows(values_only=True), 1):
    if i > 25:  # Primeras 25 filas
        break
    print(f"Row {i}: {row}")
