import openpyxl

excel_path = 'data/output/ZALAMA LORA BENITO_Checks.xlsx'
wb = openpyxl.load_workbook(excel_path)
ws = wb.active

print("\n=== EXCEL GENERADO ===\n")
for i, row in enumerate(ws.iter_rows(values_only=True), 1):
    if i > 30:  # Primeras 30 filas
        break
    print(f"Row {i}: {row}")
