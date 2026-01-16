from src.parsers.contrato_parser import ContratoParser

parser = ContratoParser('data/DE PAZ FRANCO QUINTILIANA/E1-1 CONVENIO CAE/E1-1-1 CONTRATO CESION AHORROS.pdf')
result = parser.parse()

print("\n=== CONTRATO PARSED DATA ===")
for key, value in result.items():
    print(f"{key}: {value}")
