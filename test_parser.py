from src.parsers.contrato_parser import ContratoParser

# Test with Contrato
parser = ContratoParser('data/sample_1/E01-1-1_CONTRATO CESION AHORROS.pdf')
result = parser.parse()

print("\n=== CONTRATO CESION AHORROS ===")
for key, value in result.items():
    print(f"{key}: {value}")
