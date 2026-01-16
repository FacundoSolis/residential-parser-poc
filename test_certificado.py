from src.parsers.certificado_parser import CertificadoParser

# Test with DE PAZ certificado
parser = CertificadoParser('data/DE PAZ FRANCO QUINTILIANA/E1-3 DOCUMENTOS JUSTIFICATIVOS/E1-3-5 CERTIFICADO INSTALADOR.pdf')
result = parser.parse()

print("\n=== CERTIFICADO INSTALADOR ===")
for key, value in result.items():
    print(f"{key}: {value}")
