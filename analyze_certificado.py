from src.parsers.certificado_parser import CertificadoParser

pdf_path = "data/DE PAZ FRANCO QUINTILIANA/E1-3 DOCUMENTOS JUSTIFICATIVOS/E1-3-5 CERTIFICADO INSTALADOR.pdf"

parser = CertificadoParser(pdf_path)
parser.extract_text()

print("\n=== CERTIFICADO: TEXT (first 2500 chars) ===\n")
print(parser.text[:2500])

# Buscar líneas relevantes para "Calculation Methodology"
keywords = ["R'T", "R''T", "R´T", "R’’T", "RT", "CTE", "media", "0,83", "0.83"]

print("\n=== LINES AROUND KEYWORDS ===\n")
lines = parser.text.splitlines()

def norm(s: str) -> str:
    return (
        s.replace("’", "'")
         .replace("´", "'")
         .replace("`", "'")
         .replace("″", "''")
         .upper()
    )

for i, ln in enumerate(lines):
    n = norm(ln)
    if any(norm(k) in n for k in keywords):
        print(f"\n--- around line {i} ---")
        for j in range(max(0, i-8), min(len(lines), i+8)):
            print(lines[j])

print("\n=== PARSED FIELD: calculation_methodology ===\n")
print(parser.parse().get("calculation_methodology"))
