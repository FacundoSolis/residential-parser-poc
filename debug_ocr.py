from src.parsers.base_parser import BaseDocumentParser

p = BaseDocumentParser("data/DEL CUETO DEL POZO MANOLO 50/E1-1 CONVENIO CAE/E1-1-1  CONTRATO CESION DE AHOROS.pdf")
txt = p.extract_text()

print("CHARS:", len(txt))
print("---- START ----")
print(txt[:4000])   # primer trozo
print("---- END ----")
