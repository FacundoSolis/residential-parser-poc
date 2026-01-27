import pdfplumber

pdf_path = 'data/ZALAMA LORA BENITO/ZALAMA LORA BENITO/E1-3 DOCUMENTOS JUSTIFICATIVOS/E1-3-2 DECLARACION RESPONSABLE.pdf'

with pdfplumber.open(pdf_path) as pdf:
    for i, page in enumerate(pdf.pages):
        print(f"\n{'='*80}")
        print(f"PAGE {i+1}")
        print(f"{'='*80}")
        text = page.extract_text()
        print(text)
