import pdfplumber

pdf_path = 'data/DE PAZ FRANCO QUINTILIANA/E1-3 DOCUMENTOS JUSTIFICATIVOS/E1-3-1 FICHA RES020 CUMPLIMENTADA.pdf'

with pdfplumber.open(pdf_path) as pdf:
    print(f"Total pages: {len(pdf.pages)}\n")
    for i, page in enumerate(pdf.pages):
        print(f"\n{'='*80}")
        print(f"PAGE {i+1}")
        print(f"{'='*80}")
        text = page.extract_text()
        print(text)
