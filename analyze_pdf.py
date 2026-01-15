import pdfplumber

pdf_path = 'data/sample_1/E01-1-1_CONTRATO CESION AHORROS.pdf'

with pdfplumber.open(pdf_path) as pdf:
    for i, page in enumerate(pdf.pages):
        print(f"\n{'='*80}")
        print(f"PAGE {i+1}")
        print(f"{'='*80}")
        text = page.extract_text()
        print(text)
