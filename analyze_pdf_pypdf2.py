import PyPDF2

pdf_path = 'data/sample_1/E01-1-1_CONTRATO CESION AHORROS.pdf'

with open(pdf_path, 'rb') as file:
    reader = PyPDF2.PdfReader(file)
    
    print(f"Total pages: {len(reader.pages)}\n")
    
    for i, page in enumerate(reader.pages):
        print(f"\n{'='*80}")
        print(f"PAGE {i+1}")
        print(f"{'='*80}")
        text = page.extract_text()
        print(text if text else "NO TEXT EXTRACTED")
