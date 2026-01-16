import pdfplumber

pdfs = [
    'data/DE PAZ FRANCO QUINTILIANA/E1-3 DOCUMENTOS JUSTIFICATIVOS/E1-3-6-1 CEE FINAL.pdf',
    'data/DE PAZ FRANCO QUINTILIANA/E1-3 DOCUMENTOS JUSTIFICATIVOS/E1-3-6-2 REGISTRO.pdf',
    'data/DE PAZ FRANCO QUINTILIANA/E1-4 OTROS DOCUMENTOS JUSTIFICATIVOS/DNI.pdf',
]

for pdf_path in pdfs:
    print(f"\n{'='*80}")
    print(f"📄 {pdf_path.split('/')[-1]}")
    print('='*80)
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            
            print(f"Text length: {len(text) if text else 0}")
            print(f"Has images: {len(page.images) > 0}")
            
            if text and len(text) > 0:
                print(f"Sample text: {text[:200]}")
            else:
                print("❌ No text - likely scanned image")
    except Exception as e:
        print(f"Error: {e}")
