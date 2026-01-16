import pdfplumber

pdfs = [
    ('sample_1 CEE', 'data/sample_1/E1-3-6-1 CEE FINAL.pdf'),
    ('sample_1 REGISTRO', 'data/sample_1/E1-3-6-2 REGISTRO.pdf'),
    ('DE PAZ CEE', 'data/DE PAZ FRANCO QUINTILIANA/E1-3 DOCUMENTOS JUSTIFICATIVOS/E1-3-6-1 CEE FINAL.pdf'),
    ('DE PAZ REGISTRO', 'data/DE PAZ FRANCO QUINTILIANA/E1-3 DOCUMENTOS JUSTIFICATIVOS/E1-3-6-2 REGISTRO.pdf'),
]

for name, path in pdfs:
    print(f"\n{'='*80}")
    print(f"📄 {name}")
    print('='*80)
    
    try:
        with pdfplumber.open(path) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            
            print(f"Pages: {len(pdf.pages)}")
            print(f"Text length: {len(text) if text else 0}")
            print(f"Has images: {len(page.images) > 0}")
            
            if text and len(text) > 50:
                print(f"\n✅ HAS TEXT - Sample:")
                print(text[:300])
            else:
                print(f"\n❌ NO TEXT or very little - likely scanned image")
    except Exception as e:
        print(f"❌ Error: {e}")
