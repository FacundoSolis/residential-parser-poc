from src.parsers.factura_parser import FacturaParser
from src.parsers.certificado_parser import CertificadoParser

files = ['data/sample_1/E1-3-3 FACTURA.pdf','data/sample_1/E1-3-5 CERTIFICADO INSTALADOR.pdf']
for f in files:
    if 'FACTURA' in f.upper():
        p = FacturaParser(f)
    else:
        p = CertificadoParser(f)
    p.extract_text()
    res = p.parse()
    print(f)
    print('  isolation_type ->', res.get('isolation_type'))
    print('  inferred_by_brand ->', res.get('isolation_inferred_by_brand'))
