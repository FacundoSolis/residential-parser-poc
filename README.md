# Residential Energy Certificates Parser - Phase 1

POC for extracting data from Spanish residential energy efficiency certificates.

## Project Structure
```
residential-parser-poc/
├── data/
│   ├── sample_1/       # Sample PDFs for testing
│   └── output/         # Generated Excel files
├── src/
│   ├── parsers/        # Document-specific parsers
│   ├── main.py         # Main processing script
│   └── matrix_generator.py  # Excel output generator
└── requirements.txt
```

## Phase 1 Scope
✅ Extract fields from 11 document types
✅ Generate correspondence matrix in Excel format
✅ One tab per document type

## Installation
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Usage
```bash
python src/main.py --input data/sample_1 --output data/output
```

## Document Types
1. E01-1-1: Contrato Cesión Ahorros
2. E1-3-1: Ficha RES020
3. E1-3-2: Declaración Responsable
4. E1-3-3: Factura
5. E1-3-4: Informe Fotográfico
6. E1-3-5: Certificado Instalador
7. E1-3-6-1: CEE Final
8. E1-3-6-2: Registro CEE
9. DNI
10. Calculo.xlsx
11. CSV data file