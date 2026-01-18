"""
Matrix Generator - Combines data from all document parsers into Excel output
IMPROVED: More flexible document detection for various naming conventions
"""

import openpyxl
from openpyxl.styles import Font, Alignment
from pathlib import Path
from typing import Dict, Any, Optional
import re

from src.parsers.contrato_parser import ContratoParser
from src.parsers.certificado_parser import CertificadoParser
from src.parsers.factura_parser import FacturaParser
from src.parsers.declaracion_parser import DeclaracionParser
from src.parsers.cee_parser import CeeParser
from src.parsers.registro_parser import RegistroParser
from src.parsers.dni_parser import DniParser
from src.parsers.calculo_parser import CalculoParser


class MatrixGenerator:
    """Generates correspondence matrix Excel from parsed documents"""
    
    PARSERS = {
        'CONTRATO': ContratoParser,
        'CERTIFICADO': CertificadoParser,
        'FACTURA': FacturaParser,
        'DECLARACION': DeclaracionParser,
        'CEE': CeeParser,
        'REGISTRO': RegistroParser,
        'DNI': DniParser,
        'CALCULO': CalculoParser,
    }
    
    # Flexible patterns for document identification (checked in order)
    # Format: list of (pattern, exclude_pattern) tuples
    DOCUMENT_PATTERNS = {
        'CONTRATO': [
            (r'contrato.*cesi[oó]n', None),
            (r'cesi[oó]n.*ahorro', None),
            (r'E0?4[-_\s]?1[-_\s]?1', None),  # E4-1-1 or E04-1-1
            (r'contrato', None),
            (r'convenio.*cae', None),
        ],
        'CEE': [
            (r'cee.*final', None),
            (r'E0?4[-_\s]?3[-_\s]?6[-_\s]?1', None),  # E4-3-6-1
            (r'certificado.*eficiencia', None),
        ],
        'REGISTRO': [
            (r'E0?4[-_\s]?3[-_\s]?6[-_\s]?2', None),  # E4-3-6-2
            (r'registro', r'cee|final'),  # REGISTRO but not if CEE/FINAL in name
        ],
        'CERTIFICADO': [
            (r'certificado.*instalador', None),
            (r'E0?4[-_\s]?3[-_\s]?5', None),  # E4-3-5
            (r'cert.*instalador', None),
        ],
        'FACTURA': [
            (r'factura', None),
            (r'E0?4[-_\s]?3[-_\s]?3', None),  # E4-3-3
        ],
        'DECLARACION': [
            (r'declaraci[oó]n', None),
            (r'E0?4[-_\s]?3[-_\s]?2', None),  # E4-3-2
        ],
        'DNI': [
            (r'dni', None),
            (r'E0?4[-_\s]?4[-_\s]?1', None),  # E4-4-1
        ],
        'CALCULO': [
            (r'calculo', None),
            (r'c[aá]lculo', None),
            (r'ui.*rtotal', None),
            (r'E0?4[-_\s]?4[-_\s]?2', None),  # E4-4-2
        ],
        'FICHA': [
            (r'ficha.*res', None),
            (r'E0?4[-_\s]?3[-_\s]?1', None),  # E4-3-1
            (r'ficha', None),
        ],
        'INFORME': [
            (r'informe.*fotogr[aá]fico', None),
            (r'E0?4[-_\s]?3[-_\s]?4', None),  # E4-3-4
            (r'fotografico', None),
        ],
    }
    
    def __init__(self, folder_path: str):
        self.folder_path = Path(folder_path)
        self.parsed_data = {}
        self.file_mapping = {}
        
    def identify_document_type(self, filepath: Path) -> Optional[str]:
        """
        Identify document type using flexible pattern matching.
        Checks filename AND parent folder name for patterns.
        """
        # Combine filename and parent folder for matching
        filename = filepath.stem
        parent_name = filepath.parent.name if filepath.parent != self.folder_path else ""
        search_text = f"{parent_name} {filename}".upper()
        
        # Also check full path for deeply nested structures
        full_path_text = str(filepath).upper()
        
        # Check patterns in priority order
        # CEE must be checked before REGISTRO to avoid mismatches
        priority_order = ['CEE', 'CERTIFICADO', 'CONTRATO', 'REGISTRO', 'FACTURA', 
                         'DECLARACION', 'DNI', 'CALCULO', 'FICHA', 'INFORME']
        
        for doc_type in priority_order:
            patterns = self.DOCUMENT_PATTERNS.get(doc_type, [])
            for pattern, exclude_pattern in patterns:
                # Check if pattern matches
                if re.search(pattern, search_text, re.IGNORECASE) or \
                   re.search(pattern, full_path_text, re.IGNORECASE):
                    # Check if exclude pattern also matches (skip if it does)
                    if exclude_pattern and re.search(exclude_pattern, search_text, re.IGNORECASE):
                        continue
                    return doc_type
        
        return None
    
    def parse_all_documents(self):
        """Parse all documents in folder recursively"""
        print(f"\n📄 Parsing documents from: {self.folder_path}")
        
        # Supported extensions
        supported_extensions = {'.pdf', '.jpg', '.jpeg', '.png'}
        excel_extensions = {'.xlsx', '.xls'}
        
        # Track found documents to avoid duplicates
        found_docs = {}
        
        # Recursively find all files
        all_files = list(self.folder_path.rglob("*"))
        
        # First pass: identify all documents
        for filepath in all_files:
            if not filepath.is_file():
                continue
                
            ext = filepath.suffix.lower()
            
            if ext in supported_extensions:
                doc_type = self.identify_document_type(filepath)
                if doc_type and doc_type in self.PARSERS:
                    # If we already found this type, keep the one with more specific name
                    if doc_type in found_docs:
                        # Prefer files with the actual keyword in the name
                        existing = found_docs[doc_type]
                        if doc_type.lower() in filepath.name.lower() and \
                           doc_type.lower() not in existing.name.lower():
                            found_docs[doc_type] = filepath
                    else:
                        found_docs[doc_type] = filepath
                        
            elif ext in excel_extensions:
                # Skip output files
                if 'Checks' in filepath.name or 'output' in str(filepath):
                    continue
                doc_type = self.identify_document_type(filepath)
                if doc_type == 'CALCULO':
                    found_docs['CALCULO'] = filepath
        
        # Second pass: parse found documents
        print(f"\n📋 Found {len(found_docs)} documents:")
        for doc_type, filepath in found_docs.items():
            print(f"   {doc_type}: {filepath.name}")
        
        print(f"\n🔍 Parsing...")
        for doc_type, filepath in found_docs.items():
            self.file_mapping[doc_type] = filepath
            
            if doc_type not in self.PARSERS:
                continue
                
            parser_class = self.PARSERS[doc_type]
            
            try:
                parser = parser_class(str(filepath))
                data = parser.parse()
                self.parsed_data[doc_type] = data
                print(f"   ✓ Parsed {doc_type}")
            except Exception as e:
                print(f"   ✗ Error parsing {doc_type}: {e}")
                self.parsed_data[doc_type] = {
                    'document_type': doc_type,
                    'error': str(e)
                }
    
    def generate_excel(self, output_path: str, project_name: str = "Project"):
        """Generate Excel file with correspondence matrix"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Checks"
        
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 30
        for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws.column_dimensions[col].width = 35
        
        headers = {
            'B3': 'Info',
            'C3': 'E1-1-1 CONTRATO CESION AHORROS',
            'D3': 'E1-3-1 FICHA RES020 010',
            'E3': 'E1-3-2 DECLARACION RESPONSABLE',
            'F3': 'E1-3-3 FACTURA',
            'G3': 'E1-3-4 INFORME FOTOGRAFICO',
            'H3': 'E1-3-5 CERTIFICADO INSTALADOR',
            'I3': 'E1-3-6-1\nCEE FINAL',
            'J3': 'E1-3-6-2 REGISTRO CEE',
            'K3': 'E1-4-1 \nDNI',
            'L3': 'E1-4-2 CALCULO UI RTOTAL',
        }
        
        for cell, value in headers.items():
            ws[cell] = value
            ws[cell].font = Font(bold=True)
            ws[cell].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        current_row = 4
        
        # HOME OWNER section
        ws[f'A{current_row}'] = 'HOME OWNER'
        ws[f'A{current_row}'].font = Font(bold=True)
        
        # Name
        ws[f'B{current_row}'] = 'Name'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'homeowner_name')
        ws[f'E{current_row}'] = self._get_value('DECLARACION', 'homeowner_name')
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'homeowner_name')
        ws[f'K{current_row}'] = self._get_value('DNI', 'name')
        ws[f'L{current_row}'] = self._get_value('CALCULO', 'client_name')
        current_row += 1
        
        # DNI
        ws[f'B{current_row}'] = 'DNI'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'homeowner_dni')
        ws[f'E{current_row}'] = self._get_value('DECLARACION', 'homeowner_dni')
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'homeowner_dni')
        ws[f'K{current_row}'] = self._get_value('DNI', 'dni_number')
        current_row += 1
        
        # Address
        ws[f'B{current_row}'] = 'Addres'
        homeowner_addr = self._get_value('CONTRATO', 'homeowner_address')
        if not homeowner_addr:
            homeowner_addr = self._get_value('CONTRATO', 'location')
        ws[f'C{current_row}'] = homeowner_addr
        ws[f'E{current_row}'] = self._get_value('DECLARACION', 'homeowner_address')
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'homeowner_address')
        current_row += 1
        
        # Phone
        ws[f'B{current_row}'] = 'Phone number'
        phone = self._get_value('CONTRATO', 'homeowner_phone')
        if not phone:
            phone = self._get_value('FACTURA', 'homeowner_phone')
        ws[f'C{current_row}'] = phone
        current_row += 1
        
        # Mail
        ws[f'B{current_row}'] = 'Mail'
        email = self._get_value('CONTRATO', 'homeowner_email')
        if not email:
            email = self._get_value('FACTURA', 'homeowner_email')
        ws[f'C{current_row}'] = email
        current_row += 1
        
        # Signatures
        ws[f'B{current_row}'] = 'Signatures (+format)'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'homeowner_signatures')
        ws[f'E{current_row}'] = self._get_value('DECLARACION', 'signature')
        current_row += 1
        
        # ACT section
        ws[f'A{current_row}'] = 'ACT'
        ws[f'A{current_row}'].font = Font(bold=True)
        
        # Code
        ws[f'B{current_row}'] = 'Code (010/020)'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'act_code')
        ws[f'E{current_row}'] = self._get_value('DECLARACION', 'act_code')
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'act_code')
        ws[f'L{current_row}'] = self._get_value('CALCULO', 'act_code')
        current_row += 1
        
        # Energy savings
        ws[f'B{current_row}'] = 'Energy savings (kWh)'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'energy_savings')
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'energy_savings')
        ws[f'L{current_row}'] = self._get_value('CALCULO', 'ae')
        current_row += 1
        
        # Start date
        ws[f'B{current_row}'] = 'Start date'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'start_date')
        current_row += 1
        
        # Finish date
        ws[f'B{current_row}'] = 'Finish date'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'finish_date')
        current_row += 1
        
        # Address
        ws[f'B{current_row}'] = 'Adress'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'location')
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'address')
        ws[f'I{current_row}'] = self._get_value('CEE', 'address')
        ws[f'J{current_row}'] = self._get_value('REGISTRO', 'address')
        current_row += 1
        
        # Catastral ref
        ws[f'B{current_row}'] = 'Catastral ref'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'catastral_ref')
        ws[f'E{current_row}'] = self._get_value('DECLARACION', 'catastral_ref')
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'catastral_ref')
        ws[f'I{current_row}'] = self._get_value('CEE', 'catastral_ref')
        ws[f'J{current_row}'] = self._get_value('REGISTRO', 'catastral_ref')
        current_row += 1
        
        # UTM coordinates
        ws[f'B{current_row}'] = 'UTM coordinates'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'utm_coordinates')
        current_row += 1
        
        # Lifespan
        ws[f'B{current_row}'] = 'Lifespan (años)'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'lifespan')
        current_row += 1
        
        # Surface
        ws[f'B{current_row}'] = 'Surface (m²)'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'surface')
        ws[f'L{current_row}'] = self._get_value('CALCULO', 'area_afect')
        current_row += 1
        
        # Climatic zone
        ws[f'B{current_row}'] = 'Climatic zone'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'climatic_zone')
        ws[f'L{current_row}'] = self._get_value('CALCULO', 'zone_climatique')
        current_row += 1
        
        # Invoice info
        ws[f'B{current_row}'] = 'Invoice number'
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'invoice_number')
        current_row += 1
        
        ws[f'B{current_row}'] = 'Invoice date'
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'invoice_date')
        current_row += 1
        
        ws[f'B{current_row}'] = 'Invoice amount'
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'amount')
        current_row += 1
        
        # Registration info
        ws[f'B{current_row}'] = 'Registration number'
        ws[f'J{current_row}'] = self._get_value('REGISTRO', 'registration_number')
        current_row += 1
        
        ws[f'B{current_row}'] = 'Registration date'
        ws[f'J{current_row}'] = self._get_value('REGISTRO', 'registration_date')
        current_row += 1
        
        # CEE info
        ws[f'B{current_row}'] = 'Certification date'
        ws[f'I{current_row}'] = self._get_value('CEE', 'certification_date')
        current_row += 1
        
        # Save
        wb.save(output_path)
        print(f"\n✅ Excel generated: {output_path}")
    
    def _get_value(self, doc_type: str, field: str) -> str:
        """Get value from parsed data"""
        if doc_type not in self.parsed_data:
            return ""
        
        value = self.parsed_data[doc_type].get(field, "")
        
        if value == "NOT FOUND" or value is None:
            return ""
        
        return str(value)