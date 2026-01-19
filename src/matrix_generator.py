"""
Matrix Generator - Combines data from all document parsers into Excel output
UPDATED: Matches Travis's Excel template structure
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
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
    
    DOCUMENT_PATTERNS = {
        'CONTRATO': [
            (r'contrato.*cesi[oó]n', None),
            (r'cesi[oó]n.*ahorro', None),
            (r'E0?4[-_\s]?1[-_\s]?1', None),
            (r'contrato', None),
            (r'convenio.*cae', None),
        ],
        'CEE': [
            (r'cee.*final', None),
            (r'E0?4[-_\s]?3[-_\s]?6[-_\s]?1', None),
            (r'certificado.*eficiencia', None),
            (r'certificado.*energ[eé]tico', None),
            (r'certificaci[oó]n.*energ[eé]tica', None),

        ],
        'REGISTRO': [
            (r'E0?4[-_\s]?3[-_\s]?6[-_\s]?2', None),
            (r'registro\s*cee', None),
            (r'registro', r'mercantil'),
        ],

        'CERTIFICADO': [
            (r'certificado.*instalador', None),
            (r'E0?4[-_\s]?3[-_\s]?5', None),
            (r'cert.*instalador', None),
        ],
        'FACTURA': [
            (r'factura', None),
            (r'E0?4[-_\s]?3[-_\s]?3', None),
        ],
        'DECLARACION': [
            (r'declaraci[oó]n', None),
            (r'E0?4[-_\s]?3[-_\s]?2', None),
        ],
        'DNI': [
            (r'dni', None),
            (r'E0?4[-_\s]?4[-_\s]?1', None),
        ],
        'CALCULO': [
            (r'calculo', None),
            (r'c[aá]lculo', None),
            (r'ui.*rtotal', None),
            (r'E0?4[-_\s]?4[-_\s]?2', None),
        ],
        'FICHA': [
            (r'ficha.*res', None),
            (r'E0?4[-_\s]?3[-_\s]?1', None),
            (r'ficha', None),
        ],
        'INFORME': [
            (r'informe.*fotogr[aá]fico', None),
            (r'E0?4[-_\s]?3[-_\s]?4', None),
            (r'fotografico', None),
        ],
    }
    
    def __init__(self, folder_path: str):
        self.folder_path = Path(folder_path)
        self.parsed_data = {}
        self.file_mapping = {}
        
    def identify_document_type(self, filepath: Path) -> Optional[str]:
        """Identify document type using flexible pattern matching."""
        filename = filepath.stem
        parent_name = filepath.parent.name if filepath.parent != self.folder_path else ""
        search_text = f"{parent_name} {filename}".upper()
        full_path_text = str(filepath).upper()
        
        priority_order = ['CEE', 'CERTIFICADO', 'CONTRATO', 'REGISTRO', 'FACTURA', 
                         'DECLARACION', 'DNI', 'CALCULO', 'FICHA', 'INFORME']
        
        for doc_type in priority_order:
            patterns = self.DOCUMENT_PATTERNS.get(doc_type, [])
            for pattern, exclude_pattern in patterns:
                if re.search(pattern, search_text, re.IGNORECASE) or \
                   re.search(pattern, full_path_text, re.IGNORECASE):
                    if exclude_pattern and re.search(exclude_pattern, search_text, re.IGNORECASE):
                        continue
                    return doc_type
        
        return None
    
    def parse_all_documents(self):
        """Parse all documents in folder recursively"""
        print(f"\n📄 Parsing documents from: {self.folder_path}")
        
        supported_extensions = {'.pdf', '.jpg', '.jpeg', '.png'}
        excel_extensions = {'.xlsx', '.xls'}
        
        found_docs = {}
        all_files = list(self.folder_path.rglob("*"))
        
        for filepath in all_files:
            if not filepath.is_file():
                continue
                
            ext = filepath.suffix.lower()
            
            if ext in supported_extensions:
                doc_type = self.identify_document_type(filepath)
                if doc_type and doc_type in self.PARSERS:
                    if doc_type in found_docs:
                        existing = found_docs[doc_type]
                        if doc_type.lower() in filepath.name.lower() and \
                           doc_type.lower() not in existing.name.lower():
                            found_docs[doc_type] = filepath
                    else:
                        found_docs[doc_type] = filepath
                        
            elif ext in excel_extensions:
                if 'Checks' in filepath.name or 'output' in str(filepath):
                    continue
                doc_type = self.identify_document_type(filepath)
                if doc_type == 'CALCULO':
                    found_docs['CALCULO'] = filepath
        
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
        """Generate Excel file with correspondence matrix - matches Travis's template"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Checks"
        
        # Column widths
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 25
        for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
            ws.column_dimensions[col].width = 30
        
        # Header style
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        section_font = Font(bold=True)
        
        # Headers row 3
        headers = {
            'B3': 'Info',
            'C3': 'E1-1-1\nCONTRATO CESION AHORROS',
            'D3': 'E1-3-1\nFICHA RES020',
            'E3': 'E1-3-2\nDECLARACION RESPONSABLE',
            'F3': 'E1-3-3\nFACTURA',
            'G3': 'E1-3-4\nINFORME FOTOGRAFICO',
            'H3': 'E1-3-5\nCERTIFICADO INSTALADOR',
            'I3': 'E1-3-6-1\nCEE FINAL',
            'J3': 'E1-3-6-2\nREGISTRO CEE',
            'K3': 'E1-4-1\nDNI',
            'L3': 'E1-4-2\nCALCULO UI RTOTAL',
        }
        
        for cell, value in headers.items():
            ws[cell] = value
            ws[cell].fill = header_fill
            ws[cell].font = header_font
            ws[cell].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        ws.row_dimensions[3].height = 40
        
        current_row = 4
        
        # ==================== HOME OWNER SECTION ====================
        ws[f'A{current_row}'] = 'HOME OWNER'
        ws[f'A{current_row}'].font = section_font
        
        # Name
        # Name
        ws[f'B{current_row}'] = 'Name'

        contr_name = self._get_value('CONTRATO', 'homeowner_name')
        decl_name  = self._get_value('DECLARACION', 'homeowner_name')
        fact_name  = self._get_value('FACTURA', 'homeowner_name')
        dni_name   = self._get_value('DNI', 'name')
        calc_name  = self._get_value('CALCULO', 'client_name')

        # Columna C: mejor nombre disponible (prioridad práctica con fallback fuerte)
        name_contrato = self._get_value('CONTRATO', 'homeowner_name')
        name_factura = self._get_value('FACTURA', 'homeowner_name')
        name_calc = self._get_value('CALCULO', 'client_name')
        
        ws[f'C{current_row}'] = name_contrato if name_contrato not in ("NOT FOUND", "", None) else (name_factura or name_calc)

        # Mantén las otras columnas como están (sirven para comparar)
        ws[f'E{current_row}'] = decl_name
        ws[f'F{current_row}'] = fact_name
        ws[f'K{current_row}'] = dni_name
        ws[f'L{current_row}'] = calc_name

        current_row += 1

        # DNI
        ws[f'B{current_row}'] = 'DNI'

        contr_dni = self._get_value('CONTRATO', 'homeowner_dni')
        decl_dni  = self._get_value('DECLARACION', 'homeowner_dni')
        fact_dni  = self._get_value('FACTURA', 'homeowner_dni')
        dni_num   = self._get_value('DNI', 'dni_number')

        # Columna C: fallback fuerte - usa contrato si es válido, sino factura
        ws[f'C{current_row}'] = contr_dni if contr_dni not in ("NOT FOUND", "", None) else fact_dni

        ws[f'E{current_row}'] = decl_dni
        ws[f'F{current_row}'] = fact_dni
        ws[f'K{current_row}'] = dni_num

        current_row += 1

        
        # Address
        ws[f'B{current_row}'] = 'Address'

        contr_addr = self._get_value('CONTRATO', 'homeowner_address')
        contr_loc  = self._get_value('CONTRATO', 'location')  # a veces viene bien
        decl_addr  = self._get_value('DECLARACION', 'homeowner_address')
        fact_addr  = self._get_value('FACTURA', 'homeowner_address')
        cert_addr  = self._get_value('CERTIFICADO', 'address')

        ws[f'C{current_row}'] = self._pick_best(fact_addr, cert_addr, contr_addr, contr_loc, decl_addr, min_len=8)

        ws[f'E{current_row}'] = decl_addr
        ws[f'F{current_row}'] = fact_addr

        current_row += 1

        
        # Phone
        ws[f'B{current_row}'] = 'Phone number'

        contr_phone = self._get_value('CONTRATO', 'homeowner_phone')
        fact_phone  = self._get_value('FACTURA', 'homeowner_phone')

        ws[f'C{current_row}'] = self._pick_best(fact_phone, contr_phone, min_len=9)
        ws[f'F{current_row}'] = fact_phone

        current_row += 1

        
        # Mail
        ws[f'B{current_row}'] = 'Mail'

        contr_mail = self._get_value('CONTRATO', 'homeowner_email')
        fact_mail  = self._get_value('FACTURA', 'homeowner_email')

        ws[f'C{current_row}'] = self._pick_best(fact_mail, contr_mail, min_len=6)
        ws[f'F{current_row}'] = fact_mail

        current_row += 1

        
        # Signatures
        ws[f'B{current_row}'] = 'Signatures (+format)'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'homeowner_signatures')
        ws[f'E{current_row}'] = self._get_value('DECLARACION', 'signature')
        current_row += 1
        
        # ==================== ACT SECTION ====================
        ws[f'A{current_row}'] = 'ACT'
        ws[f'A{current_row}'].font = section_font
        
        # Code
        ws[f'B{current_row}'] = 'Code (010/020)'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'act_code')
        ws[f'E{current_row}'] = self._get_value('DECLARACION', 'act_code')
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'act_code')
        ws[f'L{current_row}'] = self._get_value('CALCULO', 'act_code')
        current_row += 1
        
        # Energy savings
        ws[f'B{current_row}'] = 'Energy savings (kWh)'

        contr_ae = self._get_value('CONTRATO', 'energy_savings')
        cert_ae  = self._get_value('CERTIFICADO', 'energy_savings')
        calc_ae  = self._get_value('CALCULO', 'ae')

        # Columna C = mejor valor global con lógica inteligente
        ws[f'C{current_row}'] = self._pick_best_energy_savings(contr_ae, cert_ae, calc_ae)

        ws[f'H{current_row}'] = cert_ae
        ws[f'L{current_row}'] = calc_ae

        current_row += 1

        
        # Start date
        ws[f'B{current_row}'] = 'Start date'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'start_date')
        current_row += 1
        
        # Finish date
        ws[f'B{current_row}'] = 'Finish date'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'finish_date')
        current_row += 1
        
        # Address (ACT)
        ws[f'B{current_row}'] = 'Address'

        act_contr = self._get_value('CONTRATO', 'location')
        act_fact  = self._get_value('FACTURA', 'homeowner_address')
        act_cert  = self._get_value('CERTIFICADO', 'address')

        ws[f'C{current_row}'] = self._pick_best(act_fact, act_cert, act_contr, min_len=8)
        ws[f'F{current_row}'] = act_fact
        ws[f'H{current_row}'] = act_cert
        ws[f'I{current_row}'] = self._get_value('CEE', 'address')
        ws[f'J{current_row}'] = self._get_value('REGISTRO', 'address')
        ws[f'L{current_row}'] = ""  # no metas nombre aquí

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
        
        # Investment
        ws[f'B{current_row}'] = 'Investment (€)'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'investment')
        current_row += 1
        
        # Remuneration
        ws[f'B{current_row}'] = 'Remuneration'
        current_row += 1
        
        # Lifespan
        ws[f'B{current_row}'] = 'Lifespan (años)'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'lifespan')
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'lifespan')
        current_row += 1
        
        # Sell price
        ws[f'B{current_row}'] = 'Sell price (€/kWh)'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'sell_price')
        current_row += 1
        
        # Fp
        ws[f'B{current_row}'] = 'Fp'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'fp')
        ws[f'L{current_row}'] = self._get_value('CALCULO', 'fp')
        current_row += 1
        
        # Ui (Ki)
        ws[f'B{current_row}'] = 'Ui'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'ui')
        ws[f'L{current_row}'] = self._get_value('CALCULO', 'ki')
        current_row += 1
        
        # Uf (Kf)
        ws[f'B{current_row}'] = 'Uf'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'uf')
        ws[f'L{current_row}'] = self._get_value('CALCULO', 'kf')
        current_row += 1
        
        # S (Surface)
        ws[f'B{current_row}'] = 'S'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'surface')
        ws[f'L{current_row}'] = self._get_value('CALCULO', 's')
        current_row += 1
        
        # Climatic zone
        ws[f'B{current_row}'] = 'Climatic zone'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'climatic_zone')
        ws[f'L{current_row}'] = self._get_value('CALCULO', 'zone_climatique')
        current_row += 1
        
        # Isolation Thickness
        ws[f'B{current_row}'] = 'Isolation Thickness'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'isolation_thickness')
        current_row += 1
        
        # Surface coefficient
        ws[f'B{current_row}'] = 'Surface coefficient'
        ws[f'L{current_row}'] = self._get_value('CALCULO', 'porcentaje')
        current_row += 1
        
        # Calculation methodology
        ws[f'B{current_row}'] = 'Calculation methodology (R*t)'
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'calculation_methodology')
        ws[f'L{current_row}'] = self._get_value('CALCULO', 'calculation_methodology')
        current_row += 1

        
        # Photos
        ws[f'B{current_row}'] = 'Photos'
        current_row += 1
        
        # ==================== INSTALLER SECTION ====================
        ws[f'A{current_row}'] = 'INSTALLER'
        ws[f'A{current_row}'].font = section_font
        
        # Installer Name
        ws[f'B{current_row}'] = 'Name'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'installer')
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'installer_name')
        ws[f'H{current_row}'] = self._get_value('CERTIFICADO', 'installer_name')
        current_row += 1
        
        # Installer Address
        ws[f'B{current_row}'] = 'Address'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'installer_address')
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'installer_address')
        current_row += 1
        
        # Installer CIF
        ws[f'B{current_row}'] = 'CIF'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'installer_cif')
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'installer_cif')
        current_row += 1
        
        # ==================== BILL SECTION ====================
        ws[f'A{current_row}'] = 'BILL'
        ws[f'A{current_row}'].font = section_font
        
        # Invoice Number
        ws[f'B{current_row}'] = 'Number'
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'invoice_number')
        current_row += 1
        
        # Invoice Date
        ws[f'B{current_row}'] = 'Date'
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'invoice_date')
        current_row += 1
        
        # Subtotal
        ws[f'B{current_row}'] = 'Subtotal'
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'subtotal')
        current_row += 1
        
        # Deduction
        ws[f'B{current_row}'] = 'Deduction (CAE)'
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'deduction')
        current_row += 1
        
        # A pagar
        ws[f'B{current_row}'] = 'A pagar'
        ws[f'F{current_row}'] = self._get_value('FACTURA', 'amount')
        current_row += 1
        
        # ==================== REGISTRO SECTION ====================
        ws[f'A{current_row}'] = 'REGISTRO'
        ws[f'A{current_row}'].font = section_font
        
        # Registration Number
        ws[f'B{current_row}'] = 'Number'
        ws[f'J{current_row}'] = self._get_value('REGISTRO', 'registration_number')
        current_row += 1
        
        # Registration Date
        ws[f'B{current_row}'] = 'Date'
        ws[f'J{current_row}'] = self._get_value('REGISTRO', 'registration_date')
        current_row += 1
        
        # ==================== CEE SECTION ====================
        ws[f'A{current_row}'] = 'CEE'
        ws[f'A{current_row}'].font = section_font
        
        # Certification Date
        ws[f'B{current_row}'] = 'Certification date'
        ws[f'I{current_row}'] = self._get_value('CEE', 'certification_date')
        current_row += 1
        
        # Estado
        ws[f'B{current_row}'] = 'Estado'
        ws[f'L{current_row}'] = self._get_value('CALCULO', 'estado')
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


    def _is_valid_cell_value(self, v: Any, min_len: int = 2) -> bool:
        if v is None:
            return False
        s = str(v).strip()
        if not s:
            return False
        if s.upper() == "NOT FOUND":
            return False
        if len(s) < min_len:
            return False
        # basura típica OCR
        if s.upper() in {"C", "CL", "CALLE", "AV", "S/N"}:
            return False
        return True

    def _pick_best(self, *candidates: Any, min_len: int = 2) -> str:
        for c in candidates:
            if self._is_valid_cell_value(c, min_len=min_len):
                return str(c).strip()
        return ""

    def _pick_best_energy_savings(self, contr_ae: str, cert_ae: str, calc_ae: str) -> str:
        """Pick best energy savings value with intelligent fallback logic."""
        # Helper to check if value is valid and convert to float
        def is_valid_energy(value: str) -> tuple[bool, float]:
            if not self._is_valid_cell_value(value, min_len=2):
                return False, 0.0
            try:
                num = float(value.replace(",", "."))
                return True, num
            except ValueError:
                return False, 0.0

        contr_valid, contr_num = is_valid_energy(contr_ae)
        cert_valid, cert_num = is_valid_energy(cert_ae)
        calc_valid, calc_num = is_valid_energy(calc_ae)

        # Priority: CERTIFICADO > CALCULO > CONTRATO (but avoid small contract values)
        if cert_valid:
            return cert_ae
        elif calc_valid:
            return calc_ae
        elif contr_valid and contr_num >= 500:  # Only use contract if it's reasonably large
            return contr_ae
        else:
            return ""
