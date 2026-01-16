"""
Matrix Generator - Combines data from all document parsers into Excel output
"""

import openpyxl
from openpyxl.styles import Font, Alignment
from pathlib import Path
from typing import Dict, Any

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
    
    def __init__(self, folder_path: str):
        self.folder_path = Path(folder_path)
        self.parsed_data = {}
        
    def identify_document_type(self, filename: str) -> str:
        """Identify document type from filename"""
        filename_upper = filename.upper()
        
        if 'CONTRATO' in filename_upper or 'CONVENIO' in filename_upper:
            return 'CONTRATO'
        elif 'CERTIFICADO' in filename_upper:
            return 'CERTIFICADO'
        elif 'FACTURA' in filename_upper:
            return 'FACTURA'
        elif 'DECLARACION' in filename_upper:
            return 'DECLARACION'
        elif 'CEE' in filename_upper and 'FINAL' in filename_upper:
            return 'CEE'
        elif 'REGISTRO' in filename_upper:
            return 'REGISTRO'
        elif 'DNI' in filename_upper:
            return 'DNI'
        elif 'CALCULO' in filename_upper:
            return 'CALCULO'
        elif 'FICHA' in filename_upper:
            return 'FICHA'
        elif 'FOTOGRAFICO' in filename_upper or 'FOTOGRÁFICO' in filename_upper:
            return 'INFORME'
        else:
            return 'UNKNOWN'
    
    def parse_all_documents(self):
        """Parse all documents in folder"""
        print(f"\n📄 Parsing documents from: {self.folder_path}")
        
        # Parse PDFs
        pdf_files = list(self.folder_path.glob("**/*.pdf"))
        for pdf_file in pdf_files:
            doc_type = self.identify_document_type(pdf_file.name)
            
            if doc_type == 'UNKNOWN':
                print(f"⚠️  Skipping unknown document: {pdf_file.name}")
                continue
            
            if doc_type not in self.PARSERS:
                print(f"⚠️  No parser for: {doc_type} ({pdf_file.name})")
                continue
            
            parser_class = self.PARSERS[doc_type]
            parser = parser_class(str(pdf_file))
            
            try:
                data = parser.parse()
                self.parsed_data[doc_type] = data
                print(f"✓ Parsed {doc_type}: {pdf_file.name}")
            except Exception as e:
                print(f"✗ Error parsing {pdf_file.name}: {e}")
        
        # Parse Excel files
        xlsx_files = list(self.folder_path.glob("**/*.xlsx"))
        for xlsx_file in xlsx_files:
            # Skip output files
            if 'Checks' in xlsx_file.name or xlsx_file.parent.name == 'output':
                continue
                
            doc_type = self.identify_document_type(xlsx_file.name)
            
            if doc_type == 'CALCULO':
                parser = CalculoParser(str(xlsx_file))
                try:
                    data = parser.parse()
                    self.parsed_data['CALCULO'] = data
                    print(f"✓ Parsed CALCULO: {xlsx_file.name}")
                except Exception as e:
                    print(f"✗ Error parsing {xlsx_file.name}: {e}")
    
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
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'homeowner_phone')
        current_row += 1
        
        # Mail
        ws[f'B{current_row}'] = 'Mail'
        ws[f'C{current_row}'] = self._get_value('CONTRATO', 'homeowner_email')
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
        
        # Save
        wb.save(output_path)
        print(f"\n✅ Excel generated: {output_path}")
    
    def _get_value(self, doc_type: str, field: str) -> str:
        """Get value from parsed data"""
        if doc_type not in self.parsed_data:
            return ""
        
        value = self.parsed_data[doc_type].get(field, "")
        
        if value == "NOT FOUND":
            return ""
        
        return value
