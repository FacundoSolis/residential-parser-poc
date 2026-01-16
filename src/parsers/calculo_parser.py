"""
Parser for CALCULO.xlsx files
"""

import openpyxl
from pathlib import Path
from typing import Dict, Any


class CalculoParser:
    """Parser for Calculo Excel files"""
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.data = {}
        
    def parse(self) -> Dict[str, Any]:
        """Extract calculation data from Excel"""
        if not Path(self.file_path).exists():
            return {
                'document_type': 'CALCULO',
                'error': 'File not found'
            }
        
        try:
            wb = openpyxl.load_workbook(self.file_path)
            ws = wb.active
            
            result = {
                'document_type': 'CALCULO',
                'client_name': self._extract_client_name(ws),
                'area_total': self._extract_area_total(ws),
                'area_afect': self._extract_area_afect(ws),
                'porcentaje': self._extract_porcentaje(ws),
                'act_code': self._extract_act_code(ws),
                'fp': self._extract_fp(ws),
                'ki': self._extract_ki(ws),
                'kf': self._extract_kf(ws),
                's': self._extract_s(ws),
                'zone_climatique': self._extract_zone(ws),
                'ae': self._extract_ae(ws),
            }
            
            return result
        except Exception as e:
            return {
                'document_type': 'CALCULO',
                'error': str(e)
            }
    
    def _extract_client_name(self, ws) -> str:
        """Extract client name from Row 1"""
        try:
            cell = ws['A1'].value
            if cell and 'Client :' in str(cell):
                parts = str(cell).split('Client :')
                if len(parts) > 1:
                    name = parts[1].split('(')[0].strip()
                    return name
        except:
            pass
        return "NOT FOUND"
    
    def _extract_area_total(self, ws) -> str:
        """Extract total area"""
        try:
            for row in ws.iter_rows(min_row=3, max_row=5, values_only=True):
                if len(row) > 15 and row[14] == 'Área total' and row[15]:
                    return str(row[15])
        except:
            pass
        return "NOT FOUND"
    
    def _extract_area_afect(self, ws) -> str:
        """Extract affected area"""
        try:
            for row in ws.iter_rows(min_row=3, max_row=5, values_only=True):
                if len(row) > 15 and row[14] == 'Área afect' and row[15]:
                    return str(row[15])
        except:
            pass
        return "NOT FOUND"
    
    def _extract_porcentaje(self, ws) -> str:
        """Extract percentage"""
        try:
            for row in ws.iter_rows(min_row=5, max_row=6, values_only=True):
                if len(row) > 15 and row[14] == 'Porcentaje' and row[15]:
                    return f"{row[15]}%"
        except:
            pass
        return "NOT FOUND"
    
    def _extract_act_code(self, ws) -> str:
        """Extract RES code - look for RES020 in column O (index 14)"""
        try:
            for row in ws.iter_rows(min_row=10, max_row=15, values_only=True):
                if len(row) > 14 and row[14] and 'RES' in str(row[14]):
                    return str(row[14])
        except:
            pass
        return "NOT FOUND"
    
    def _extract_fp(self, ws) -> str:
        """Extract Fp value - same row as RES020, column P (index 15)"""
        try:
            for row in ws.iter_rows(min_row=10, max_row=15, values_only=True):
                if len(row) > 15 and row[14] and 'RES' in str(row[14]):
                    return str(row[15])
        except:
            pass
        return "NOT FOUND"
    
    def _extract_ki(self, ws) -> str:
        """Extract Ki value - same row as RES020, column Q (index 16)"""
        try:
            for row in ws.iter_rows(min_row=10, max_row=15, values_only=True):
                if len(row) > 16 and row[14] and 'RES' in str(row[14]):
                    return str(row[16])
        except:
            pass
        return "NOT FOUND"
    
    def _extract_kf(self, ws) -> str:
        """Extract Kf value - same row as RES020, column R (index 17)"""
        try:
            for row in ws.iter_rows(min_row=10, max_row=15, values_only=True):
                if len(row) > 17 and row[14] and 'RES' in str(row[14]):
                    return str(row[17])
        except:
            pass
        return "NOT FOUND"
    
    def _extract_s(self, ws) -> str:
        """Extract S (surface) value - same row as RES020, column S (index 18)"""
        try:
            for row in ws.iter_rows(min_row=10, max_row=15, values_only=True):
                if len(row) > 18 and row[14] and 'RES' in str(row[14]):
                    return str(row[18])
        except:
            pass
        return "NOT FOUND"
    
    def _extract_zone(self, ws) -> str:
        """Extract climatic zone - same row as RES020, column T (index 19)"""
        try:
            for row in ws.iter_rows(min_row=10, max_row=15, values_only=True):
                if len(row) > 19 and row[14] and 'RES' in str(row[14]):
                    return str(row[19])
        except:
            pass
        return "NOT FOUND"
    
    def _extract_ae(self, ws) -> str:
        """Extract AE (energy savings) - same row as RES020, column U (index 20)"""
        try:
            for row in ws.iter_rows(min_row=10, max_row=15, values_only=True):
                if len(row) > 20 and row[14] and 'RES' in str(row[14]):
                    return str(row[20])
        except:
            pass
        return "NOT FOUND"
