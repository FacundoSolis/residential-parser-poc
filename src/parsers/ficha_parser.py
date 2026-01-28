"""
Parser for FICHA RES020 documents
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any
import re


class FichaParser(BaseDocumentParser):
    """Parser for Ficha RES020 documents"""

    def _normalize(self, t: str) -> str:
        """Normalize text for better pattern matching"""
        if not t:
            return ""

        # Basic normalization for OCR text
        t = (
            t.replace("’", "'")
            .replace("´", "'")
            .replace("`", "'")
            .replace("″", "''")
        )

        # Common OCR errors in Spanish
        t = t.replace("Le6n", "León").replace("Direcci6n", "Dirección").replace("ubicaci6n", "ubicación")

        return t

    def parse(self) -> Dict[str, Any]:
        self.extract_text()
        t = self._normalize(self.text)
        # Extract lifespan and provide fallback from context (CERTIFICADO/CONTRATO)
        lifespan = self._extract_lifespan(t)
        if not lifespan:
            ctx = getattr(self, 'context_data', None)
            if ctx and isinstance(ctx, dict):
                # Prefer CERTIFICADO then CONTRATO
                cert = ctx.get('CERTIFICADO')
                if cert and isinstance(cert, dict):
                    lifespan = cert.get('lifespan') or lifespan
                if not lifespan:
                    contr = ctx.get('CONTRATO')
                    if contr and isinstance(contr, dict):
                        lifespan = contr.get('lifespan') or lifespan

        return {
            'document_type': 'FICHA',
            'homeowner_name': self._extract_homeowner_name(t),
            'homeowner_address': self._extract_homeowner_address(t),
            'homeowner_dni': self._extract_homeowner_dni(t),
            'act_code': self._extract_act_code(t),
            'catastral_ref': self._extract_catastral_ref(t),
            'energy_savings': self._extract_energy_savings(t),
            'start_date': self._extract_start_date(t),
            'finish_date': self._extract_finish_date(t),
            'sell_price': self._extract_sell_price(t),
            'fp': self._extract_fp(t),
            'ui': self._extract_ui(t),
            'uf': self._extract_uf(t),
            'surface': self._extract_surface(t),
            'climatic_zone': self._extract_climatic_zone(t),
            'isolation_thickness': self._extract_isolation_thickness(t),
            'lifespan': lifespan or "",
        }

    def _extract_homeowner_name(self, text: str) -> str:
        """Extract homeowner name from ficha"""
        patterns = [
            r'(?:propietario|titular|dueño)[:\s]*([^\n\r]{10,50})',
            r'(?:nombre|name)[:\s]*([^\n\r]{10,50})',
        ]
        name = self._find_first_pattern(text, patterns)
        if name and "BONO SOCIAL" not in name.upper() and "PERCEPTORES" not in name.upper():
            return name
        return "NOT FOUND"

    def _extract_homeowner_address(self, text: str) -> str:
        """Extract homeowner address from ficha"""
        patterns = [
            r'(?:domicilio|dirección|address)[:\s]*([^\n\r]{15,80})',
            r'(?:ubicación|location)[:\s]*([^\n\r]{15,80})',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_homeowner_dni(self, text: str) -> str:
        """Extract homeowner DNI from ficha"""
        patterns = [
            r'(?:dni|nie|documento)[:\s]*([0-9]{8}[A-Z]|[XYZ][0-9]{7}[A-Z])',
            r'(?:nif|cif)[:\s]*([0-9]{8}[A-Z]|[XYZ][0-9]{7}[A-Z])',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_act_code(self, text: str) -> str:
        """Extract ACT code from ficha - busca RES010 o RES020"""
        m = re.search(r"(RES0*\d{2,3})", text, re.IGNORECASE)
        if m:
            code = m.group(1).upper()
            code = re.sub(r"RES0+(\d)", r"RES0\1", code)  # normaliza RES00020 -> RES020
            return code
        
        patterns = [
            r'(?:código\s*act|act\s*code)[:\s]*(RES\s*0*\d{2,3})',
            r'(?:tipo\s*de\s*act|act\s*type)[:\s]*(RES\s*0*\d{2,3})',
            r'(?:medida|actuación)[:\s]*(RES\s*0*\d{2,3})',
        ]
        
        for pattern in patterns:
            m = re.search(pattern, text, re.IGNORECASE)
            if m:
                code = m.group(1).upper().replace(" ", "")
                code = re.sub(r"RES0+(\d)", r"RES0\1", code)
                return code
        
        return ""

    def _extract_catastral_ref(self, text: str) -> str:
        """Extract catastral reference from ficha"""
        patterns = [
            r'(?:referencia\s*catastral|catastral\s*ref)[:\s]*([0-9A-Z\s]{14,20})',
            r'(?:ref\s*catastral)[:\s]*([0-9A-Z\s]{14,20})',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_energy_savings(self, text: str) -> str:
        """Extract energy savings from ficha"""
        patterns = [
            r'(?:ahorro\s*energético|energy\s*savings)[:\s]*([0-9.,]+)',
            r'(?:kwh|kilowatts? hora)[:\s]*([0-9.,]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_sell_price(self, text: str) -> str:
        """Extract sell price (€/kWh) from ficha"""
        patterns = [
            r"precio\s+de\s+venta[^€\d]*([\d\s,]+(?:\.\d+)?)\s*€/[kI]Wh",
            r"venta[^€\d]*([\d\s,]+(?:\.\d+)?)\s*€/[kI]Wh",
            r"([\d\s,]+(?:\.\d+)?)\s*€/[kI]Wh",
            r"precio\s+de\s+venta[^€\d]*([\d\s,]+(?:\.\d+)?)\s*euros?\s*por\s*[kI]Wh",
        ]
        for pattern in patterns:
            m = re.search(pattern, text, re.IGNORECASE)
            if m:
                price = m.group(1).replace(",", ".").replace(" ", "")
                try:
                    float(price)  # validar que sea número
                    return price
                except ValueError:
                    continue
        return ""

    def _extract_start_date(self, text: str) -> str:
        """Extract start date from ficha"""
        patterns = [
            r"inici[oó]\s+el\s+(\d+)\s+de\s+(\w+)\s+de\s+(\d{4})",
            r"[Ff]echa.*?inicio[:\s]+([\d/]+)",
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                if len(match.groups()) == 3:
                    day, month, year = match.groups()
                    months = {
                        "enero": "01",
                        "febrero": "02",
                        "marzo": "03",
                        "abril": "04",
                        "mayo": "05",
                        "junio": "06",
                        "julio": "07",
                        "agosto": "08",
                        "septiembre": "09",
                        "octubre": "10",
                        "noviembre": "11",
                        "diciembre": "12",
                    }
                    month_num = months.get(month.lower(), month)
                    return f"{day}/{month_num}/{year}"
                else:
                    return match.group(1)

        return ""

    def _extract_finish_date(self, text: str) -> str:
        """Extract finish date from ficha"""
        patterns = [
            r"finaliz[oó].*?el\s+(\d+)\s+de\s+(\w+)\s+de\s+(\d{4})",
            r"[Ff]echa.*?fin[:\s]+([\d/]+)",
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                if len(match.groups()) == 3:
                    day, month, year = match.groups()
                    months = {
                        "enero": "01",
                        "febrero": "02",
                        "marzo": "03",
                        "abril": "04",
                        "mayo": "05",
                        "junio": "06",
                        "julio": "07",
                        "agosto": "08",
                        "septiembre": "09",
                        "octubre": "10",
                        "noviembre": "11",
                        "diciembre": "12",
                    }
                    month_num = months.get(month.lower(), month)
                    return f"{day}/{month_num}/{year}"
                else:
                    return match.group(1)

        return ""

    def _extract_fp(self, text: str) -> str:
        """Extract Fp from ficha"""
        patterns = [
            r'(?:fp|factor\s*p)[:=\s]*([0-9.,]+)',
            r'(?:p\s*factor)[:=\s]*([0-9.,]+)',
            r'Fp[:=\s]*([0-9.,]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_ui(self, text: str) -> str:
        """Extract Ui from ficha"""
        patterns = [
            r'(?:ui|transmitancia)[:=\s]*([0-9.,]+)',
            r'(?:u\s*inicial)[:=\s]*([0-9.,]+)',
            r'Ui[:=\s]*([0-9.,]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_uf(self, text: str) -> str:
        """Extract Uf from ficha"""
        patterns = [
            r'(?:uf|u\s*final)[:=\s]*([0-9.,]+)',
            r'(?:transmitancia\s*final)[:=\s]*([0-9.,]+)',
            r'Uf[:=\s]*([0-9.,]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_surface(self, text: str) -> str:
        """Extract surface area from ficha"""
        patterns = [
            r'(?:superficie|surface|área)[:=\s]*([0-9.,]+)',
            r'(?:s\s*=|superficie\s*=)[:=\s]*([0-9.,]+)',
            r'S[:=\s]*([0-9.,]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_climatic_zone(self, text: str) -> str:
        """Extract climatic zone from ficha"""
        patterns = [
            r'(?:zona\s*climática|climatic\s*zone)[:\s]*([A-Z0-9]+)',
            r'(?:zona\s*=|zone\s*=)[:\s]*([A-Z0-9]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_isolation_thickness(self, text: str) -> str:
        """Extract isolation thickness from ficha"""
        patterns = [
            r'(?:espesor\s*aislamiento|isolation\s*thickness)[:\s]*([0-9.,]+)',
            r'(?:grosor\s*aislante)[:\s]*([0-9.,]+)',
        ]
        return self._find_first_pattern(text, patterns)

    def _extract_energy_savings(self, text: str) -> str:
        """Extract energy savings from ficha"""
        m = re.search(r"(\d[\d\s.,]{2,})\s*k[wW]?[hH]\s*/\s*a(?:ñ|n)o", text, re.IGNORECASE)
        if m:
            raw = m.group(1)
            value = re.sub(r"[^\d]", "", raw)
            if not value:
                return ""
            try:
                num_value = int(value)
                if num_value < 500:
                    return ""
                if num_value > 100000:
                    num_value //= 100
                return str(num_value)
            except ValueError:
                return value
        return ""

    def _extract_lifespan(self, text: str) -> str:
        """Extract lifespan from ficha"""
        # Normalize common OCR/encoding issues for matching
        t = (text or "")
        # Replace common accented variants and OCR mistakes
        t_norm = t.replace("ú", "u").replace("ó", "o").replace("ñ", "n")
        t_norm = t_norm.replace("años", "anos").replace("a\xc3\xb1os", "anos")

        # Try explicit labels first: 'vida útil', 'duracion', 'lifespan'
        m = re.search(r"(?:vida\s*util|vida\s*útil|duraci[oó]n|duracion|lifespan)[:\s\-]*([0-9]{1,3})", t_norm, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            if val.isdigit() and 1 <= int(val) <= 100:
                return val

        # General fallback: any number followed by años/anos/years
        m2 = re.search(r"\b([0-9]{1,3})\s*(?:a[nn]os|anos|years)\b", t_norm, re.IGNORECASE)
        if m2:
            val = m2.group(1).strip()
            if val.isdigit() and 1 <= int(val) <= 100:
                return val

        # Last resort: any standalone small integer that looks like years near the word 'vida' or 'util'
        m3 = re.search(r"(vida|util|lifespan).{0,30}?([0-9]{1,3})", t_norm, re.IGNORECASE)
        if m3:
            val = m3.group(2).strip()
            if val.isdigit() and 1 <= int(val) <= 100:
                return val

        return ""

    def _find_first_pattern(self, text: str, patterns: list) -> str:
        """Find first matching pattern"""
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                result = match.group(1).strip()
                if len(result) > 3:  # Avoid garbage matches
                    return result
        return ""
