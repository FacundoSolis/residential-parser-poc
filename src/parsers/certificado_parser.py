"""
Parser for CERTIFICADO INSTALADOR documents
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any
import re


class CertificadoParser(BaseDocumentParser):
    """Parser for Certificado Instalador"""

    def parse(self) -> Dict[str, Any]:
        """Extract all fields from Certificado"""
        self.extract_text()

        result = {
            "document_type": "CERTIFICADO",
            "act_code": self._extract_act_code(),
            "energy_savings": self._extract_energy_savings(),
            "start_date": self._extract_start_date(),
            "finish_date": self._extract_finish_date(),
            "address": self._extract_property_address(),
            "catastral_ref": self._extract_catastral_ref(),
            "lifespan": self._extract_lifespan(),
            "surface": self._extract_surface(),
            "climatic_zone": self._extract_climatic_zone(),
            "calculation_methodology": self._extract_calculation_methodology(),
            "fp": self._extract_fp(),
            "ui": self._extract_ui(),
            "uf": self._extract_uf(),
            "g": self._extract_g(),
            "b": self._extract_b(),
            "isolation_thickness": self._extract_isolation_thickness(),
            "isolation_type": self._extract_isolation_type(),

        }

        return result

    def _extract_act_code(self) -> str:
        """Extract RES code"""
        match = re.search(r"(RES0*\d{2,3})", self.text, re.IGNORECASE)
        if match:
            code = match.group(1).upper()
            code = re.sub(r"RES0+(\d)", r"RES0\1", code)
            return code
        return "NOT FOUND"

    def _extract_energy_savings(self) -> str:
        """Extract energy savings"""
        patterns = [
            r"Ahorro.*?AE\s+([\d.]+)",
            r"AE\s+([\d.]+)",
        ]

        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                return f"{match.group(1)}"
        return "NOT FOUND"

    def _extract_start_date(self) -> str:
        """Extract start date"""
        patterns = [
            r"inici[oó]\s+el\s+(\d+)\s+de\s+(\w+)\s+de\s+(\d{4})",
            r"[Ff]echa.*?inicio[:\s]+([\d/]+)",
        ]

        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
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
                    month_num = months.get(month.lower(), "00")
                    return f"{day}/{month_num}/{year}"
                else:
                    return match.group(1)

        return "NOT FOUND"

    def _extract_finish_date(self) -> str:
        r"""Extract finish date - handle newlines with \s+"""

        patterns = [
            r"finaliz[oó].*?el\s+(\d+)\s+de\s+(\w+)\s+de\s+(\d{4})",
            r"[Ff]echa.*?fin[:\s]+([\d/]+)",
        ]

        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE | re.DOTALL)
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
                    month_num = months.get(month.lower(), "00")
                    return f"{day}/{month_num}/{year}"
                else:
                    return match.group(1)

        return "NOT FOUND"

    def _extract_property_address(self) -> str:
        """Extract property address"""
        pattern = r"situado en\s+(.+?)(?:\.|La referencia|$)"
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            address = match.group(1).strip()
            address = re.sub(r"\s+", " ", address)
            return address
        return "NOT FOUND"

    def _extract_catastral_ref(self) -> str:
        """Extract catastral reference"""
        pattern = r"referencia catastral[^.]*es\s+([0-9A-Z]+)"
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            catastral = match.group(1).strip()
            catastral = re.sub(r"\s+", "", catastral)
            if re.match(r"^[0-9]+[A-Z]+[0-9]+[A-Z]+[0-9]+[A-Z]+$", catastral):
                return catastral
        return "NOT FOUND"

    def _extract_lifespan(self) -> str:
        """Extract lifespan in years"""
        pattern = r"duración.*?actuación.*?(\d+)\s*años"
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            return f"{match.group(1)}"
        return "NOT FOUND"

    def _extract_surface(self) -> str:
        t = self.text or ""
        m = re.search(r"superficie de la envolvente térmica final es de\s+([0-9]+(?:[.,]\d+)?)", t, re.IGNORECASE)
        if m:
            return m.group(1).replace(".", ",")
        return "NOT FOUND"

    def _extract_climatic_zone(self) -> str:
        """Extract climatic zone"""
        pattern = r"zona climática.*?es.*?([A-E]\d)"
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            return match.group(1).upper()
        return "NOT FOUND"

    def _extract_calculation_methodology(self) -> str:
        t = self.text or ""
        if "resistencia térmica" in t.lower():
            return "R t"
        return "NOT FOUND"
    def _extract_table_value(self, key: str) -> str:
        """
        Busca líneas tipo:
        Fp  1
        Ui: 3,26
        Uf = 0.23
        """
        if not self.text:
            return "NOT FOUND"

        t = self.text.replace("O", "0")  # OCR típico
        # ^\s*Fp\s*[:=]?\s*(valor)\s*$
        m = re.search(
            rf"(?im)^\s*{re.escape(key)}\s*[:=]?\s*([0-9]+(?:[.,]\d+)?)\s*$",
            t
        )
        if not m:
            return "NOT FOUND"

        return m.group(1).replace(".", ",")

    def _extract_fp(self) -> str:
        t = self.text or ""
        m = re.search(r"\bF\s*P\b\s*[:=]?\s*([0-9]+(?:[.,]\d+)?)", t, re.IGNORECASE)
        if not m:
            m = re.search(r"\bFp\b\s*[:=]?\s*([0-9]+(?:[.,]\d+)?)", t, re.IGNORECASE)
        return m.group(1).replace(".", ",") if m else "NOT FOUND"

    def _extract_ui(self) -> str:
        t = self.text or ""
        m = re.search(r"\bU\s*I\b\s*[:=]?\s*([0-9]+(?:[.,]\d+)?)", t, re.IGNORECASE)
        if not m:
            m = re.search(r"\bUi\b\s*[:=]?\s*([0-9]+(?:[.,]\d+)?)", t, re.IGNORECASE)
        return m.group(1).replace(".", ",") if m else "NOT FOUND"

    def _extract_uf(self) -> str:
        t = self.text or ""
        m = re.search(r"\bU\s*F\b\s*[:=]?\s*([0-9]+(?:[.,]\d+)?)", t, re.IGNORECASE)
        if not m:
            m = re.search(r"\bUf\b\s*[:=]?\s*([0-9]+(?:[.,]\d+)?)", t, re.IGNORECASE)
        return m.group(1).replace(".", ",") if m else "NOT FOUND"

    def _extract_g(self) -> str:
        t = self.text or ""
        # G o Gj o G j
        m = re.search(r"\bG\s*j?\b\s*[:=]?\s*([0-9]+(?:[.,]\d+)?)", t, re.IGNORECASE)
        return m.group(1).replace(".", ",") if m else "NOT FOUND"

    def _extract_b(self) -> str:
        t = self.text or ""
        m = re.search(r'valor\s+de\s+b\s+de\s+([0-9]+(?:[.,]\d+)?)', t, re.IGNORECASE)
        if m:
            return m.group(1).replace(".", ",")
        # Fallback
        return "0,70"


    def _extract_isolation_thickness(self) -> str:
        t = self.text or ""

        # 1) mm directo
        m = re.search(r"(?:espesor|aislamiento|thickness|e\s*=)[^0-9]{0,40}(\d{2,4})\s*mm", t, re.IGNORECASE)
        if m:
            return f"{m.group(1)} mm"

        # 2) cm (tu caso: "Cubierta:16 cm de Aislamiento ...")
        m = re.search(r"(?:cubierta|fachada|muro|cerramiento|aislamiento|espesor)[^\n]{0,80}?(\d+(?:[.,]\d+)?)\s*cm", t, re.IGNORECASE)
        if m:
            cm = float(m.group(1).replace(",", "."))
            mm = int(round(cm * 10))
            return f"{mm} mm"

        # 3) fallback bruto (último recurso)
        m = re.search(r"\b(\d{2,4})\s*mm\b", t, re.IGNORECASE)
        return f"{m.group(1)} mm" if m else "NOT FOUND"
    
    def _extract_isolation_type(self) -> str:
        t = self.text or ""
        # Look for "tipo Soplado" or "tipo Rollo"
        m = re.search(r'tipo\s+(Soplado|Rollo)', t, re.IGNORECASE)
        if m:
            return m.group(1).upper()
        # Fallback: if URSA, assume SOPLADO
        if 'URSA' in t.upper():
            return 'SOPLADO'
        return "NOT FOUND"







