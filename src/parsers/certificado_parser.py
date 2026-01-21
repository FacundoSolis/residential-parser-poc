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
            "document_type": "CERTIFICADO_INSTALADOR",
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
        """Extract surface in m2"""
        pattern = r"superficie tratada.*?([\d.]+)\s*m"
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            return f"{match.group(1)}"
        return "NOT FOUND"

    def _extract_climatic_zone(self) -> str:
        """Extract climatic zone"""
        pattern = r"zona climática.*?es.*?([A-E]\d)"
        match = re.search(pattern, self.text, re.IGNORECASE)
        if match:
            return match.group(1).upper()
        return "NOT FOUND"

    def _extract_calculation_methodology(self) -> str:
        """
        Certificado Instalador - Calculation methodology.
        Queremos:
        - un número tipo 0,70 / 0,83 (normalmente b o metodología)
        - o detectar R'T / R''T
        Evitar falsos positivos: 0,047 W/m2K, 0,1 m2K/W, Ui/Uf, etc.
        """
        if not self.text:
            return "NOT FOUND"

        t = (
            self.text.replace("’", "'")
            .replace("´", "'")
            .replace("`", "'")
            .replace("″", "''")
        )

        # 1) Detectar marcadores R'T / R''T
        has_r1t = bool(re.search(r"R\s*'\s*T", t, re.IGNORECASE))
        has_r2t = bool(re.search(r"R\s*''\s*T", t, re.IGNORECASE))
        has_media = bool(re.search(r"media\s+aritm[eé]tica", t, re.IGNORECASE))
        has_formula = bool(re.search(r"R\s*=\s*\(?\s*R\s*'\s*\+\s*R\s*''\s*\)?\s*/\s*2", t, re.IGNORECASE))

        if (has_media or has_formula) and (has_r1t and has_r2t):
            return "CTE_RT_MEDIA_ARITMETICA (R'T + R''T)/2"
        if has_r1t and has_r2t:
            return "R'T_R''T_PRESENT"
        if has_r2t:
            return "R''T_PRESENT"
        if has_r1t:
            return "R'T_PRESENT"

        # 2) Buscar número “metodología” con contexto BUENO (b / defecto / se establece)
        good_context = [
            r"(?:valor\s+de\s+b|se\s+establece\s+por\s+defecto\s+un\s+valor\s+de\s+b)[^\n]{0,120}?(\d+[.,]\d+)",
            r"(?:metodolog[ií]a|m[eé]todo)[^\n]{0,120}?(\d+[.,]\d+)",
        ]
        for p in good_context:
            m = re.search(p, t, re.IGNORECASE)
            if m:
                val = m.group(1).replace(".", ",")
                return val

        # 3) Fallback: primer decimal pequeño, PERO filtrando unidades típicas que NO son metodología
        # Captura 0,xx o 1,xx (máx 2 decimales)
        for m in re.finditer(r"\b([01][.,]\d{1,2})\b", t):
            val = m.group(1).replace(".", ",")
            start = max(0, m.start() - 25)
            end = min(len(t), m.end() + 25)
            ctx = t[start:end].lower()

            # contexto malo general
            if re.search(r"(w/mk|w/m2k|m2k/w|conductividad|λ|lambda|ui\b|uf\b|rtotal|resistencia)", ctx):
                continue

            # ✅ descarta 0,0x y 0,1 típicos de conductividad / resistencias aire
            if re.fullmatch(r"0,[0-1]\d", val) or val in {"0,1", "0,10"}:
                continue

            # ✅ refuerza contexto malo adicional (por si OCR mete 0,047 etc)
            if re.search(r"(w/mk|w/m2k|m2k/w|0,047|0\.047|0,1)", ctx):
                continue

            return val



        return "NOT FOUND"



