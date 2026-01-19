"""
Parser for CONTRATO CESION AHORROS documents

✅ Robust against:
- OCR noise (broken accents, random newlines)
- different contract templates (DE UNA PARTE / Y DE OTRA vs CESIONARIO / CEDENTE)
- garbage captures like "C" for address/location
- mis-encoded characters (rare case: '6' used as 'ó' in some PDFs)

Returns fields used in MatrixGenerator:
- sd_do_* (cesionario)
- homeowner_* (cedente)
- installer/location/utm/catastral/energy_savings/act_code
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any, Tuple
import re


class ContratoParser(BaseDocumentParser):
    """Parser for Contrato Cesión Ahorros"""

    def parse(self) -> Dict[str, Any]:
        """Extract all fields from Contrato"""
        self.extract_text()
        t = self._normalize(self.text)

        cesionario_txt, cedente_txt = self._split_sections(t)
        self._cesionario_txt = cesionario_txt
        self._cedente_txt = cedente_txt
        self._text_norm = t

        result = {
            "document_type": "CONTRATO_CESION_AHORROS",

            # Cesionario (empresa)
            "sd_do_company_name": self._extract_cesionario_company(),
            "sd_do_cif": self._extract_cesionario_cif(),
            "sd_do_address": self._extract_cesionario_address(),
            "sd_do_representative_name": self._extract_cesionario_representative(),
            "sd_do_representative_dni": self._extract_cesionario_dni(),
            "sd_do_signature": self._check_cesionario_signature(),

            # Cedente (homeowner)
            "homeowner_name": self._extract_cedente_name(),
            "homeowner_dni": self._extract_cedente_dni(),
            "homeowner_address": self._extract_cedente_address(),
            "homeowner_phone": self._extract_cedente_phone(),
            "homeowner_email": self._extract_cedente_email(),
            "homeowner_notifications": self._extract_notifications(),
            "homeowner_signatures": self._check_cedente_signature(),

            # Act / obra
            "installer": self._extract_installer(),
            "location": self._extract_location(),  # dirección de la actuación
            "catastral_ref": self._extract_catastral_ref(),
            "utm_coordinates": self._extract_utm_coordinates(),
            "energy_savings": self._extract_energy_savings(),
            "act_code": self._extract_act_code(),
        }

        return result

    # ------------------------
    # Normalization & helpers
    # ------------------------

    def _normalize(self, t: str) -> str:
        """Normalize OCR artifacts and whitespace."""
        if not t:
            return ""

        # Normalize punctuation / quotes
        t = (
            t.replace("’", "'")
            .replace("´", "'")
            .replace("`", "'")
            .replace("″", "''")
        )

        # OPTIONAL: only if you *really* see "Castilla y Le6n" etc.
        # If it's too aggressive, comment this out.
        t = t.replace("Le6n", "León").replace("Direcci6n", "Dirección").replace("ubicaci6n", "ubicación")

        # Collapse whitespace but preserve newlines somewhat
        t = re.sub(r"[ \t]+", " ", t)
        t = re.sub(r"\n{3,}", "\n\n", t)

        return t.strip()

    def _is_valid_text(self, s: str, min_len: int = 8) -> bool:
        if not s:
            return False
        s = s.strip()
        if not s or s == "NOT FOUND":
            return False
        if len(s) < min_len:
            return False
        # Reject typical garbage single-letter / short tokens
        if s.upper() in {"C", "CL", "CALLE", "AV", "AVD", "S/N"}:
            return False
        return True

    def _find_first(self, pattern: str, txt: str, flags=re.IGNORECASE | re.DOTALL) -> str:
        m = re.search(pattern, txt or "", flags)
        if not m:
            return "NOT FOUND"
        val = re.sub(r"\s+", " ", m.group(1)).strip()
        return val if self._is_valid_text(val, 3) else "NOT FOUND"

    def _find_dni(self, txt: str) -> str:
        """DNI: 8 digits + letter (OCR tolerant)."""
        if not txt:
            return "NOT FOUND"
        t = txt.upper()
        t = t.replace(" ", "").replace("-", "").replace(".", "").replace(":", "")
        t = t.replace("O", "0").replace("I", "1").replace("L", "1")
        m = re.search(r"(\d{8}[A-Z])", t)
        return m.group(1) if m else "NOT FOUND"

    def _find_cif(self, txt: str) -> str:
        """CIF (common): Letter + 8 digits."""
        if not txt:
            return "NOT FOUND"
        m = re.search(r"\b([A-Z]\d{8})\b", txt.upper())
        return m.group(1) if m else "NOT FOUND"

    def _split_sections(self, t: str) -> Tuple[str, str]:
        """
        Returns (cesionario_text, cedente_text)

        Supports multiple templates:
        - DE UNA PARTE ... Y DE OTRA ...
        - CESIONARIO ... CEDENTE ...
        - If not found -> ("", full_text)
        """
        if not t:
            return "", ""

        # Template A: DE UNA PARTE / Y DE OTRA
        m1 = re.search(r"\bDE\s+UNA\s+PARTE\b", t, re.IGNORECASE)
        m2 = re.search(r"\bY\s+DE\s+OTRA\b", t, re.IGNORECASE)
        if m1 and m2 and m2.start() > m1.start():
            return t[m1.start():m2.start()], t[m2.start():]

        # Template B: CESIONARIO / CEDENTE
        m3 = re.search(r"\bCESIONARIO\b", t, re.IGNORECASE)
        m4 = re.search(r"\bCEDENTE\b", t, re.IGNORECASE)
        if m3 and m4 and m4.start() > m3.start():
            return t[m3.start():m4.start()], t[m4.start():]

        # Fallback: keep all text as cedente (safer than mixing)
        return "", t

    # ------------------------
    # Extractors (ACT / common)
    # ------------------------

    def _extract_act_code(self) -> str:
        """Extract RES code (RES010, RES020)."""
        t = getattr(self, "_text_norm", "") or self.text
        m = re.search(r"(RES0*\d{2,3})", t, re.IGNORECASE)
        if m:
            code = m.group(1).upper()
            code = re.sub(r"RES0+(\d)", r"RES0\1", code)
            return code
        return "NOT FOUND"

    def _extract_energy_savings(self) -> str:
        """Extract energy savings in kWh/año (OCR tolerant)."""
        t = getattr(self, "_text_norm", "") or self.text

        # Prefer explicit kWh/año occurrences
        m = re.search(r"([\d.,]+)\s*kWh\s*/\s*a(?:ñ|n)o", t, re.IGNORECASE)
        if m:
            value = m.group(1).replace(",", ".").strip()
            try:
                num_value = float(value)
                if num_value < 500:
                    return "NOT FOUND"   # <--- antes era ""
                return value
            except ValueError:
                return value

        # fallback older pattern
        m = re.search(r"([\d.]+)\s*kWh/a[ñnrio]+(?:\s|,|\.)", t, re.IGNORECASE)
        if m:
            value = m.group(1).strip()
            try:
                num_value = float(value)
                if num_value < 500:
                    return "NOT FOUND"   # <--- antes era ""
                return value
            except ValueError:
                return value

        return "NOT FOUND"


    def _extract_catastral_ref(self) -> str:
        """Extract catastral reference."""
        t = getattr(self, "_text_norm", "") or self.text

        # Pattern 1: "Referencia catastral: XXXXX"
        m = re.search(r"Referencia\s+catastral\s*:\s*([0-9A-Z\s]+)", t, re.IGNORECASE)
        if m:
            ref = re.sub(r"\s+", "", m.group(1)).strip()
            if re.match(r"^[0-9A-Z]{10,25}$", ref):
                return ref

        # Pattern 2: OCR sometimes splits
        m = re.search(r"\b(\d{6,8}[A-Z]{1,3}\d{2,6}[A-Z]{1,4}\d{1,6}[A-Z]{1,4})\b", t)
        if m:
            return m.group(1)

        return "NOT FOUND"

    def _extract_utm_coordinates(self) -> str:
        """Extract UTM coordinates in various formats."""
        t = getattr(self, "_text_norm", "") or self.text

        # Format 1: "UTM 30, X:275624.89, Y:4741864.43"
        m = re.search(r"UTM\s+(\d+),?\s*X:\s*([\d.]+),?\s*Y:\s*([\d.\s]+)", t, re.IGNORECASE)
        if m:
            zone = m.group(1)
            x = m.group(2)
            y = m.group(3).replace(" ", "")
            return f"X:{x} Y:{y} HUSO:{zone}"

        # Format 2: "X: ... Y: ... HUSO:..."
        m = re.search(r"X:\s*([\d.]+)[^Y]*Y:\s*([\d.]+)", t, re.IGNORECASE)
        if m:
            x = m.group(1)
            y = m.group(2)
            zm = re.search(r"HUSO[:\s]*(\d+)", t, re.IGNORECASE)
            zone = zm.group(1) if zm else "30"
            return f"X:{x} Y:{y} HUSO:{zone}"

        return "NOT FOUND"

    def _extract_installer(self) -> str:
        txt = self._normalize(self.text)

        # Busca "Instalador: XXXXX" o "Instalador XXXXX"
        patterns = [
            r"Instalador[:\s]+([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s&\.-]{3,80})",
            r"El\s+Instalador[:\s]+([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s&\.-]{3,80})",
        ]
        for p in patterns:
            m = re.search(p, txt, re.IGNORECASE)
            if m:
                name = re.sub(r"\s+", " ", m.group(1)).strip()
                # Limpia basura típica OCR
                name = re.sub(r"\b(NIF|CIF)\b.*$", "", name, flags=re.IGNORECASE).strip()
                if len(name.split()) >= 2:
                    return name
        return "NOT FOUND"

    def _extract_location(self) -> str:
        """
        Extract location/address of installation (avoid garbage like 'C').
        """
        t = getattr(self, "_text_norm", "") or self.text

        # Pattern 1: "Dirección: ...."
        m = re.search(r"Direcci[oó]n\s*:\s*([^\n\.]{8,220})", t, re.IGNORECASE)
        if m:
            loc = re.sub(r"\s+", " ", m.group(1)).strip().rstrip(".")
            if self._is_valid_text(loc, 8):
                return loc

        # Pattern 2: "en CALLE/PLAZA ..."
        m = re.search(
            r"\ben\s+(.{8,200}?(?:PLAZA|CALLE|AVENIDA|CARRETERA|CL)\b.{0,80}?(?:\b\d{5}\b|Castilla y Le[oó]n))",
            t,
            re.IGNORECASE,
        )
        if m:
            loc = re.sub(r"\s+", " ", m.group(1)).strip()
            if self._is_valid_text(loc, 8):
                return loc

        # Pattern 3: "localidad de ..."
        m = re.search(r"localidad\s+de\s+(.{3,80}?)(?:\s+Castilla y Le[oó]n|\s+\b\d{5}\b)", t, re.IGNORECASE)
        if m:
            loc = re.sub(r"\s+", " ", m.group(1)).strip()
            if self._is_valid_text(loc, 8):
                return loc

        return "NOT FOUND"

    # ------------------------
    # Cesionario fields
    # ------------------------

    def _extract_cesionario_company(self) -> str:
        txt = getattr(self, "_cesionario_txt", "") or ""
        t = txt if txt else (getattr(self, "_text_norm", "") or self.text)

        # Typical: "DE UNA PARTE, <empresa>, con CIF ..."
        val = self._find_first(r"DE\s+UNA\s+PARTE[,\s]+(.+?)(?:,?\s*con\s+(?:CIF|NIF)|\s+(?:CIF|NIF)\b)", t)
        if val != "NOT FOUND":
            return val

        # Alternative: "CESIONARIO: <empresa> ..."
        val = self._find_first(r"\bCESIONARIO\b\s*[:\-]?\s*(.+?)(?:\s+(?:CIF|NIF)\b|,|\n)", t)
        if val != "NOT FOUND":
            return val

        return "NOT FOUND"

    def _extract_cesionario_cif(self) -> str:
        txt = getattr(self, "_cesionario_txt", "") or ""
        t = txt if txt else (getattr(self, "_text_norm", "") or self.text)
        return self._find_cif(t)

    def _extract_cesionario_address(self) -> str:
        txt = getattr(self, "_cesionario_txt", "") or ""
        t = txt if txt else (getattr(self, "_text_norm", "") or self.text)
        val = self._find_first(r"domicilio[^:]*:\s*(.+?)(?:\n|,?\s*CP|\s+C\.P\.)", t)
        return val if self._is_valid_text(val, 8) else "NOT FOUND"

    def _extract_cesionario_representative(self) -> str:
        txt = getattr(self, "_cesionario_txt", "") or ""
        t = txt if txt else (getattr(self, "_text_norm", "") or self.text)
        val = self._find_first(r"representad[oa]\s+por\s+D[./]?\s*([A-ZÑÁÉÍÓÚ][A-Za-zÑÁÉÍÓÚ\s]{5,60}?)(?:,|\s+con|\n)", t)
        return val

    def _extract_cesionario_dni(self) -> str:
        txt = getattr(self, "_cesionario_txt", "") or ""
        t = txt if txt else (getattr(self, "_text_norm", "") or self.text)
        return self._find_dni(t)

    def _check_cesionario_signature(self) -> str:
        txt = getattr(self, "_cesionario_txt", "") or ""
        t = txt if txt else (getattr(self, "_text_norm", "") or self.text)
        if re.search(r"Firma.*Cesionario|Cesionario.*Firma", t, re.IGNORECASE):
            return "Present"
        return "NOT FOUND"

    # ------------------------
    # Cedente fields
    # ------------------------

    def _extract_cedente_name(self) -> str:
        txt = getattr(self, "_cedente_txt", "") or self.text
        txt = self._normalize(txt)

        # Patrones típicos en contratos: "Y DE OTRA" / "CEDENTE"
        patterns = [
            r"Y\s+DE\s+OTRA\s+PARTE[:,\s]*D\.?\s*(?:ÑA|NA|N)?\.?\s*([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]{6,60}?)\s*,\s*mayor\s+de\s+edad",
            r"CEDENTE[:\s]*D\.?\s*(?:ÑA|NA|N)?\.?\s*([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]{6,60}?)\s*,\s*mayor\s+de\s+edad",
            r"D\.?\s*(?:ÑA|NA|N)?\.?\s*([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]{6,60}?)\s*,\s*mayor\s+de\s+edad,\s*con\s+DNI",
        ]

        bad_words = re.compile(
            r"\b(haya|debido|cantidad|instalador|por\s+cuanto|notificaciones|cl[aá]usula|presente)\b",
            re.IGNORECASE,
        )

        for p in patterns:
            m = re.search(p, txt, re.IGNORECASE | re.DOTALL)
            if m:
                name = re.sub(r"\s+", " ", m.group(1)).strip()
                # Heurística: nombre real suele ser 2-5 palabras
                if 2 <= len(name.split()) <= 6 and not bad_words.search(name):
                    return name

        return "NOT FOUND"


    def _extract_cedente_dni(self) -> str:
        txt = getattr(self, "_cedente_txt", "") or self.text
        txt = self._normalize(txt)

        # 1) súper prioridad: DNI en frase típica
        m = re.search(
            r"mayor\s+de\s+edad,\s*con\s+DNI[:\s]*([0-9OIlL]{7,8}\s*[A-Z])",
            txt,
            re.IGNORECASE,
        )
        if m:
            cand = m.group(1).upper().replace(" ", "")
            cand = cand.replace("O", "0").replace("I", "1").replace("L", "1")
            if self._is_valid_spanish_dni(cand):
                return cand

        # 2) DNI cerca de palabra DNI
        m = re.search(r"\bDNI[:\s]*([0-9OIlL]{7,8}\s*[A-Z])\b", txt, re.IGNORECASE)
        if m:
            cand = m.group(1).upper().replace(" ", "")
            cand = cand.replace("O", "0").replace("I", "1").replace("L", "1")
            if self._is_valid_spanish_dni(cand):
                return cand

        # 3) fallback global
        cand = self._find_dni(txt)
        if cand != "NOT FOUND" and self._is_valid_spanish_dni(cand):
            return cand

        # si no pasa checksum -> NO LO DEVUELVAS
        return "NOT FOUND"


    def _extract_cedente_address(self) -> str:
        """Best effort: location is usually the address of the action."""
        loc = self._extract_location()
        return loc if self._is_valid_text(loc, 8) else "NOT FOUND"

    def _extract_cedente_phone(self) -> str:
        txt = getattr(self, "_cedente_txt", "") or ""
        t = txt if txt else (getattr(self, "_text_norm", "") or self.text)

        # Near keyword
        m = re.search(r"tel[eé]fono[:\s]*([+\d][\d\s-]{8,})", t, re.IGNORECASE)
        if m:
            phone = re.sub(r"[^\d+]", "", m.group(1))
            if len(re.sub(r"\D", "", phone)) >= 9:
                return phone

        # fallback: Spanish 9-digit
        m = re.search(r"\b(\+34)?\s*(6\d{8}|7\d{8}|8\d{8}|9\d{8})\b", t)
        if m:
            return re.sub(r"\s+", "", m.group(0)).replace(" ", "")

        return "NOT FOUND"

    def _extract_cedente_email(self) -> str:
        txt = getattr(self, "_cedente_txt", "") or ""
        t = txt if txt else (getattr(self, "_text_norm", "") or self.text)
        m = re.search(r"[\w\.-]+@[\w\.-]+\.\w+", t)
        return m.group(0) if m else "NOT FOUND"

    def _extract_notifications(self) -> str:
        txt = getattr(self, "_cedente_txt", "") or ""
        t = txt if txt else (getattr(self, "_text_norm", "") or self.text)
        m = re.search(r"[Nn]otificaciones[:\s]+(.+?)(?:\n|$)", t, re.IGNORECASE)
        if m:
            val = re.sub(r"\s+", " ", m.group(1)).strip()
            return val if self._is_valid_text(val, 8) else "NOT FOUND"
        return "NOT FOUND"

    def _check_cedente_signature(self) -> str:
        txt = getattr(self, "_cedente_txt", "") or ""
        t = txt if txt else (getattr(self, "_text_norm", "") or self.text)
        signature_count = len(re.findall(r"Firma|Firmado|Fdo\.", t, re.IGNORECASE))
        return f"{signature_count} signature(s) found" if signature_count > 0 else "NOT FOUND"

    def _is_valid_spanish_dni(self, dni: str) -> bool:
        dni = (dni or "").strip().upper().replace(" ", "")
        if not re.match(r"^\d{8}[A-Z]$", dni):
            return False
        letters = "TRWAGMYFPDXBNJZSQVHLCKE"
        num = int(dni[:8])
        return dni[-1] == letters[num % 23]


    
