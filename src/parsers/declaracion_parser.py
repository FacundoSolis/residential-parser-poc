"""
Parser for DECLARACION RESPONSABLE documents
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any
import re


class DeclaracionParser(BaseDocumentParser):
    """Parser for Declaración Responsable"""

    def parse(self) -> Dict[str, Any]:
        self.extract_text()

        return {
            "document_type": "DECLARACION_RESPONSABLE",
            "homeowner_name": self._extract_homeowner_name(),
            "homeowner_dni": self._extract_homeowner_dni(),
            "homeowner_address": self._extract_homeowner_address(),
            "catastral_ref": self._extract_catastral_ref(),
            "act_code": self._extract_act_code(),
            "signature": self._check_signature(),
        }

    # ------------------------
    # DNI helpers (OCR tolerant)
    # ------------------------

    def _is_valid_spanish_dni(self, dni: str) -> bool:
        dni = (dni or "").strip().upper().replace(" ", "")
        if not re.match(r"^\d{8}[A-Z]$", dni):
            return False
        letters = "TRWAGMYFPDXBNJZSQVHLCKE"
        num = int(dni[:8])
        return dni[-1] == letters[num % 23]

    def _normalize_dni_parts(self, raw_num: str, raw_last: str) -> str:
        """
        Normaliza SOLO la parte numérica (8 dígitos) tolerando OCR O/I/L.
        La letra final NO se convierte a 1.
        También tolera caso típico OCR: '6' => 'G' como letra final.
        """
        if not raw_num or not raw_last:
            return ""

        num = raw_num.upper().replace(" ", "")
        num = num.replace("O", "0").replace("I", "1").replace("L", "1")

        last = raw_last.upper().replace(" ", "")
        last_candidates = []

        if last == "6":
            last_candidates = ["G"]
        elif last in {"1", "I", "L"}:
            # a veces OCR confunde la letra final con I/L/1
            last_candidates = ["I", "L"]
        else:
            last_candidates = [last]

        for L in last_candidates:
            cand = f"{num}{L}"
            if self._is_valid_spanish_dni(cand):
                return cand

        return ""

    def _find_any_valid_dni(self, t: str) -> str:
        """
        Fallback global: busca cualquier DNI válido en todo el texto.
        Soporta OCR en dígitos (O/I/L) y letra final (6->G).
        """
        if not t:
            return "NOT FOUND"

        txt = (t or "").upper()

        # buscamos 8 "dígitos" tolerantes + 1 letra tolerante
        for m in re.finditer(r"\b([0-9OIL]{8})\s*([A-Z6IL1])\b", txt):
            dni = self._normalize_dni_parts(m.group(1), m.group(2))
            if dni:
                return dni

        # fallback clásico por si el OCR vino limpio
        for m in re.finditer(r"\b(\d{8})([A-Z])\b", txt):
            cand = f"{m.group(1)}{m.group(2)}"
            if self._is_valid_spanish_dni(cand):
                return cand

        return "NOT FOUND"

    # ------------------------
    # Extractors
    # ------------------------

    def _extract_homeowner_name(self) -> str:
        t = self.text or ""

        # 1) Línea NOMBRE + DNI (prioridad) (mayúsculas típicas OCR)
        m = re.search(
            r"\n([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]{10,80})\s+([0-9OIL]{8}\s*[A-Z6IL1])\b",
            t.upper(),
            re.IGNORECASE
        )
        if m:
            name = re.sub(r"\s+", " ", m.group(1)).strip()
            if 2 <= len(name.split()) <= 8:
                return name

        # 2) Titular / Solicitante / Beneficiario
        m = re.search(
            r"(?:TITULAR|SOLICITANTE|BENEFICIARIO)\s*[:\-]?\s*([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]{10,80})",
            t.upper(),
            re.IGNORECASE,
        )
        if m:
            name = re.sub(r"\s+", " ", m.group(1)).strip()
            if 2 <= len(name.split()) <= 8:
                return name

        # 3) Don/Doña/D.
        m = re.search(
            r"\bD(?:ON|OÑA|\.)\s+([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]{10,80})\b",
            t.upper(),
            re.IGNORECASE,
        )
        if m:
            name = re.sub(r"\s+", " ", m.group(1)).strip()
            if 2 <= len(name.split()) <= 10:
                return name

        return "NOT FOUND"

    def _extract_homeowner_dni(self) -> str:
        t = self.text or ""
        up = t.upper()

        # 1) DNI en línea NOMBRE + DNI
        m = re.search(
            r"\n([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]{10,80})\s+([0-9OIL]{8})\s*([A-Z6IL1])\b",
            up,
            re.IGNORECASE
        )
        if m:
            dni = self._normalize_dni_parts(m.group(2), m.group(3))
            if dni:
                return dni

        # 2) Cerca de "NIF/NIE"
        m = re.search(
            r"NIF\s*/\s*NIE\s*[:\-]?\s*([0-9OIL]{8})\s*([A-Z6IL1])\b",
            up,
            re.IGNORECASE,
        )
        if m:
            dni = self._normalize_dni_parts(m.group(1), m.group(2))
            if dni:
                return dni

        # 3) Fallback global (lo más útil en OCR raro)
        return self._find_any_valid_dni(t)

    def _extract_homeowner_address(self) -> str:
        t = self.text or ""

        def is_garbage(val: str) -> bool:
            v = (val or "").strip()
            if not v:
                return True
            v_low = v.lower()
            if v_low in {"teléfono", "telefono", "tel", "telf", "tfno", "firma", "firmado"}:
                return True
            # basura corta tipo "C" / "CL"
            if v.strip().upper() in {"C", "CL", "CALLE", "AV", "AVD", "S/N"}:
                return True
            return False

        patterns = [
            # 1) Domicilio: ... (hasta CP / salto / Teléfono / Firma)
            r"[Dd]omicilio[:\s]+(.+?)(?=\n|,\s*\d{5}|CP|C\.P\.|Tel[eé]fono|Firma|Firmado)",
            # 2) Dirección: ...
            r"[Dd]irecci[oó]n[:\s]+(.+?)(?=\n|CP|C\.P\.|Tel[eé]fono|Firma|Firmado)",
            # 3) Si viene en la siguiente línea (OCR típico)
            r"(?:[Dd]omicilio|[Dd]irecci[oó]n)\s*[:\-]?\s*\n\s*(.{10,220})",
        ]

        for pattern in patterns:
            m = re.search(pattern, t, re.IGNORECASE | re.DOTALL)
            if m:
                val = re.sub(r"\s+", " ", m.group(1)).strip().rstrip(",;.")
                val = self._clean_address(val)
                if not is_garbage(val):
                    return val

        return "NOT FOUND"

    def _extract_catastral_ref(self) -> str:
        t = self.text or ""

        m = re.search(
            r"Referencia\s+catastral.*?\n\s*([0-9A-Z\s]{10,40})",
            t,
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            ref = re.sub(r"\s+", "", m.group(1)).strip().upper()
            if re.match(r"^(CL|C|CALLE|AV|AVENIDA|PL|PLAZA)", ref):
                return "NOT FOUND"
            if re.search(r"[A-Z]", ref) and re.search(r"\d", ref) and 10 <= len(ref) <= 25:
                return ref

        m = re.search(
            r"\b(\d{6,8}[A-Z]{1,3}\d{2,6}[A-Z]{1,4}\d{1,6}[A-Z]{1,4})\b",
            t,
        )
        if m:
            return m.group(1).upper()

        return "NOT FOUND"

    def _extract_act_code(self) -> str:
        t = self.text or ""
        m = re.search(r"(RES0*\d{2,3})", t, re.IGNORECASE)
        if m:
            code = m.group(1).upper()
            code = re.sub(r"RES0+(\d)", r"RES0\1", code)
            return code
        return "NOT FOUND"

    def _check_signature(self) -> str:
        t = self.text or ""
        return "Present" if re.search(r"Firma|Firmado|Fdo\.", t, re.IGNORECASE) else "NOT FOUND"

    def _clean_address(self, addr: str) -> str:
        if not addr:
            return addr

        a = re.sub(r"\s+", " ", addr).strip()

        # Quita prefijos típicos de plantilla/OCR
        a = re.sub(
            r"^(?:domicilio|direcci[oó]n|direccion|postal|c[oó]digo\s*postal|postal de la instalaci[oó]n.*?)(?:\s+en\s+que\s+se\s+ejecut[oó].*?)?\s*[:\-]?\s*",
            "",
            a,
            flags=re.IGNORECASE,
        )

        # Si metió separadores tipo "|" o "—", quédate con el lado más “dirección”
        if "|" in a:
            parts = [p.strip() for p in a.split("|") if p.strip()]
            # normalmente la dirección es la parte más corta y con números
            parts = sorted(parts, key=len)
            for p in parts:
                if re.search(r"\d", p):
                    a = p
                    break
            else:
                a = parts[-1]

        # Limpieza final
        a = a.strip(" ,;.-")
        return a

