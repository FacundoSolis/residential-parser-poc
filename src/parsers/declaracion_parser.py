"""
Parser for DECLARACION RESPONSABLE documents
"""

from .base_parser import BaseDocumentParser
from typing import Dict, Any
import re


class DeclaracionParser(BaseDocumentParser):
    """Parser for Declaración Responsable"""

    def parse(self) -> Dict[str, Any]:
        """Extract all fields from Declaracion"""
        self.extract_text()

        result = {
            "document_type": "DECLARACION_RESPONSABLE",
            "homeowner_name": self._extract_homeowner_name(),
            "homeowner_dni": self._extract_homeowner_dni(),
            "homeowner_address": self._extract_homeowner_address(),
            "catastral_ref": self._extract_catastral_ref(),
            "act_code": self._extract_act_code(),
            "signature": self._check_signature(),
        }

        return result

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

    def _normalize_dni_candidate(self, raw: str) -> str:
        if not raw:
            return ""
        cand = raw.upper().replace(" ", "").replace("-", "").replace(".", "").replace(":", "")
        cand = cand.replace("O", "0").replace("I", "1").replace("L", "1")

        # Caso típico OCR: última letra "G" -> "6"
        if len(cand) == 9 and cand[-1] == "6":
            cand_g = cand[:-1] + "G"
            if self._is_valid_spanish_dni(cand_g):
                return cand_g

        if self._is_valid_spanish_dni(cand):
            return cand

        return ""

    # ------------------------
    # Extractors
    # ------------------------

    def _extract_homeowner_name(self) -> str:
        """
        En tus declaraciones el OCR saca una línea tipo:
        'DEL CUETO DEL POZO MANOLO 711094496'
        Vamos a sacar el nombre de ahí (más robusto que patrones "D./titular").
        """
        t = self.text or ""

        # 1) Línea NOMBRE + DNI (prioridad)
        m = re.search(
            r"\n([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]{10,80})\s+([0-9OIlL]{7,8}\s*[A-Z6])\b",
            t,
            re.IGNORECASE
        )
        if m:
            name = re.sub(r"\s+", " ", m.group(1)).strip()
            if 2 <= len(name.split()) <= 8:
                return name

        # 2) Fallback: tus patrones antiguos (por si viene distinto)
        patterns = [
            r"D[./]?\s*([A-ZÑ][a-zñ]+\s+[A-ZÑ][a-zñ]+(?:\s+[A-ZÑ][a-zñ]+)?)",
            r"titular[:\s]+([A-ZÑ][a-zñ]+\s+[A-ZÑ][a-zñ]+)",
        ]
        for pattern in patterns:
            match = re.search(pattern, t)
            if match:
                return match.group(1).strip()

        return "NOT FOUND"

    def _extract_homeowner_dni(self) -> str:
        """
        Maneja OCR donde la letra final puede venir como "6".
        Ej: 711094496 -> 71109449G (si checksum cuadra)
        """
        t = self.text or ""

        # 1) DNI en línea NOMBRE + DNI
        m = re.search(
            r"\n([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]{10,80})\s+([0-9OIlL]{7,8}\s*[A-Z6])\b",
            t,
            re.IGNORECASE
        )
        if m:
            dni = self._normalize_dni_candidate(m.group(2))
            if dni:
                return dni

        # 2) Cerca de "NIF/NIE"
        m = re.search(r"NIF\s*/\s*NIE\s*([0-9OIlL]{7,8}\s*[A-Z6])", t, re.IGNORECASE)
        if m:
            dni = self._normalize_dni_candidate(m.group(1))
            if dni:
                return dni

        # 3) Fallback: busca cualquier candidato 8 dígitos + letra/6
        m = re.search(r"\b(\d{8}[A-Z6])\b", t.upper().replace(" ", ""))
        if m:
            dni = self._normalize_dni_candidate(m.group(1))
            if dni:
                return dni

        return "NOT FOUND"

    def _extract_homeowner_address(self) -> str:
        """Extract address"""
        t = self.text or ""
        patterns = [
            r"[Dd]omicilio[:\s]+(.+?)(?:\n|,\s*\d{5})",
            r"[Dd]irecci[oó]n[:\s]+(.+?)(?:\n|CP)",
        ]

        for pattern in patterns:
            match = re.search(pattern, t, re.IGNORECASE)
            if match:
                return match.group(1).strip()

        return "NOT FOUND"

    def _extract_catastral_ref(self) -> str:
        t = self.text or ""

        # 1) Busca algo tipo referencia catastral cerca del label
        m = re.search(
            r"Referencia\s+catastral.*?\n\s*([0-9A-Z\s]{10,40})",
            t,
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            ref = re.sub(r"\s+", "", m.group(1)).strip().upper()

            # ❌ descarta direcciones típicas OCR (CL/C/CALLE/AV/etc.)
            if re.match(r"^(CL|C|CALLE|AV|AVENIDA|PL|PLAZA)", ref):
                return "NOT FOUND"

            # ✅ debe tener mezcla real de letras y números
            if re.search(r"[A-Z]", ref) and re.search(r"\d", ref) and 10 <= len(ref) <= 25:
                return ref

        # 2) Fallback: patrón típico RC española (alfanumérico largo con mezcla)
        m = re.search(
            r"\b(\d{6,8}[A-Z]{1,3}\d{2,6}[A-Z]{1,4}\d{1,6}[A-Z]{1,4})\b",
            t,
        )
        if m:
            return m.group(1).upper()

        return "NOT FOUND"



    def _extract_act_code(self) -> str:
        """Extract RES code"""
        t = self.text or ""
        match = re.search(r"(RES0*\d{2,3})", t, re.IGNORECASE)
        if match:
            code = match.group(1).upper()
            code = re.sub(r"RES0+(\d)", r"RES0\1", code)
            return code

        return "NOT FOUND"

    def _check_signature(self) -> str:
        """Check if signature exists"""
        t = self.text or ""
        if re.search(r"Firma|Firmado|Fdo\.", t, re.IGNORECASE):
            return "Present"

        return "NOT FOUND"
