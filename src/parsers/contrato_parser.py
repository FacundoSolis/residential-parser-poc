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
        self.extract_text()
        t = self._normalize(self.text)

        cesionario_txt, cedente_txt = self._split_sections(t)
        self._cesionario_txt = cesionario_txt
        self._cedente_txt = cedente_txt
        self._text_norm = t

        location = self._extract_location()
        if location == "NOT FOUND":
            location = self._extract_cedente_address()

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
            "location": location,  
            "catastral_ref": self._extract_catastral_ref(),
            "utm_coordinates": self._extract_utm_coordinates(),
            "energy_savings": self._extract_energy_savings(),
            "act_code": self._extract_act_code(),
            "sell_price": self._extract_sell_price(),
        }

        return result


    # ------------------------
    # Normalization & helpers
    # ------------------------

    def _normalize(self, t: str) -> str:
        if not t:
            return ""

        t = (
            t.replace("’", "'")
            .replace("´", "'")
            .replace("`", "'")
            .replace("″", "''")
        )

        # OCR rare encodings (safe minimal)
        t = t.replace("Le6n", "León").replace("Direcci6n", "Dirección").replace("ubicaci6n", "ubicación")

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
        if not txt:
            return "NOT FOUND"

        t = txt.upper()
        t = t.replace(" ", "").replace("-", "").replace(".", "").replace(":", "")

        # Buscar candidatos incluso embebidos en tokens (IDESP...13103004L)
        for m in re.finditer(r"([0-9OIL]{8})([A-Z6IL1])", t):
            raw_num, raw_last = m.groups()

            # Normaliza SOLO los 8 "dígitos"
            num = raw_num.replace("O", "0").replace("I", "1").replace("L", "1")

            # La letra final se conserva; sólo tratamos casos OCR típicos
            if raw_last == "6":
                last_candidates = ["G"]
            elif raw_last in {"1", "I", "L"}:
                # OCR puede confundir I/L con 1 en la última letra
                last_candidates = ["I", "L"]
            else:
                last_candidates = [raw_last]

            for last in last_candidates:
                cand = f"{num}{last}"
                if self._is_valid_spanish_dni(cand):
                    return cand

        # Fallback: si hay uno limpio directo
        m = re.search(r"(\d{8}[A-Z])", t)
        if m and self._is_valid_spanish_dni(m.group(1)):
            return m.group(1)

        return "NOT FOUND"


    def _find_cif(self, txt: str) -> str:
        if not txt:
            return "NOT FOUND"
        m = re.search(r"\b([A-Z]\d{8})\b", txt.upper())
        return m.group(1) if m else "NOT FOUND"

    def _normalize_nie(self, nie: str) -> str:
        if not nie:
            return "NOT FOUND"
        nie = re.sub(r"[\s\-]", "", nie).upper()
        return nie

    def _split_sections(self, t: str) -> Tuple[str, str]:
        """
        Returns (cesionario_text, cedente_text)

        IMPORTANT:
        In your real contract:
        - "De una parte" = CEDENTE
        - "Y de otra parte" = CESIONARIO
        """
        if not t:
            return "", ""

        # ✅ Template A: DE UNA PARTE / Y DE OTRA PARTE
        m1 = re.search(r"\bDE\s+UNA\s+PARTE\b", t, re.IGNORECASE)
        m2 = re.search(r"\bY\s+DE\s+OTRA\s+PARTE\b|\bY\s+DE\s+OTRA\b", t, re.IGNORECASE)

        if m1 and m2 and m2.start() > m1.start():
            cedente = t[m1.start():m2.start()]
            cesionario = t[m2.start():]
            return cesionario, cedente

        # Template B: CESIONARIO / CEDENTE
        m3 = re.search(r"\bCESIONARIO\b", t, re.IGNORECASE)
        m4 = re.search(r"\bCEDENTE\b", t, re.IGNORECASE)
        if m3 and m4 and m4.start() > m3.start():
            return t[m3.start():m4.start()], t[m4.start():]

        return "", t

    # ------------------------
    # Extractors (ACT / common)
    # ------------------------

    def _extract_act_code(self) -> str:
        t = getattr(self, "_text_norm", "") or self.text
        m = re.search(r"(RES0*\d{2,3})", t, re.IGNORECASE)
        if m:
            code = m.group(1).upper()
            code = re.sub(r"RES0+(\d)", r"RES0\1", code)
            return code
        return "NOT FOUND"

    def _extract_energy_savings(self) -> str:
        """
        Must handle: "9 180 kWh/año"
        """
        t = getattr(self, "_text_norm", "") or self.text

        m = re.search(r"(\d[\d\s.,]{2,})\s*kWh\s*/\s*a(?:ñ|n)o", t, re.IGNORECASE)
        if m:
            raw = m.group(1)
            # ✅ deja solo dígitos (mata espacios, puntos, comas y saltos)
            value = re.sub(r"[^\d]", "", raw)

            if not value:
                return "NOT FOUND"
            try:
                num_value = int(value)
                if num_value < 500:
                    return "NOT FOUND"
                return str(num_value)
            except ValueError:
                return value


        return "NOT FOUND"

    def _extract_catastral_ref(self) -> str:
        t = getattr(self, "_text_norm", "") or self.text

        m = re.search(r"Referencia\s+catastral\s*:\s*([0-9A-Z\s]+)", t, re.IGNORECASE)
        if m:
            ref = re.sub(r"\s+", "", m.group(1)).strip()
            if re.match(r"^[0-9A-Z]{10,25}$", ref):
                return ref

        m = re.search(r"\b(\d{6,8}[A-Z]{1,3}\d{2,6}[A-Z]{1,4}\d{1,6}[A-Z]{1,4})\b", t)
        if m:
            return m.group(1)

        return "NOT FOUND"

    def _extract_utm_coordinates(self) -> str:
        t = getattr(self, "_text_norm", "") or self.text

        m = re.search(r"UTM\s+(\d+),?\s*X:\s*([\d.]+),?\s*Y:\s*([\d.\s]+)", t, re.IGNORECASE)
        if m:
            zone = m.group(1)
            x = m.group(2)
            y = m.group(3).replace(" ", "")
            return f"X:{x} Y:{y} HUSO:{zone}"

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
        patterns = [
            r"Instalador[:\s]+([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s&\.-]{3,80})",
            r"El\s+Instalador[:\s]+([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s&\.-]{3,80})",
        ]
        for p in patterns:
            m = re.search(p, txt, re.IGNORECASE)
            if m:
                name = re.sub(r"\s+", " ", m.group(1)).strip()
                name = re.sub(r"\b(NIF|CIF)\b.*$", "", name, flags=re.IGNORECASE).strip()
                if len(name.split()) >= 2:
                    return name
        return "NOT FOUND"

    def _extract_location(self) -> str:
        t = getattr(self, "_text_norm", "") or self.text

        m = re.search(r"Direcci[oó]n\s*:\s*([^\n\.]{8,220})", t, re.IGNORECASE)
        if m:
            loc = re.sub(r"\s+", " ", m.group(1)).strip().rstrip(".")
            if self._is_valid_text(loc, 8):
                return loc

        m = re.search(
            r"\ben\s+(.{8,200}?(?:PLAZA|CALLE|AVENIDA|CARRETERA|CL)\b.{0,80}?(?:\b\d{5}\b|Castilla y Le[oó]n))",
            t,
            re.IGNORECASE,
        )
        if m:
            loc = re.sub(r"\s+", " ", m.group(1)).strip()
            if self._is_valid_text(loc, 8):
                return loc

        m = re.search(r"localidad\s+de\s+(.{3,80}?)(?:\s+Castilla y Le[oó]n|\s+\b\d{5}\b)", t, re.IGNORECASE)
        if m:
            loc = re.sub(r"\s+", " ", m.group(1)).strip()
            if self._is_valid_text(loc, 8):
                return loc

        return "NOT FOUND"

    # ------------------------
    # Cesionario fields (empresa)
    # ------------------------

    def _extract_cesionario_company(self) -> str:
        # ✅ extractor directo (para tu formato real)
        t = getattr(self, "_text_norm", "") or self.text
        m = re.search(
            r"Y\s+de\s+otra\s+parte,?\s*([A-Z0-9ÑÁÉÍÓÚ\.\s,]{3,80}?),\s*con\s+CIF",
            t,
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            val = re.sub(r"\s+", " ", m.group(1)).strip()
            if self._is_valid_text(val, 3):
                return val

        # fallback a secciones
        txt = getattr(self, "_cesionario_txt", "") or ""
        t2 = txt if txt else t

        val = self._find_first(
            r"DE\s+UNA\s+PARTE[,\s]+(.+?)(?:,?\s*con\s+(?:CIF|NIF)|\s+(?:CIF|NIF)\b)",
            t2
        )
        if val != "NOT FOUND":
            return val

        val = self._find_first(r"\bCESIONARIO\b\s*[:\-]?\s*(.+?)(?:\s+(?:CIF|NIF)\b|,|\n)", t2)
        if val != "NOT FOUND":
            return val

        return "NOT FOUND"

    def _extract_cesionario_cif(self) -> str:
        t = getattr(self, "_text_norm", "") or self.text
        m = re.search(r"\bCIF\s*([A-Z]\d{8})\b", t, re.IGNORECASE)
        if m:
            return m.group(1).upper()

        txt = getattr(self, "_cesionario_txt", "") or ""
        t2 = txt if txt else t
        return self._find_cif(t2)

    def _extract_cesionario_address(self) -> str:
        t = getattr(self, "_text_norm", "") or self.text

        # ✅ Captura completo aunque haya salto de línea: "Conde de Aranda\n1, 29..."
        m = re.search(
            r"domicilio\s+social\s+en\s+(.{10,220}?)(?:;|,?\s*debidamente|\n\s*debidamente)",
            t,
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            val = re.sub(r"\s+", " ", m.group(1)).strip().rstrip(",;.")
            return val if self._is_valid_text(val, 8) else "NOT FOUND"

        # fallback antiguo (por si otro template)
        txt = getattr(self, "_cesionario_txt", "") or ""
        t2 = txt if txt else t
        val = self._find_first(r"domicilio[^:]*:\s*(.+?)(?:\n|,?\s*CP|\s+C\.P\.)", t2)
        return val if self._is_valid_text(val, 8) else "NOT FOUND"


    def _extract_cesionario_representative(self) -> str:
        t = getattr(self, "_text_norm", "") or self.text

        # ✅ extractor directo
        m = re.search(
            r"representada\s+por\s+D\.\s*([A-ZÑÁÉÍÓÚ][A-Za-zÑÁÉÍÓÚ\s]{5,80}?)\s*,\s*mayor\s+de\s+edad",
            t,
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            val = re.sub(r"\s+", " ", m.group(1)).strip()
            return val if self._is_valid_text(val, 5) else "NOT FOUND"

        # fallback
        txt = getattr(self, "_cesionario_txt", "") or ""
        t2 = txt if txt else t
        val = self._find_first(
            r"representad[oa]\s+por\s+D[./]?\s*([A-ZÑÁÉÍÓÚ][A-Za-zÑÁÉÍÓÚ\s]{5,60}?)(?:,|\s+con|\n)",
            t2,
        )
        return val

    def _extract_cesionario_dni(self) -> str:
        """
        In your real docs it's NIE: 'con NIE Z161 6694-Y'
        """
        t = getattr(self, "_text_norm", "") or self.text

        m = re.search(r"\bNIE\s*([A-Z]\s*\d[\d\s\-]{5,}\s*[A-Z])\b", t, re.IGNORECASE)
        if m:
            return self._normalize_nie(m.group(1))

        # fallback (might return DNI format)
        txt = getattr(self, "_cesionario_txt", "") or ""
        t2 = txt if txt else t
        return self._find_dni(t2)

    def _check_cesionario_signature(self) -> str:
        t = getattr(self, "_text_norm", "") or self.text
        if re.search(r"Firma|Firmado|Fdo\.", t, re.IGNORECASE):
            return "Present"
        return "NOT FOUND"


    # ------------------------
    # Cedente fields (homeowner)
    # ------------------------

    def _extract_cedente_name(self) -> str:
        txt = getattr(self, "_cedente_txt", "") or self.text
        txt = self._normalize(txt)

        # ✅ extractor directo “De una parte, <NOMBRE>, mayor de edad, con DNI”
        m = re.search(
            r"De\s+una\s+parte,?\s*([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]{6,80}?)\s*,\s*mayor\s+de\s+edad,\s*con\s+DNI",
            txt,
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            name = re.sub(r"\s+", " ", m.group(1)).strip()
            if 2 <= len(name.split()) <= 6:
                return name

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
                if 2 <= len(name.split()) <= 6 and not bad_words.search(name):
                    return name

        return "NOT FOUND"

    def _extract_cedente_dni(self) -> str:
        """
        Cedente DNI extractor.

        Fixes OCR bug where DNIs ending with 'L' or 'I' were being broken by
        globally converting L/I -> 1. We now:
        - capture candidates allowing OCR noise
        - normalize ONLY the 8-digit portion
        - keep/resolve the final letter properly
        - validate via checksum
        """
        txt = getattr(self, "_cedente_txt", "") or self.text
        txt = self._normalize(txt)

        # Helper: take a captured candidate and run it through the robust finder
        def _from_candidate(candidate: str) -> str:
            val = self._find_dni(candidate)
            return val if val != "NOT FOUND" else "NOT FOUND"

        # ✅ 1) Direct pattern “De una parte ... con DNI XXXXX”
        m = re.search(
            r"De\s+una\s+parte.*?\bDNI\s*([0-9OIlL]{8}\s*[A-Z6IL1])",
            txt,
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            val = _from_candidate(m.group(1))
            if val != "NOT FOUND":
                return val

        # ✅ 2) “mayor de edad, con DNI …”
        m = re.search(
            r"mayor\s+de\s+edad,\s*con\s+DNI[:\s]*([0-9OIlL]{8}\s*[A-Z6IL1])",
            txt,
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            val = _from_candidate(m.group(1))
            if val != "NOT FOUND":
                return val

        # ✅ 3) Generic “DNI: …”
        m = re.search(
            r"\bDNI[:\s]*([0-9OIlL]{8}\s*[A-Z6IL1])\b",
            txt,
            re.IGNORECASE,
        )
        if m:
            val = _from_candidate(m.group(1))
            if val != "NOT FOUND":
                return val

        # ✅ 4) Fallback: scan the whole text (also catches IDESP…13103004L)
        val = self._find_dni(txt)
        return val if val != "NOT FOUND" else "NOT FOUND"


    def _extract_cedente_address(self) -> str:
        """
        ✅ En tu contrato real está en:
        'con domicilio en ... , teléfono ...'
        """
        t = getattr(self, "_cedente_txt", "") or (getattr(self, "_text_norm", "") or self.text)

        m = re.search(
            r"domicilio\s+en\s+(.{10,140}?)(?:,?\s*tel[eé]fono|\s+tel[eé]fono|\n)",
            t,
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            addr = re.sub(r"\s+", " ", m.group(1)).strip().rstrip(",;.")
            return addr if self._is_valid_text(addr, 8) else "NOT FOUND"

        # fallback a location
        loc = self._extract_location()
        return loc if self._is_valid_text(loc, 8) else "NOT FOUND"

    def _extract_cedente_phone(self) -> str:
        t = getattr(self, "_cedente_txt", "") or (getattr(self, "_text_norm", "") or self.text)

        m = re.search(r"tel[eé]fono[^+\d]*([+\d][\d\s-]{8,})", t, re.IGNORECASE)
        if m:
            phone = re.sub(r"[^\d+]", "", m.group(1))
            if len(re.sub(r"\D", "", phone)) >= 9:
                return phone

        m = re.search(r"\b(\+34)?\s*(6\d{8}|7\d{8}|8\d{8}|9\d{8})\b", t)
        if m:
            return re.sub(r"\s+", "", m.group(0)).replace(" ", "")

        return "NOT FOUND"

    def _extract_cedente_email(self) -> str:
        t = getattr(self, "_cedente_txt", "") or (getattr(self, "_text_norm", "") or self.text)
        m = re.search(r"[\w\.-]+@[\w\.-]+\.\w+", t)
        return m.group(0) if m else "NOT FOUND"

    def _extract_notifications(self) -> str:
        t = getattr(self, "_cedente_txt", "") or (getattr(self, "_text_norm", "") or self.text)
        m = re.search(r"[Nn]otificaciones[:\s]+(.+?)(?:\n|$)", t, re.IGNORECASE)
        if m:
            val = re.sub(r"\s+", " ", m.group(1)).strip()
            return val if self._is_valid_text(val, 8) else "NOT FOUND"
        return "NOT FOUND"

    def _check_cedente_signature(self) -> str:
        t = getattr(self, "_text_norm", "") or self.text
        signature_count = len(re.findall(r"Firma|Firmado|Fdo\.", t, re.IGNORECASE))
        return f"{signature_count} signature(s) found" if signature_count > 0 else "NOT FOUND"


    def _is_valid_spanish_dni(self, dni: str) -> bool:
        dni = (dni or "").strip().upper().replace(" ", "")
        if not re.match(r"^\d{8}[A-Z]$", dni):
            return False
        letters = "TRWAGMYFPDXBNJZSQVHLCKE"
        num = int(dni[:8])
        return dni[-1] == letters[num % 23]

    def _extract_sell_price(self) -> str:
        """
        Extrae precio de venta (€/kWh o €/IWh) del contrato.
        Busca patrones como "precio de venta", "€/kWh", "€/IWh", etc.
        """
        t = getattr(self, "_text_norm", "") or self.text

        # Patrones comunes para precio de venta
        patterns = [
            r"precio\s+de\s+venta[^€\d]*([\d,]+(?:\.\d+)?)\s*€/[kI]Wh",
            r"venta[^€\d]*([\d,]+(?:\.\d+)?)\s*€/[kI]Wh",
            r"([\d,]+(?:\.\d+)?)\s*€/[kI]Wh",
            r"precio\s+de\s+venta[^€\d]*([\d,]+(?:\.\d+)?)\s*euros?\s*por\s*[kI]Wh",
        ]

        for pattern in patterns:
            m = re.search(pattern, t, re.IGNORECASE)
            if m:
                price = m.group(1).replace(",", ".")
                try:
                    float(price)  # validar que sea número
                    return price
                except ValueError:
                    continue

        return "NOT FOUND"
