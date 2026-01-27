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

        # Fallback helpers
        def fallback_context(field, prefer_types=None, match_name=None, match_dni=None):
            context = getattr(self, 'context_data', None)
            if not context or not isinstance(context, dict):
                return None
            prefer_types = prefer_types or ["CONTRATO", "FACTURA", "CEE", "CEE_FINAL"]
            for doc_type in prefer_types:
                docs = context.get(doc_type)
                if not docs:
                    continue
                if isinstance(docs, dict):
                    docs = [docs]
                for doc in docs:
                    # Si se pide coincidencia de nombre/dni, filtra
                    if match_name and match_dni:
                        n2 = doc.get("homeowner_name") or doc.get("client_name")
                        d2 = doc.get("homeowner_dni") or doc.get("dni_number")
                        if not (n2 and d2 and n2.strip().upper() == match_name.strip().upper() and d2.strip().upper() == match_dni.strip().upper()):
                            continue
                    val = doc.get(field)
                    if val and val != "NOT FOUND":
                        return val
            return None

        # 1. Nombre
        name = self._extract_homeowner_name()
        if not name or name == "NOT FOUND":
            name_ctx = fallback_context("homeowner_name")
            if name_ctx:
                name = name_ctx

        # 2. DNI
        dni = self._extract_homeowner_dni()
        if not dni or dni == "NOT FOUND":
            dni_ctx = fallback_context("homeowner_dni")
            if dni_ctx:
                dni = dni_ctx

        # 3. Dirección (ya tiene fallback propio)
        address = self._extract_homeowner_address()

        # 4. Referencia catastral
        catastral = self._extract_catastral_ref()
        if not catastral or catastral == "NOT FOUND":
            catastral_ctx = fallback_context("catastral_ref", match_name=name, match_dni=dni)
            if catastral_ctx:
                catastral = catastral_ctx

        # 5. Código de actuación
        act_code = self._extract_act_code()
        if not act_code or act_code == "NOT FOUND":
            act_code_ctx = fallback_context("act_code", match_name=name, match_dni=dni)
            if act_code_ctx:
                act_code = act_code_ctx

        # 6. Ahorro energético
        energy_savings = self._extract_energy_savings()
        if not energy_savings or energy_savings == "NOT FOUND":
            energy_ctx = fallback_context("energy_savings", match_name=name, match_dni=dni)
            if energy_ctx:
                energy_savings = energy_ctx

        # 7. Superficie
        surface = self._extract_surface()
        if not surface or surface == "NOT FOUND":
            surface_ctx = fallback_context("surface", match_name=name, match_dni=dni)
            if surface_ctx:
                surface = surface_ctx

        # 8. Tipo de aislamiento - DESHABILITADO PARA DECLARACION
        isolation_type = "NOT FOUND"

        # 9. Firma (no tiene sentido fallback)
        signature = self._check_signature()

        return {
            "document_type": "DECLARACION",
            "homeowner_name": name,
            "homeowner_dni": dni,
            "homeowner_address": address,
            "catastral_ref": catastral,
            "act_code": act_code,
            "energy_savings": energy_savings,
            "surface": surface,
            "isolation_type": isolation_type,
            "signature": signature,
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
            if 2 <= len(name.split()) <= 8 and "PROPIETARIO INICIAL DEL AHORRO" not in name.upper():
                return name

        # 2) Titular / Solicitante / Beneficiario
        m = re.search(
            r"(?:TITULAR|SOLICITANTE|BENEFICIARIO)\s*[:\-]?\s*([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]{10,80})",
            t.upper(),
            re.IGNORECASE,
        )
        if m:
            name = re.sub(r"\s+", " ", m.group(1)).strip()
            if 2 <= len(name.split()) <= 8 and "PROPIETARIO INICIAL DEL AHORRO" not in name.upper():
                return name

        # 3) Don/Doña/D.
        m = re.search(
            r"\bD(?:ON|OÑA|\.)\s+([A-ZÑÁÉÍÓÚ][A-ZÑÁÉÍÓÚ\s]{10,80})\b",
            t.upper(),
            re.IGNORECASE,
        )
        if m:
            name = re.sub(r"\s+", " ", m.group(1)).strip()
            if 2 <= len(name.split()) <= 10 and "PROPIETARIO INICIAL DEL AHORRO" not in name.upper():
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
        debug_label = "[DECLARACION] Dirección extraída:"
        def is_garbage(val: str) -> bool:
            v = (val or "").strip()
            if not v:
                return True
            v_low = v.lower()
            if v_low in {"teléfono", "telefono", "tel", "telf", "tfno", "firma", "firmado"}:
                return True
            if v.strip().upper() in {"C", "CL", "CALLE", "AV", "AVD", "S/N"}:
                return True
            return False

        # 1) Patrones clásicos
        patterns = [
            r"Domicilio\s+(.+?)(?=\n|,\s*\d{5}|CP|C\.P\.|Tel|Firma)",
            r"Direcci[oó]n\s+(.+?)(?=\n|CP|C\.P\.|Tel|Firma)",
            r"(?:Domicilio|Direcci[oó]n)\s*[:\-]?\s*\n\s*(.{10,220})",
        ]
        for pattern in patterns:
            m = re.search(pattern, t, re.IGNORECASE | re.DOTALL)
            if m:
                val = re.sub(r"\s+", " ", m.group(1)).strip().rstrip(",;.")
                val = self._clean_address(val)
                if not is_garbage(val) and val != "NOT FOUND":
                    print(f"{debug_label} (patrón clásico) '{val}'")
                    return val
                # Si el valor es 'NOT FOUND', ignora y sigue buscando

        # 2) Fallback: busca la primera línea larga con número (como en contrato)
        lines = [ln.strip() for ln in t.splitlines() if len(ln.strip()) > 10]
        candidates = [ln for ln in lines if re.search(r"\d", ln) and not is_garbage(ln)]
        if candidates:
            # Prefiere la más larga con número
            best = max(candidates, key=len)
            best = self._clean_address(best)
            if not is_garbage(best):
                print(f"{debug_label} (línea larga con número) '{best}'")
                return best

        # 3) Fallback: busca cualquier cosa que parezca dirección
        m = re.search(r"([A-ZÁÉÍÓÚÑ]{2,}\s+\d{1,4}(?:[\w\s,.-]{0,40})\d{5})", t, re.IGNORECASE)
        if m:
            val = self._clean_address(m.group(1))
            if not is_garbage(val):
                print(f"{debug_label} (regex dirección) '{val}'")
                return val

        # 4) Fallback: buscar dirección en otros documentos si hay acceso a todos los datos
        # Busca en self.context_data si está disponible (debe ser dict con otros docs)
        name = self._extract_homeowner_name()
        dni = self._extract_homeowner_dni()
        context = getattr(self, 'context_data', None)
        if context and isinstance(context, dict):
            # Busca en CONTRATO primero, aunque no coincida nombre/dni
            contrato = context.get("CONTRATO")
            if contrato:
                if isinstance(contrato, dict):
                    contrato = [contrato]
                for doc in contrato:
                    addr = doc.get("homeowner_address")
                    if addr and not is_garbage(addr) and addr != "NOT FOUND":
                        print(f"{debug_label} (copiado de CONTRATO) '{addr}'")
                        return addr
            # Si no hay en CONTRATO, busca en FACTURA, CEE, CEE_FINAL con coincidencia de nombre/dni
            for doc_type in ["FACTURA", "CEE", "CEE_FINAL"]:
                docs = context.get(doc_type)
                if not docs:
                    continue
                if isinstance(docs, dict):
                    docs = [docs]
                for doc in docs:
                    n2 = doc.get("homeowner_name") or doc.get("client_name")
                    d2 = doc.get("homeowner_dni") or doc.get("dni_number")
                    if n2 and d2 and n2.strip().upper() == name.strip().upper() and d2.strip().upper() == dni.strip().upper():
                        for addr_field in ["homeowner_address", "address", "location"]:
                            addr = doc.get(addr_field)
                            if addr and not is_garbage(addr) and addr != "NOT FOUND":
                                print(f"{debug_label} (copiado de {doc_type} campo {addr_field}) '{addr}'")
                                return addr

        print(f"{debug_label} (NO ENCONTRADA) 'NOT FOUND'")
        return "NOT FOUND"

    def _extract_catastral_ref(self) -> str:
        t = self.text or ""

        m = re.search(
            r"Referencia\s+catastral.*?([0-9A-Z]{10,25})",
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

    def _extract_energy_savings(self) -> str:
        patterns = [
            r"Ahorro.*?AE\s+([\d.]+)",
            r"AE\s+([\d.]+)",
            r"AE.*?([\d.]+)",
        ]

        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                return f"{match.group(1)}"
        return "NOT FOUND"

    def _extract_surface(self) -> str:
        patterns = [
            r"superficie.*?(\d+(?:[.,]\d+)?)",
            r"s\s*=\s*(\d+(?:[.,]\d+)?)",
            r"área.*?(\d+(?:[.,]\d+)?)",
        ]
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                return match.group(1).replace(",", ".")
        return "NOT FOUND"

    def _extract_isolation_type(self) -> str:
        patterns = [
            r"(ROLLO|SOPLADO)",
        ]
        for pattern in patterns:
            match = re.search(pattern, self.text, re.IGNORECASE)
            if match:
                return match.group(1).upper()
        return "NOT FOUND"

    def _check_signature(self) -> str:
        t = self.text or ""
        return "Present" if re.search(r"Firma|Firmado|Fdo\.", t, re.IGNORECASE) else "NOT FOUND"

    def _clean_address(self, addr: str) -> str:
        if not addr:
            return addr

        a = re.sub(r"\s+", " ", addr).strip()

        # Quita prefijos específicos de basura OCR/plantilla
        a = re.sub(r"de la instalación en que se ejecutó\s*", "", a, flags=re.IGNORECASE)
        a = re.sub(r'^[!"]*\s*', '', a)  # Quita " ! al inicio

        # Quita prefijos típicos de plantilla/OCR
        a = re.sub(
            r"^(?:domicilio|direcci[oó]n|direccion|postal|c[oó]digo\s*postal|postal de la instalaci[oó]n.*?|de la instalaci[oó]n.*?)(?:\s+en\s+que\s+se\s+ejecut[oó].*?)?\s*[:\-]?\s*",
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

        # Si contiene teléfono o firma, no es dirección
        if re.search(r"Tel[eé]fono|Correo electr[oó]nico|Firma|Firmado", a, re.IGNORECASE):
            return "NOT FOUND"

        return a

