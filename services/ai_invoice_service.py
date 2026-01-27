import json
import logging
import os
import re
from datetime import date, timedelta
from typing import Any, Dict, Optional, List, Tuple

import mimetypes
import fitz  # PyMuPDF
import httpx
import openai
from openai import OpenAI

logger = logging.getLogger(__name__)

DEFAULT_MODEL = os.getenv("OPENAI_CHAT_MODEL", os.getenv("OPENAI_VISION_MODEL", "gpt-4o-mini"))
MAX_OUTPUT_TOKENS = int(os.getenv("OPENAI_MAX_OUTPUT_TOKENS", "500"))
PDF_TEXT_THRESHOLD = int(os.getenv("PDF_TEXT_THRESHOLD", "100"))
PDF_OCR_ZOOM = float(os.getenv("PDF_OCR_ZOOM", "2.0"))
_client: Optional[OpenAI] = None
_ocr_reader = None


def _get_client() -> OpenAI:
    global _client
    if _client is not None:
        return _client

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY no configurada")

    logger.info("OpenAI SDK version: %s", openai.__version__)
    logger.info("httpx version: %s", httpx.__version__)

    _client = OpenAI(api_key=api_key)
    return _client


def _extract_json(text: str) -> Dict[str, Any]:
    if not text:
        return {}
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```[a-zA-Z]*", "", cleaned)
        cleaned = cleaned.strip("` \n")
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", cleaned, re.DOTALL)
        if match:
            try:
                return json.loads(match.group(0))
            except json.JSONDecodeError:
                return {}
    return {}


def _normalize_date(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    value = value.strip()
    if re.match(r"\d{4}-\d{2}-\d{2}$", value):
        return value
    match = re.match(r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", value)
    if match:
        day, month, year = match.groups()
        if len(year) == 2:
            year = f"20{year}"
        return f"{year.zfill(4)}-{month.zfill(2)}-{day.zfill(2)}"
    match = re.match(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})", value)
    if match:
        year, month, day = match.groups()
        return f"{year.zfill(4)}-{month.zfill(2)}-{day.zfill(2)}"
    return None


def _extract_first_date(text: str) -> Optional[str]:
    if not text:
        return None
    normalized = text.replace(".", "/")
    patterns = [
        r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",
        r"(\d{4}[/-]\d{1,2}[/-]\d{1,2})",
    ]
    for pattern in patterns:
        match = re.search(pattern, normalized)
        if match:
            return _normalize_date(match.group(1))
    return None


def _find_payment_date_by_keywords(text: str) -> Optional[str]:
    if not text:
        return None
    keywords = [
        "fecha de vencimiento",
        "vencimiento",
        "vence el",
        "fecha de pago",
        "fecha pago",
    ]
    for line in text.splitlines():
        lowered = line.lower()
        if any(keyword in lowered for keyword in keywords):
            found = _extract_first_date(line)
            if found:
                return found
    return None


def _normalize_rate(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        numeric = float(value)
    else:
        cleaned = str(value).replace("%", "").strip()
        cleaned = cleaned.replace(",", ".")
        try:
            numeric = float(cleaned)
        except ValueError:
            return None
    if not (numeric >= 0):
        return None
    rounded_int = round(numeric)
    if abs(numeric - rounded_int) < 0.001:
        return float(rounded_int)
    return float(round(numeric, 2))


def _pick_first_non_empty(*values: Any) -> Any:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and value.strip() == "":
            continue
        return value
    return None


def _normalize_amount(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    raw = str(value).replace("EUR", "").replace("euro", "").strip()
    raw = raw.replace(" ", "")
    if raw.count(",") >= 1 and raw.count(".") >= 1:
        raw = raw.replace(".", "").replace(",", ".")
    elif raw.count(",") == 1 and raw.count(".") == 0:
        raw = raw.replace(",", ".")
    elif raw.count(".") >= 1 and raw.count(",") == 0:
        parts = raw.split(".")
        if len(parts[-1]) == 2:
            raw = "".join(parts[:-1]) + "." + parts[-1]
        else:
            raw = raw.replace(".", "")
    try:
        return float(raw)
    except ValueError:
        return None


def _normalize_entity_name(value: Optional[str]) -> str:
    if not value:
        return ""
    return re.sub(r"[^a-z0-9]", "", value.lower())


def _is_same_entity(candidate: Optional[str], company_names) -> bool:
    if not candidate:
        return False
    normalized = _normalize_entity_name(candidate)
    if not normalized:
        return False
    for name in company_names or []:
        if normalized == _normalize_entity_name(name):
            return True
    return False


def looks_like_person(name: Optional[str]) -> bool:
    if not name:
        return False
    cleaned = re.sub(r"[^\w\s]", " ", name).strip()
    if not cleaned:
        return False
    tokens = [token for token in cleaned.split() if token.isalpha()]
    if not tokens:
        return False
    if has_legal_form(name):
        return False
    if len(tokens) in {2, 3} and all(len(token) > 1 for token in tokens):
        return True
    return False


def has_legal_form(name: Optional[str]) -> bool:
    if not name:
        return False
    return bool(
        re.search(
            r"\b(S\.?L\.?U?\.?|S\.?A\.?U?\.?|S\.?C\.?|COOP(?:ERATIVA)?|S\.?L\.?P\.?|UTE|CB|LTD|LIMITED|INC|GMBH|SARL|BV|NV)\b",
            name,
            re.IGNORECASE,
        )
    )


def contains_forbidden_keyword(name: Optional[str]) -> bool:
    if not name:
        return False
    lowered = name.lower()
    forbidden = [
        "vendedor",
        "comercial",
        "agente",
        "transporte",
        "reparto",
        "envío",
        "envio",
        "logística",
        "logistica",
        "shipping",
    ]
    return any(keyword in lowered for keyword in forbidden)


def _is_valid_supplier(candidate: Optional[str], company_names, text: Optional[str] = None) -> bool:
    if candidate is None:
        return False
    value = str(candidate).strip()
    if not value:
        return False
    if looks_like_person(value):
        return False
    if contains_forbidden_keyword(value):
        return False
    if not has_legal_form(value):
        return False
    if _is_same_entity(value, company_names):
        return False
    if text and not _supplier_has_near_tax_id(text, value):
        return False
    return True


def _looks_like_metadata(line: str) -> bool:
    lowered = line.lower()
    blocked = [
        "factura",
        "fecha",
        "nif",
        "cif",
        "dni",
        "iva",
        "total",
        "base",
        "importe",
        "pedido",
    ]
    if any(word in lowered for word in blocked):
        return True
    letters = sum(char.isalpha() for char in line)
    digits = sum(char.isdigit() for char in line)
    if letters < 3:
        return True
    if digits > letters * 2:
        return True
    return False


def _contains_legal_form(line: str) -> bool:
    return has_legal_form(line)


def _has_tax_id(line: str) -> bool:
    if not line:
        return False
    patterns = [
        r"\b[A-HJ-NP-SUVW]\s?-?\d{7}\s?-?[0-9A-J]\b",  # CIF con separadores
        r"\b\d{8}\s?-?[A-Z]\b",  # NIF con separador
        r"\b[A-Z]{2}\s?-?\d{6,12}\b",  # VAT/IVA intracomunitario
    ]
    return any(re.search(pattern, line, re.IGNORECASE) for pattern in patterns)


def _supplier_has_near_tax_id(text: str, supplier: str, window: int = 2) -> bool:
    if not text or not supplier:
        return False
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    normalized_supplier = _normalize_entity_name(supplier)
    if not normalized_supplier:
        return False
    for idx, line in enumerate(lines):
        if normalized_supplier in _normalize_entity_name(line):
            start = max(0, idx - window)
            end = min(len(lines), idx + window + 1)
            for candidate in lines[start:end]:
                if _has_tax_id(candidate):
                    return True
            return False
    return False


def _extract_supplier_candidates(text: str, company_names=None) -> List[Tuple[str, int]]:
    if not text:
        return []
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return []

    supplier_keywords = [
        "expedido por",
        "emisor",
        "proveedor",
        "facturado por",
        "vendedor",
        "issued by",
        "seller",
    ]
    client_keywords = [
        "cliente",
        "enviado a",
        "destinatario",
        "facturado a",
        "receptor",
        "bill to",
        "ship to",
    ]
    operational_keywords = [
        "transporte",
        "envío",
        "expedición",
        "mensajería",
        "portes",
        "logística",
        "shipping",
    ]

    header_lines = lines[:8]
    line_counts = {}
    for line in lines:
        key = _normalize_entity_name(line)
        if key:
            line_counts[key] = line_counts.get(key, 0) + 1

    candidates: List[Tuple[str, int]] = []
    for idx, line in enumerate(lines):
        lowered = line.lower()
        if any(keyword in lowered for keyword in client_keywords):
            continue
        if any(keyword in lowered for keyword in operational_keywords) and not _contains_legal_form(line):
            continue
        if _looks_like_metadata(line):
            continue
        if _is_same_entity(line, company_names):
            continue
        if contains_forbidden_keyword(line):
            continue

        score = 0
        if line in header_lines:
            score += 15
        if _contains_legal_form(line):
            score += 80
        if _has_tax_id(line):
            score += 30
        if line_counts.get(_normalize_entity_name(line), 0) > 1:
            score += 10
        if any(keyword in lowered for keyword in supplier_keywords):
            score += 25

        if score <= 0:
            continue
        candidates.append((line, score))

    return candidates


def _select_best_supplier(text: str, company_names=None) -> Optional[str]:
    if not text:
        return None
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    supplier_keywords = [
        "expedido por",
        "emisor",
        "proveedor",
        "facturado por",
        "vendedor",
        "issued by",
        "seller",
    ]
    client_keywords = [
        "cliente",
        "enviado a",
        "destinatario",
        "facturado a",
        "receptor",
        "bill to",
        "ship to",
    ]
    anchor_keywords = [
        "titular",
        "iban",
        "datos bancarios",
        "datos fiscales",
    ]

    for idx, line in enumerate(lines):
        lowered = line.lower()
        if any(keyword in lowered for keyword in supplier_keywords):
            for keyword in supplier_keywords:
                if keyword in lowered:
                    parts = re.split(keyword, line, flags=re.IGNORECASE)
                    if len(parts) > 1:
                        candidate = parts[1].strip(" :-")
                        if _is_valid_supplier(candidate, company_names, text):
                            return candidate
            for offset in (1, 2):
                if idx + offset < len(lines):
                    candidate = lines[idx + offset].strip()
                    if _is_valid_supplier(candidate, company_names, text):
                        return candidate

    for idx, line in enumerate(lines):
        lowered = line.lower()
        if any(anchor in lowered for anchor in anchor_keywords):
            parts = line.split(":", 1)
            if len(parts) > 1 and _is_valid_supplier(parts[1], company_names, text):
                return parts[1].strip()
            for offset in (1, 2):
                if idx + offset < len(lines):
                    candidate = lines[idx + offset].strip()
                    if _is_valid_supplier(candidate, company_names, text):
                        return candidate

    candidates = _extract_supplier_candidates(text, company_names)
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[1], reverse=True)
    best, score = candidates[0]
    if _is_valid_supplier(best, company_names, text):
        return best
    return None


def _extract_supplier_from_text(text: str, company_names=None) -> Optional[str]:
    return _select_best_supplier(text, company_names)


def _validate_math(
    base_amount: Optional[float],
    vat_amount: Optional[float],
    total_amount: Optional[float],
) -> Dict[str, Any]:
    if base_amount is None or vat_amount is None or total_amount is None:
        return {"is_consistent": None, "difference": None}
    difference = round((base_amount + vat_amount) - total_amount, 2)
    tolerance = max(0.05, total_amount * 0.01)
    return {
        "is_consistent": abs(difference) <= tolerance,
        "difference": difference,
    }


def _round_amount(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return round(float(value), 2)


def _has_vat_exemption_indicators(text: str) -> bool:
    if not text:
        return False
    lowered = text.lower()
    keywords = [
        "exento",
        "exenta",
        "exencion",
        "inversion del sujeto pasivo",
        "inversion sujeto pasivo",
        "intracomunitaria",
        "iva incluido",
        "iva incluida",
        "iva incl",
    ]
    return any(keyword in lowered for keyword in keywords)


def _is_text_significant(text: str, min_chars: int = 100) -> bool:
    if not text:
        return False
    useful_chars = sum(1 for char in text if char.isalnum())
    return useful_chars >= min_chars


def _extract_pdf_text(file_path: str) -> str:
    with fitz.open(file_path) as doc:
        parts = []
        for page in doc:
            parts.append(page.get_text("text"))
        return "\n".join(parts).strip()


def _extract_pdf_text_from_bytes(data: bytes) -> str:
    with fitz.open(stream=data, filetype="pdf") as doc:
        parts = []
        for page in doc:
            parts.append(page.get_text("text"))
        return "\n".join(parts).strip()


def _get_ocr_reader():
    global _ocr_reader
    if _ocr_reader is not None:
        return _ocr_reader
    try:
        import easyocr
    except ImportError as exc:
        logger.warning("EasyOCR no disponible: %s", exc)
        return None
    model_dir = os.getenv("EASYOCR_MODEL_STORAGE_DIRECTORY", "/opt/easyocr-models")
    if not os.path.isdir(model_dir):
        logger.warning("Directorio de modelos EasyOCR no existe: %s", model_dir)
    download_env = os.getenv("EASYOCR_DOWNLOAD_ENABLED", "").strip().lower()
    download_enabled = download_env in {"1", "true", "yes"}
    runtime_env = os.getenv("ENV", "").strip().lower()
    model_contents = []
    if os.path.isdir(model_dir):
        try:
            model_contents = os.listdir(model_dir)
        except OSError:
            model_contents = []
    if not os.path.isdir(model_dir) or not model_contents:
        if runtime_env == "production" and not download_enabled:
            logger.warning(
                "Modelos EasyOCR no encontrados y descarga deshabilitada en producción. OCR omitido."
            )
            return None
        if not download_enabled:
            download_enabled = True
            logger.warning("Modelos EasyOCR no encontrados. Se habilita descarga automática.")
    try:
        _ocr_reader = easyocr.Reader(
            ["es", "en"],
            gpu=False,
            model_storage_directory=model_dir,
            download_enabled=download_enabled,
        )
    except Exception as exc:
        logger.warning("Error inicializando EasyOCR (model_dir=%s): %s", model_dir, exc)
        return None
    return _ocr_reader


def _extract_pdf_text_ocr(file_path: str) -> str:
    reader = _get_ocr_reader()
    if reader is None:
        return ""
    try:
        import numpy as np
    except ImportError as exc:
        logger.warning("NumPy no disponible para OCR: %s", exc)
        return ""

    parts = []
    matrix = fitz.Matrix(PDF_OCR_ZOOM, PDF_OCR_ZOOM)
    with fitz.open(file_path) as doc:
        for page in doc:
            pix = page.get_pixmap(matrix=matrix)
            image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.height, pix.width, pix.n
            )
            if pix.n == 4:
                image = image[:, :, :3]
            lines = reader.readtext(image, detail=0)
            if lines:
                parts.append("\n".join(lines))
    return "\n".join(parts).strip()


def _extract_pdf_text_ocr_from_bytes(data: bytes) -> str:
    reader = _get_ocr_reader()
    if reader is None:
        return ""
    try:
        import numpy as np
    except ImportError as exc:
        logger.warning("NumPy no disponible para OCR: %s", exc)
        return ""

    parts = []
    matrix = fitz.Matrix(PDF_OCR_ZOOM, PDF_OCR_ZOOM)
    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            pix = page.get_pixmap(matrix=matrix)
            image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.height, pix.width, pix.n
            )
            if pix.n == 4:
                image = image[:, :, :3]
            lines = reader.readtext(image, detail=0)
            if lines:
                parts.append("\n".join(lines))
    return "\n".join(parts).strip()


def _extract_image_text_ocr(file_path: str) -> str:
    reader = _get_ocr_reader()
    if reader is None:
        return ""
    lines = reader.readtext(file_path, detail=0)
    if not lines:
        return ""
    return "\n".join(lines).strip()


def _extract_image_text_ocr_from_bytes(data: bytes) -> str:
    reader = _get_ocr_reader()
    if reader is None:
        return ""
    try:
        import numpy as np
        import cv2
    except ImportError as exc:
        logger.warning("Dependencias de OCR no disponibles: %s", exc)
        return ""

    image_array = np.frombuffer(data, dtype=np.uint8)
    image = cv2.imdecode(image_array, cv2.IMREAD_COLOR)
    if image is None:
        return ""
    lines = reader.readtext(image, detail=0)
    if not lines:
        return ""
    return "\n".join(lines).strip()


def analyze_invoice(
    file_path: Optional[str] = None,
    file_bytes: Optional[bytes] = None,
    filename: Optional[str] = None,
    mime_type: Optional[str] = None,
    document_type: str = "expense",
    company_names: Optional[list] = None,
) -> Dict[str, Any]:
    client = _get_client()
    if file_bytes is None:
        if not file_path:
            raise ValueError("file_path o file_bytes es requerido")
        with open(file_path, "rb") as handle:
            file_bytes = handle.read()

    if not filename:
        filename = os.path.basename(file_path) if file_path else "archivo"

    extension = os.path.splitext(filename)[1].lower()
    mime_type = mime_type or mimetypes.guess_type(filename)[0] or ""

    is_pdf = extension == ".pdf" or mime_type == "application/pdf"
    is_image = extension in {".jpg", ".jpeg", ".png"} or mime_type in {
        "image/jpeg",
        "image/png",
    }
    file_kind = "pdf" if is_pdf else "image" if is_image else "unknown"
    logger.info(
        "Tipo de archivo procesado (%s): tipo=%s extension=%s mime=%s",
        filename,
        file_kind,
        extension,
        mime_type,
    )

    extracted_text = ""
    embedded_text = ""
    used_ocr = False
    pdf_kind = None
    if is_pdf:
        embedded_text = _extract_pdf_text_from_bytes(file_bytes)
        text_length = len(embedded_text.strip())
        is_significant = _is_text_significant(embedded_text, PDF_TEXT_THRESHOLD)
        is_scanned = not is_significant
        ocr_text = ""
        if is_scanned:
            ocr_text = _extract_pdf_text_ocr_from_bytes(file_bytes)
            extracted_text = ocr_text
            used_ocr = True
            pdf_kind = "scanned"
        else:
            extracted_text = embedded_text
            pdf_kind = "original"
        logger.info("PDF tratado como escaneado (%s): %s", filename, is_scanned)
        logger.info("Longitud texto extraido (%s): %s", filename, text_length)
        logger.info("Texto significativo (%s): %s", filename, is_significant)
        logger.info("Longitud texto OCR (%s): %s", filename, len(ocr_text.strip()))
        logger.info("Texto PDF extraido (%s):\n%s", filename, extracted_text)
    elif is_image:
        extracted_text = _extract_image_text_ocr_from_bytes(file_bytes)
        used_ocr = True
        pdf_kind = "image"
        logger.info("OCR aplicado a imagen (%s).", filename)
    else:
        logger.warning("Tipo de archivo no soportado (%s). Texto vacio enviado.", filename)

    logger.info(
        "OCR usado (%s): %s | Longitud texto final: %s",
        filename,
        used_ocr,
        len(extracted_text.strip()),
    )

    is_income = document_type == "income"
    if is_income:
        prompt = (
            "Analiza el siguiente texto extraido de una factura emitida (ingreso). "
            "Devuelve SOLO JSON valido con estas claves: "
            "client, invoice_date, payment_date, base_amount, vat_rate, vat_amount, total_amount. "
            "Usa null si no puedes inferir un dato con seguridad. "
            "No incluyas texto adicional fuera del JSON.\n\n"
            f"TEXTO_FACTURA:\n{extracted_text}"
        )
    else:
        prompt = (
            "Analiza el siguiente texto extraido de una factura recibida (gasto). "
            "Devuelve SOLO JSON valido con estas claves: "
            "supplier, invoice_date, payment_date, base_amount, vat_rate, vat_amount, total_amount. "
            "El supplier debe ser la razon social del emisor (forma juridica si aparece) "
            "y no debe ser el cliente/receptor. "
            "Usa null si no puedes inferir un dato con seguridad. "
            "No incluyas texto adicional fuera del JSON.\n\n"
            f"TEXTO_FACTURA:\n{extracted_text}"
        )

    logger.info("Prompt enviado (%s): %s", filename, prompt)

    response = client.chat.completions.create(
        model=DEFAULT_MODEL,
        max_tokens=MAX_OUTPUT_TOKENS,
        temperature=0,
        messages=[
            {"role": "user", "content": prompt},
        ],
    )

    raw_text = ""
    if response.choices:
        raw_text = response.choices[0].message.content or ""
    logger.info("Respuesta cruda modelo (%s): %s", filename, raw_text)

    data = _extract_json(raw_text)

    provider_name = (
        data.get("supplier")
        or data.get("proveedor")
        or data.get("provider_name")
        or data.get("provider")
    )
    client_name = (
        data.get("client")
        or data.get("cliente")
        or data.get("customer")
        or data.get("client_name")
    )
    invoice_date = _normalize_date(
        data.get("invoice_date") or data.get("fecha_factura") or data.get("fecha")
    )
    payment_date = _normalize_date(
        data.get("payment_date")
        or data.get("fecha_pago")
        or data.get("fecha_vencimiento")
        or data.get("vencimiento")
    )
    base_amount = _normalize_amount(
        data.get("base_amount") or data.get("base_imponible") or data.get("base")
    )
    vat_rate = _normalize_rate(
        _pick_first_non_empty(
            data.get("vat_rate"),
            data.get("iva_rate"),
            data.get("tipo_iva"),
            data.get("iva"),
        )
    )
    vat_amount = _normalize_amount(
        data.get("vat_amount") or data.get("importe_iva") or data.get("iva_importe")
    )
    total_amount = _normalize_amount(
        data.get("total_amount") or data.get("total_factura") or data.get("total")
    )

    if company_names is None:
        company_names = []

    if document_type != "income":
        supplier_source_text = embedded_text if pdf_kind == "original" else extracted_text
        provider_name = provider_name.strip() if isinstance(provider_name, str) else provider_name
        if provider_name is not None and not _is_valid_supplier(
            provider_name, company_names, supplier_source_text
        ):
            provider_name = None
        if provider_name is None:
            heuristic_supplier = _extract_supplier_from_text(
                supplier_source_text,
                company_names,
            )
            if heuristic_supplier is not None and not _is_valid_supplier(
                heuristic_supplier, company_names, supplier_source_text
            ):
                heuristic_supplier = None
            provider_name = heuristic_supplier

    if payment_date is None:
        payment_date = _find_payment_date_by_keywords(extracted_text)
    if payment_date is None and invoice_date:
        try:
            payment_date = (date.fromisoformat(invoice_date) + timedelta(days=30)).isoformat()
        except ValueError:
            payment_date = None

    assumed_vat = False
    if vat_rate is None and not _has_vat_exemption_indicators(extracted_text):
        vat_rate = 21.0
        assumed_vat = True
        if base_amount is not None:
            vat_amount = round(base_amount * 0.21, 2)
            total_amount = round(base_amount + vat_amount, 2)
        elif total_amount is not None:
            base_amount = round(total_amount / 1.21, 2)
            vat_amount = round(total_amount - base_amount, 2)

    base_amount = _round_amount(base_amount)
    vat_amount = _round_amount(vat_amount)
    total_amount = _round_amount(total_amount)

    validation = _validate_math(base_amount, vat_amount, total_amount)

    if assumed_vat:
        logger.info(
            "IVA asumido automaticamente (%s): rate=%s base=%s iva=%s total=%s",
            filename,
            vat_rate,
            base_amount,
            vat_amount,
            total_amount,
        )

    logger.info(
        "Valores detectados (%s): proveedor=%s cliente=%s fecha=%s pago=%s base=%s iva_rate=%s iva_importe=%s total=%s",
        filename,
        provider_name,
        client_name,
        invoice_date,
        payment_date,
        base_amount,
        vat_rate,
        vat_amount,
        total_amount,
    )

    return {
        "provider_name": provider_name,
        "client_name": client_name,
        "invoice_date": invoice_date,
        "payment_date": payment_date,
        "base_amount": base_amount,
        "vat_rate": vat_rate,
        "vat_amount": vat_amount,
        "total_amount": total_amount,
        "analysis_text": raw_text[:500],
        "validation": validation,
    }
