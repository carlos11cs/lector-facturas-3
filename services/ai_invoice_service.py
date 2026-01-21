import json
import logging
import os
import re
from typing import Any, Dict, Optional

import mimetypes
import fitz  # PyMuPDF
import httpx
import openai
from openai import OpenAI

logger = logging.getLogger(__name__)

DEFAULT_MODEL = os.getenv("OPENAI_CHAT_MODEL", os.getenv("OPENAI_VISION_MODEL", "gpt-4o-mini"))
MAX_OUTPUT_TOKENS = int(os.getenv("OPENAI_MAX_OUTPUT_TOKENS", "500"))
PDF_TEXT_THRESHOLD = int(os.getenv("PDF_TEXT_THRESHOLD", "50"))
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
    _ocr_reader = easyocr.Reader(["es", "en"], gpu=False)
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
    used_ocr = False
    if is_pdf:
        extracted_text = _extract_pdf_text_from_bytes(file_bytes)
        text_length = len(extracted_text.strip())
        is_scanned = text_length < PDF_TEXT_THRESHOLD
        ocr_text = ""
        if is_scanned:
            ocr_text = _extract_pdf_text_ocr_from_bytes(file_bytes)
            if ocr_text:
                extracted_text = ocr_text
            used_ocr = True
        logger.info("PDF tratado como escaneado (%s): %s", filename, is_scanned)
        logger.info("Longitud texto extraido (%s): %s", filename, text_length)
        logger.info("Longitud texto OCR (%s): %s", filename, len(ocr_text.strip()))
        logger.info("Texto PDF extraido (%s):\n%s", filename, extracted_text)
    elif is_image:
        extracted_text = _extract_image_text_ocr_from_bytes(file_bytes)
        used_ocr = True
        logger.info("OCR aplicado a imagen (%s).", filename)
    else:
        logger.warning("Tipo de archivo no soportado (%s). Texto vacio enviado.", filename)

    logger.info(
        "OCR usado (%s): %s | Longitud texto final: %s",
        filename,
        used_ocr,
        len(extracted_text.strip()),
    )

    prompt = (
        "Analiza el siguiente texto extraido de una factura. "
        "Devuelve SOLO JSON valido con estas claves: "
        "supplier, invoice_date, base_amount, vat_rate, vat_amount, total_amount. "
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
    invoice_date = _normalize_date(
        data.get("invoice_date") or data.get("fecha_factura") or data.get("fecha")
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
        "Valores detectados (%s): proveedor=%s fecha=%s base=%s iva_rate=%s iva_importe=%s total=%s",
        filename,
        provider_name,
        invoice_date,
        base_amount,
        vat_rate,
        vat_amount,
        total_amount,
    )

    return {
        "provider_name": provider_name,
        "invoice_date": invoice_date,
        "base_amount": base_amount,
        "vat_rate": vat_rate,
        "vat_amount": vat_amount,
        "total_amount": total_amount,
        "analysis_text": raw_text[:500],
        "validation": validation,
    }
