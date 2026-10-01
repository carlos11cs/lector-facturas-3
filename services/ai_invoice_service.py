import base64
import gc
import json
import logging
import math
import os
import re
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from difflib import SequenceMatcher
from email.utils import parsedate_to_datetime
from html import unescape
from itertools import combinations
from typing import Any, Dict, Optional, List, Tuple, Union

import mimetypes
try:
    import fitz  # PyMuPDF
except ModuleNotFoundError:
    fitz = None
try:
    import httpx
except ModuleNotFoundError:
    httpx = None
try:
    import openai
    from openai import OpenAI
except ModuleNotFoundError:
    openai = None
    OpenAI = None

logger = logging.getLogger(__name__)

# Financial document extraction has its own model configured at request time.
# Generic assistant features keep using DEFAULT_MODEL below.
DEFAULT_MODEL = os.getenv("OPENAI_CHAT_MODEL", os.getenv("OPENAI_VISION_MODEL", "gpt-4o-mini"))
MAX_OUTPUT_TOKENS = int(os.getenv("OPENAI_MAX_OUTPUT_TOKENS", "500"))
DEFAULT_INVOICE_MAX_OUTPUT_TOKENS = 32768
DEFAULT_INVOICE_TIMEOUT_SECONDS = 240
DEFAULT_INVOICE_REASONING_EFFORT = "low"
DEFAULT_INVOICE_AUDIT_REASONING_EFFORT = "high"
INVOICE_REASONING_EFFORTS = {"none", "minimal", "low", "medium", "high", "xhigh"}
PDF_TEXT_THRESHOLD = int(os.getenv("PDF_TEXT_THRESHOLD", "100"))
DEFAULT_INVOICE_V2_FAST_TEXT_MAX_CHARS = 30000
PDF_OCR_ZOOM = float(os.getenv("PDF_OCR_ZOOM", "2.0"))
OCR_MAX_PAGES = int(os.getenv("OCR_MAX_PAGES", "5"))
OCR_MAX_SECONDS = int(os.getenv("OCR_MAX_SECONDS", "7"))
OCR_MAX_DIM = int(os.getenv("OCR_MAX_DIM", "1600"))
OCR_TIMEOUT_SECONDS = 60
INVOICE_VISION_MAX_PAGES = max(1, int(os.getenv("OPENAI_INVOICE_VISION_MAX_PAGES", "2")))
INVOICE_VISION_MAX_DIMENSION = max(
    900, int(os.getenv("OPENAI_INVOICE_VISION_MAX_DIMENSION", "1400"))
)
_client: Optional[OpenAI] = None
_ocr_reader = None
_EU_AMOUNT_RE = re.compile(r"\d{1,3}(?:[.\s]\d{3})*,\d{2}|\d+,\d{2}")
_EU_THOUSANDS_RE = re.compile(r"^\d{1,3}\.\d{3},\d{2}$")


class InvoiceAnalysisResponseError(RuntimeError):
    """A terminal API or response-contract failure for invoice extraction."""

    def __init__(
        self,
        status: str,
        detail: Optional[str] = None,
        *,
        metadata: Optional[Dict[str, Any]] = None,
    ):
        self.status = status
        self.detail = detail or ""
        # Metadata is explicitly whitelisted operational data. It must never
        # contain a request body, document content, prompt, or credentials.
        self.metadata = metadata or {}
        super().__init__(f"Invoice analysis response is not usable: {status}")


_RATE_LIMIT_HEADER_FIELDS = {
    "x-ratelimit-limit-requests": "requests_limit",
    "x-ratelimit-remaining-requests": "requests_remaining",
    "x-ratelimit-reset-requests": "requests_reset",
    "x-ratelimit-limit-tokens": "tokens_limit",
    "x-ratelimit-remaining-tokens": "tokens_remaining",
    "x-ratelimit-reset-tokens": "tokens_reset",
}
_NON_RETRYABLE_RATE_LIMIT_CODES = {
    "insufficient_quota",
    "project_spend_limit_exceeded",
    "organization_spend_limit_exceeded",
    "credit_balance_exhausted",
    "billing_hard_limit_reached",
}


def _safe_error_text(value: Any, limit: int = 255) -> Optional[str]:
    if value is None:
        return None
    text_value = str(value).strip()
    return text_value[:limit] if text_value else None


def _retry_after_seconds(headers: Any) -> Optional[int]:
    """Read a server-directed retry delay without retaining raw headers."""
    if not headers:
        return None
    retry_after_ms = headers.get("retry-after-ms")
    if retry_after_ms:
        try:
            return max(1, math.ceil(float(retry_after_ms) / 1000))
        except (TypeError, ValueError):
            pass
    retry_after = headers.get("retry-after")
    if not retry_after:
        return None
    try:
        return max(1, math.ceil(float(retry_after)))
    except (TypeError, ValueError):
        try:
            retry_at = parsedate_to_datetime(str(retry_after))
            if retry_at.tzinfo is None:
                retry_at = retry_at.replace(tzinfo=timezone.utc)
            return max(
                1,
                math.ceil(
                    (retry_at.astimezone(timezone.utc) - datetime.now(timezone.utc)).total_seconds()
                ),
            )
        except (TypeError, ValueError, IndexError, OverflowError):
            return None


def _classify_rate_limit(metadata: Dict[str, Any]) -> Tuple[str, bool]:
    """Classify a 429 using only documented error fields and rate-limit headers."""
    error_code = (metadata.get("error_code") or "").lower()
    error_type = (metadata.get("error_type") or "").lower()
    combined = f"{error_code} {error_type}"
    for code in _NON_RETRYABLE_RATE_LIMIT_CODES:
        if code in combined:
            return code, False
    if "spend_limit" in combined:
        return "spend_limit_exceeded", False
    if "credit" in combined and ("exhaust" in combined or "balance" in combined):
        return "credit_balance_exhausted", False
    if "quota" in combined:
        return "insufficient_quota", False
    if "token" in combined:
        return "tpm", True
    if "request" in combined or "rpm" in combined:
        return "rpm", True
    if metadata.get("tokens_remaining") == "0":
        return "tpm", True
    if metadata.get("requests_remaining") == "0":
        return "rpm", True
    return "burst_or_unknown", True


def _safe_openai_error_metadata(exc: Exception) -> Dict[str, Any]:
    """Return only whitelisted API-error metadata suitable for logs and metrics."""
    response = getattr(exc, "response", None)
    headers = getattr(response, "headers", None) or {}
    metadata: Dict[str, Any] = {
        "http_status": getattr(exc, "status_code", None),
        "error_class": type(exc).__name__,
        "error_code": _safe_error_text(getattr(exc, "code", None)),
        "error_type": _safe_error_text(getattr(exc, "type", None)),
        "error_param": _safe_error_text(getattr(exc, "param", None)),
        "request_id": _safe_error_text(
            getattr(exc, "request_id", None) or headers.get("x-request-id")
        ),
        "retry_after_seconds": _retry_after_seconds(headers),
    }
    for header_name, field_name in _RATE_LIMIT_HEADER_FIELDS.items():
        header_value = _safe_error_text(headers.get(header_name), limit=64)
        if header_value is not None:
            metadata[field_name] = header_value
    if not isinstance(metadata["http_status"], int):
        metadata["http_status"] = None
    return {key: value for key, value in metadata.items() if value is not None}


def _get_invoice_model() -> str:
    model = os.getenv("OPENAI_INVOICE_MODEL", "").strip()
    if not model:
        raise RuntimeError("OPENAI_INVOICE_MODEL is not configured")
    return model


def _get_invoice_max_output_tokens() -> int:
    configured = os.getenv("OPENAI_INVOICE_MAX_OUTPUT_TOKENS", str(DEFAULT_INVOICE_MAX_OUTPUT_TOKENS)).strip()
    try:
        value = int(configured)
    except ValueError as exc:
        raise RuntimeError("OPENAI_INVOICE_MAX_OUTPUT_TOKENS must be a positive integer") from exc
    if value <= 0:
        raise RuntimeError("OPENAI_INVOICE_MAX_OUTPUT_TOKENS must be a positive integer")
    return value


def _get_invoice_timeout_seconds() -> int:
    configured = os.getenv("OPENAI_INVOICE_TIMEOUT_SECONDS", str(DEFAULT_INVOICE_TIMEOUT_SECONDS)).strip()
    try:
        value = int(configured)
    except ValueError as exc:
        raise RuntimeError("OPENAI_INVOICE_TIMEOUT_SECONDS must be a positive integer") from exc
    if value <= 0:
        raise RuntimeError("OPENAI_INVOICE_TIMEOUT_SECONDS must be a positive integer")
    return value


def _get_invoice_v2_fast_text_max_chars() -> int:
    configured = os.getenv(
        "INVOICE_V2_FAST_TEXT_MAX_CHARS",
        str(DEFAULT_INVOICE_V2_FAST_TEXT_MAX_CHARS),
    ).strip()
    try:
        value = int(configured)
    except ValueError as exc:
        raise RuntimeError("INVOICE_V2_FAST_TEXT_MAX_CHARS must be a positive integer") from exc
    if value < 1000:
        raise RuntimeError("INVOICE_V2_FAST_TEXT_MAX_CHARS must be at least 1000")
    return value


def _get_invoice_reasoning_effort(audit: bool = False) -> str:
    env_key = "OPENAI_INVOICE_AUDIT_REASONING_EFFORT" if audit else "OPENAI_INVOICE_REASONING_EFFORT"
    default = DEFAULT_INVOICE_AUDIT_REASONING_EFFORT if audit else DEFAULT_INVOICE_REASONING_EFFORT
    effort = os.getenv(env_key, default).strip().lower()
    if effort not in INVOICE_REASONING_EFFORTS:
        raise RuntimeError(f"{env_key} must be a supported reasoning effort")
    return effort


def _ocr_is_enabled() -> bool:
    """Avoid loading the heavyweight OCR model on constrained production plans."""
    configured = os.getenv("OCR_ENABLED", "").strip().lower()
    if configured in {"1", "true", "yes"}:
        return True
    if configured in {"0", "false", "no"}:
        return False
    return os.getenv("ENV", "").strip().lower() != "production"


def _production_ocr_limits():
    if os.getenv("ENV", "").strip().lower() != "production":
        return OCR_MAX_PAGES, PDF_OCR_ZOOM, OCR_MAX_DIM
    return (
        min(OCR_MAX_PAGES, int(os.getenv("OCR_PRODUCTION_MAX_PAGES", "1"))),
        min(PDF_OCR_ZOOM, float(os.getenv("OCR_PRODUCTION_MAX_ZOOM", "1.2"))),
        min(OCR_MAX_DIM, int(os.getenv("OCR_PRODUCTION_MAX_DIM", "1200"))),
    )


def _get_client() -> OpenAI:
    global _client
    if _client is not None:
        return _client

    if OpenAI is None or openai is None:
        raise RuntimeError("openai no esta instalado")

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY no configurada")

    logger.info("OpenAI SDK version: %s", openai.__version__)
    if httpx is not None:
        logger.info("httpx version: %s", httpx.__version__)

    _client = OpenAI(api_key=api_key)
    logger.info("OpenAI client retry policy: max_retries=%s", _client.max_retries)
    return _client


def extract_first_json_object(text: str) -> Optional[Dict[str, Any]]:
    if not text:
        return None
    cleaned = text.strip()
    cleaned = re.sub(r"```[a-zA-Z]*", "", cleaned)
    cleaned = cleaned.replace("```", "")
    start = cleaned.find("{")
    if start == -1:
        return None
    depth = 0
    end = None
    for idx in range(start, len(cleaned)):
        char = cleaned[idx]
        if char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                end = idx
                break
    if end is None:
        return None
    snippet = cleaned[start : end + 1]
    try:
        return json.loads(snippet)
    except json.JSONDecodeError:
        return None


def _extract_json(text: str) -> Dict[str, Any]:
    data = extract_first_json_object(text)
    return data if isinstance(data, dict) else {}


def _normalize_date(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    value = str(value).strip()
    day = month = year = None
    if re.match(r"\d{4}-\d{2}-\d{2}$", value):
        year, month, day = value.split("-")
    else:
        match = re.match(r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$", value)
        if match:
            day, month, year = match.groups()
            if len(year) == 2:
                year = f"20{year}"
        else:
            match = re.match(r"(\d{4})[/-](\d{1,2})[/-](\d{1,2})$", value)
            if not match:
                return None
            year, month, day = match.groups()
    try:
        parsed = date(int(year), int(month), int(day))
    except (TypeError, ValueError):
        return None
    if not 2000 <= parsed.year <= date.today().year + 2:
        return None
    return parsed.isoformat()


_ISSUE_DATE_EVIDENCE_PATTERN = re.compile(
    r"(?<!\d)(\d{1,2})([-/.])(\d{1,2})\2(\d{2})(?!\d)"
)


def _inspect_issue_date_evidence(evidence: Any) -> Dict[str, Optional[str]]:
    """Classify one field-specific issue-date evidence value without guessing.

    The evidence contract already associates this text with the invoice issue
    date. A two-digit day at most 12 remains ambiguous between day/month
    conventions, so callers must never use it to overwrite model output.
    """
    text = str(evidence or "").strip()
    matches = list(_ISSUE_DATE_EVIDENCE_PATTERN.finditer(text))
    if not matches:
        return {"status": "missing", "normalized_date": None}
    if len(matches) != 1:
        return {"status": "multiple", "normalized_date": None}

    day, _separator, month, year = matches[0].groups()
    # Reuse Ledged's established two-digit-year policy through _normalize_date.
    normalized_date = _normalize_date(f"{day}/{month}/{year}")
    if normalized_date is None:
        return {"status": "invalid", "normalized_date": None}
    if int(day) <= 12:
        return {"status": "ambiguous", "normalized_date": normalized_date}
    return {"status": "unambiguous", "normalized_date": normalized_date}


def _resolve_issue_date_from_evidence(
    model_issue_date: Any, evidence: Any
) -> Tuple[Optional[str], Optional[Dict[str, str]], Optional[str]]:
    """Correct only a contradictory model date with unambiguous date evidence."""
    normalized_model_date = _normalize_date(model_issue_date)
    evidence_details = _inspect_issue_date_evidence(evidence)
    evidence_date = evidence_details.get("normalized_date")
    evidence_status = evidence_details.get("status")

    if (
        evidence_status == "unambiguous"
        and normalized_model_date is not None
        and evidence_date is not None
        and normalized_model_date != evidence_date
    ):
        return (
            evidence_date,
            {
                "code": "issue_date_corrected_from_evidence",
                "original_model_issue_date": str(model_issue_date),
                "normalized_evidence_issue_date": evidence_date,
            },
            None,
        )
    if (
        evidence_status == "ambiguous"
        and normalized_model_date is not None
        and evidence_date is not None
        and normalized_model_date != evidence_date
    ):
        return normalized_model_date, None, "issue_date_evidence_ambiguous"
    return normalized_model_date, None, None


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


def extract_payment_terms_days(text: str) -> Optional[int]:
    if not text:
        return None
    match = re.search(
        r"RECIBO\s+(\d+)\s+DIAS\s+FECHA\s+FACTURA",
        text,
        flags=re.IGNORECASE,
    )
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            return None
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


def _find_due_dates_in_due_context(text: str) -> List[str]:
    if not text:
        return []
    due_keywords = [
        "fecha de vencimiento",
        "vencimiento",
        "vence el",
    ]
    dates: List[str] = []
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for idx, line in enumerate(lines):
        lowered = line.lower()
        if not any(keyword in lowered for keyword in due_keywords):
            continue
        for match in re.findall(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", line):
            normalized = _normalize_date(match)
            if normalized:
                dates.append(normalized)
        for match in re.findall(r"\d{4}[/-]\d{1,2}[/-]\d{1,2}", line):
            normalized = _normalize_date(match)
            if normalized:
                dates.append(normalized)
        found_in_block = bool(dates)
        for offset in range(1, 13):
            if idx + offset >= len(lines):
                break
            candidate_line = lines[idx + offset]
            candidate_lower = candidate_line.lower()
            if any(keyword in candidate_lower for keyword in due_keywords):
                break
            if not _is_payment_schedule_continuation_line(candidate_line):
                if found_in_block:
                    break
                continue
            matched_any = False
            for match in re.findall(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", candidate_line):
                normalized = _normalize_date(match)
                if normalized:
                    dates.append(normalized)
                    matched_any = True
            for match in re.findall(r"\d{4}[/-]\d{1,2}[/-]\d{1,2}", candidate_line):
                normalized = _normalize_date(match)
                if normalized:
                    dates.append(normalized)
                    matched_any = True
            if matched_any:
                found_in_block = True
    return sorted({d for d in dates if d})


def _is_payment_schedule_continuation_line(line: str) -> bool:
    if not line:
        return False
    lowered = line.lower()
    if re.search(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", line):
        return True
    if re.search(r"\d{4}[/-]\d{1,2}[/-]\d{1,2}", line):
        return True
    if re.search(r"\bES\d{2}\b", line):
        return True
    if re.search(r"CC\s*:", line, flags=re.IGNORECASE):
        return True
    if re.search(r"\d{1,3}\s*d[ií]as", lowered):
        return True
    if re.search(r"\d+[.,]\d{2}", line):
        return True
    schedule_keywords = (
        "vencimiento",
        "fecha de vencimiento",
        "fecha de pago",
        "fecha pago",
        "forma de pago",
        "importe",
        "pendiente",
        "domicili",
        "transfer",
        "banco",
        "iban",
        "giro",
        "cuota",
        "cliente",
        "pago",
        "pagos",
    )
    return any(keyword in lowered for keyword in schedule_keywords)


def _find_payment_dates_by_keywords(text: str, invoice_date_iso: Optional[str]) -> List[str]:
    if not text:
        return []
    dates: List[str] = []
    keywords = [
        "fecha de vencimiento",
        "vencimiento",
        "vence el",
        "fecha de pago",
        "fecha pago",
        "pago",
        "pagos",
        "cuota",
        "cuotas",
    ]
    for line in text.splitlines():
        lowered = line.lower()
        has_payment_context = any(keyword in lowered for keyword in keywords)
        if has_payment_context:
            for match in re.findall(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", line):
                normalized = _normalize_date(match)
                if normalized:
                    dates.append(normalized)
            for match in re.findall(r"\d{4}[/-]\d{1,2}[/-]\d{1,2}", line):
                normalized = _normalize_date(match)
                if normalized:
                    dates.append(normalized)
            day_matches = re.findall(r"(\d{1,3})\s*d[ií]as", lowered)
            if day_matches and invoice_date_iso:
                try:
                    base_date = date.fromisoformat(invoice_date_iso)
                except ValueError:
                    base_date = None
                if base_date:
                    for days_str in day_matches:
                        try:
                            days = int(days_str)
                        except ValueError:
                            continue
                        due_date = (base_date + timedelta(days=days)).isoformat()
                        dates.append(due_date)
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for idx, line in enumerate(lines):
        lowered = line.lower()
        if not any(keyword in lowered for keyword in keywords):
            continue
        found_in_block = False
        for offset in range(1, 12):
            if idx + offset >= len(lines):
                break
            candidate_line = lines[idx + offset]
            if any(keyword in candidate_line.lower() for keyword in keywords):
                break
            if not _is_payment_schedule_continuation_line(candidate_line):
                if found_in_block:
                    break
                continue
            for match in re.findall(r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", candidate_line):
                normalized = _normalize_date(match)
                if normalized:
                    dates.append(normalized)
                    found_in_block = True
            for match in re.findall(r"\d{4}[/-]\d{1,2}[/-]\d{1,2}", candidate_line):
                normalized = _normalize_date(match)
                if normalized:
                    dates.append(normalized)
                    found_in_block = True
    unique_dates = sorted({d for d in dates if d})
    return unique_dates


def _resolve_payment_schedule(
    extracted_text: str,
    invoice_date: Optional[str],
    raw_payment_dates: Any,
    single_payment_date_raw: Any,
    payment_terms_days_raw: Any,
    *,
    issue_date_corrected_from_evidence: bool = False,
) -> Tuple[List[str], Optional[int]]:
    payment_terms_days = _pick_first_non_empty(payment_terms_days_raw)
    try:
        payment_terms_days = int(payment_terms_days) if payment_terms_days is not None else None
    except (TypeError, ValueError):
        payment_terms_days = None
    if payment_terms_days is not None and payment_terms_days <= 0:
        payment_terms_days = None

    payment_dates: List[str] = []
    if isinstance(raw_payment_dates, list):
        for item in raw_payment_dates:
            normalized = _normalize_date(str(item)) if item is not None else None
            if normalized:
                payment_dates.append(normalized)
    elif isinstance(raw_payment_dates, str):
        for chunk in re.split(r"[;,]\s*", raw_payment_dates):
            normalized = _normalize_date(chunk.strip())
            if normalized:
                payment_dates.append(normalized)

    single_payment_date = _normalize_date(single_payment_date_raw)
    if single_payment_date:
        payment_dates.append(single_payment_date)

    explicit_due_dates = _find_due_dates_in_due_context(extracted_text)
    if explicit_due_dates:
        payment_dates = explicit_due_dates
    elif issue_date_corrected_from_evidence:
        # This literal term is a deterministic due-date rule. It supersedes a
        # model-derived installment only after the issue date itself has been
        # corrected from unambiguous, field-specific evidence.
        document_terms_days = extract_payment_terms_days(extracted_text)
        if document_terms_days is not None and invoice_date:
            try:
                payment_dates = [
                    (date.fromisoformat(invoice_date) + timedelta(days=document_terms_days)).isoformat()
                ]
                payment_terms_days = document_terms_days
            except ValueError:
                payment_dates = []
        else:
            text_payment_dates = _find_payment_dates_by_keywords(extracted_text, invoice_date)
            if text_payment_dates:
                payment_dates = text_payment_dates
    else:
        text_payment_dates = _find_payment_dates_by_keywords(extracted_text, invoice_date)
        if text_payment_dates:
            payment_dates = text_payment_dates

    if not payment_dates and payment_terms_days is None:
        payment_terms_days = extract_payment_terms_days(extracted_text)
        if payment_terms_days is not None and payment_terms_days <= 0:
            payment_terms_days = None

    if not payment_dates and payment_terms_days is not None and invoice_date:
        try:
            base_date = date.fromisoformat(invoice_date)
            payment_dates = [(base_date + timedelta(days=payment_terms_days)).isoformat()]
        except ValueError:
            payment_dates = []

    normalized_dates = sorted({d for d in payment_dates if d})
    if invoice_date:
        later_dates = [value for value in normalized_dates if value > invoice_date]
        # Invoice issue date is frequently returned by generic models as a
        # payment date. Keep it only for immediate-payment invoices with no
        # actual later due date in the document.
        if later_dates:
            normalized_dates = later_dates

    return normalized_dates, payment_terms_days


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


def _is_llm_amounts_trustworthy(
    base_amount: Optional[float],
    vat_rate: Optional[float],
    vat_amount: Optional[float],
    total_amount: Optional[float],
) -> bool:
    if base_amount is None or vat_rate is None or vat_amount is None or total_amount is None:
        return False
    if vat_rate < 0 or vat_rate > 30:
        return False
    if base_amount < 0 or vat_amount < 0 or total_amount < 0:
        return False
    if total_amount < base_amount:
        return False
    if abs(total_amount - (base_amount + vat_amount)) > 0.02:
        return False
    return True


def _pick_first_non_empty(*values: Any) -> Any:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and value.strip() == "":
            continue
        return value
    return None


def parse_eu_amount(value: Any) -> Optional[float]:
    if value is None:
        return None
    raw = str(value)
    raw = raw.replace("EUR", "").replace("€", "").replace("EURO", "").replace("EUROS", "")
    raw = raw.strip()
    raw = re.sub(r"\s+", "", raw)
    if not raw:
        return None
    sign = -1 if raw.startswith("-") else 1
    raw = raw.lstrip("+-")
    raw = raw.replace(".", "").replace(",", ".")
    raw = re.sub(r"[^\d.]", "", raw)
    if not raw:
        return None
    try:
        return sign * float(raw)
    except ValueError:
        return None


def _normalize_ocr_amount_text(text: str) -> str:
    if not text:
        return text
    text = unescape(text).replace("\xa0", " ")
    # Fix OCR patterns like "1,042 79" -> "1.042,79"
    return re.sub(r"(\d{1,3})[.,](\d{3})\s(\d{2})", r"\1.\2,\3", text)


def _normalize_amount(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    raw = str(value).replace("EUR", "").replace("euro", "").replace("€", "").strip()
    raw = raw.replace(" ", "")
    if re.match(r"^\d{1,3}(?:,\d{3})+\.\d{2}$", raw):
        try:
            return float(raw.replace(",", ""))
        except ValueError:
            return None
    if "," in raw:
        parsed = parse_eu_amount(raw)
        if parsed is not None:
            return parsed
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


def _normalize_vat_breakdown(raw_value: Any) -> List[Dict[str, Any]]:
    if not raw_value:
        return []
    value = raw_value
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except json.JSONDecodeError:
            return []
    if isinstance(value, dict):
        value = [value]
    if not isinstance(value, list):
        return []

    lines: List[Dict[str, Any]] = []
    for entry in value:
        if not isinstance(entry, dict):
            continue
        rate = _normalize_rate(
            entry.get("rate")
            or entry.get("vat_rate")
            or entry.get("vat")
            or entry.get("iva_rate")
            or entry.get("iva")
        )
        if rate is None or rate not in {0, 4, 10, 21}:
            continue
        base_amount = _normalize_amount(entry.get("base") or entry.get("base_amount"))
        vat_amount = _normalize_amount(entry.get("vat_amount") or entry.get("iva_amount"))
        total_amount = _normalize_amount(entry.get("total") or entry.get("total_amount"))
        if base_amount is None and total_amount is None:
            continue
        if base_amount is None and total_amount is not None:
            base_amount = round(total_amount / (1 + rate / 100), 2)
        if base_amount is not None and vat_amount is None:
            vat_amount = round(base_amount * (rate / 100), 2)
        if base_amount is not None and total_amount is None and vat_amount is not None:
            total_amount = round(base_amount + vat_amount, 2)
        lines.append(
            {
                "rate": float(rate),
                "base": _round_amount(base_amount),
                "vat_amount": _round_amount(vat_amount),
                "total": _round_amount(total_amount),
            }
        )
    return lines


def _extract_amounts_from_text(text: str) -> Dict[str, Optional[float]]:
    if not text:
        return {"base": None, "vat": None, "total": None}
    text = _normalize_ocr_amount_text(text)
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    def pick_best_amount(numbers: List[str]) -> Optional[float]:
        if not numbers:
            return None
        # Prefer amounts with explicit decimals.
        decimal_numbers = [
            n for n in numbers if re.search(r"[,.]\d{2}$", n.strip())
        ]
        candidates = decimal_numbers or numbers
        values = [_normalize_amount(value) for value in candidates]
        values = [value for value in values if value is not None]
        if not values:
            return None
        large_values = [value for value in values if value > 30]
        if large_values:
            return max(large_values)
        return values[-1]

    def math_consistent(
        base_amount: Optional[float],
        vat_amount: Optional[float],
        total_amount: Optional[float],
    ) -> bool:
        if base_amount is None or vat_amount is None or total_amount is None:
            return False
        difference = round((base_amount + vat_amount) - total_amount, 2)
        tolerance = max(0.05, total_amount * 0.01)
        return abs(difference) <= tolerance

    def reconcile_base_and_vat_from_total(
        base_amount: Optional[float],
        vat_amount: Optional[float],
        total_amount: Optional[float],
    ) -> Tuple[Optional[float], Optional[float]]:
        if total_amount is None or math_consistent(base_amount, vat_amount, total_amount):
            return base_amount, vat_amount

        currency_matches = re.findall(
            r"(-?\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})|-?\d+[.,]\d{2})\s*(?:EUR|€)",
            text,
            flags=re.IGNORECASE,
        )
        candidates: List[float] = []
        for raw in currency_matches:
            parsed = _normalize_amount(raw)
            if parsed is None or parsed < 0 or parsed > total_amount + 0.05:
                continue
            rounded = round(parsed, 2)
            if rounded not in candidates:
                candidates.append(rounded)

        pair_candidates: List[Tuple[int, float, float, float]] = []
        for left in candidates:
            for right in candidates:
                bigger = max(left, right)
                smaller = min(left, right)
                if bigger <= 0 or smaller < 0:
                    continue
                if abs((bigger + smaller) - total_amount) > 0.05:
                    continue
                implied_rate = (smaller / bigger) * 100 if bigger else None
                rate_penalty = 1
                if implied_rate is not None and any(
                    abs(implied_rate - standard) <= 0.5 for standard in (0.0, 4.0, 10.0, 21.0)
                ):
                    rate_penalty = 0
                pair_candidates.append((rate_penalty, -bigger, smaller, bigger))

        if not pair_candidates:
            return base_amount, vat_amount

        pair_candidates.sort()
        _, _, best_vat, best_base = pair_candidates[0]
        return best_base, best_vat

    def find_amount_for_keywords(
        keywords: List[str],
        *,
        forbid_if_contains: Optional[List[str]] = None,
        require_currency_on_keyword_line: bool = False,
        require_single_amount: bool = False,
        lookahead: int = 7,
        lookbehind: int = 0,
        ignore_inline_percentages: bool = False,
        require_keyword_at_line_start: bool = False,
        search_previous_before_next: bool = False,
        skip_previous_if_prior_label_contains: Optional[List[str]] = None,
    ) -> Optional[float]:
        for idx, line in enumerate(lines):
            upper = line.upper()
            if any(keyword in upper for keyword in keywords):
                if require_keyword_at_line_start:
                    if not any(upper.startswith(keyword) for keyword in keywords):
                        continue
                if forbid_if_contains and any(token in upper for token in forbid_if_contains):
                    continue
                amount = None
                numbers = []
                if not (ignore_inline_percentages and "%" in line):
                    numbers = re.findall(r"\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})|\d+[.,]\d{2}", line)
                if require_currency_on_keyword_line and not numbers:
                    if "€" not in line and "EUR" not in upper:
                        continue
                if require_single_amount and len(numbers) > 1:
                    continue
                if numbers:
                    amount = pick_best_amount(numbers)
                if amount is None:
                    # OCR may drop decimal separators; try to rebuild from plain digits.
                    raw_digits = re.findall(r"\b\d{4,6}\b", line)
                    if raw_digits:
                        candidate = raw_digits[-1]
                        amount = parse_eu_amount(candidate[:-2] + "," + candidate[-2:])
                if amount is None and search_previous_before_next and lookbehind > 0:
                    for offset in range(1, lookbehind + 1):
                        if idx - offset < 0:
                            break
                        if (
                            skip_previous_if_prior_label_contains
                            and idx - offset - 1 >= 0
                            and any(
                                token in lines[idx - offset - 1].upper()
                                for token in skip_previous_if_prior_label_contains
                            )
                        ):
                            continue
                        previous_line = lines[idx - offset]
                        numbers = re.findall(
                            r"\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})|\d+[.,]\d{2}",
                            previous_line,
                        )
                        if not numbers:
                            continue
                        amount = pick_best_amount(numbers)
                        if amount is not None:
                            break
                if amount is None and idx + 1 < len(lines):
                    candidates = []
                    for offset in range(1, lookahead + 1):
                        if idx + offset >= len(lines):
                            break
                        next_line = lines[idx + offset]
                        numbers = re.findall(
                            r"\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})|\d+[.,]\d{2}",
                            next_line,
                        )
                        if not numbers:
                            continue
                        has_currency = "€" in next_line or "EUR" in next_line.upper()
                        candidates.append((has_currency, numbers, next_line))
                        if has_currency:
                            amount = pick_best_amount(numbers)
                            break
                    if amount is None and candidates:
                        amount = pick_best_amount(candidates[0][1])
                if amount is None and lookbehind > 0 and not search_previous_before_next:
                    for offset in range(1, lookbehind + 1):
                        if idx - offset < 0:
                            break
                        previous_line = lines[idx - offset]
                        numbers = re.findall(
                            r"\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})|\d+[.,]\d{2}",
                            previous_line,
                        )
                        if not numbers:
                            continue
                        amount = pick_best_amount(numbers)
                        if amount is not None:
                            break
                if amount is not None:
                    return amount
        return None

    base_amount = (
        find_amount_for_keywords(
            ["BASE IMPONIBLE", "BASE IVA", "BASE I.V.A", "TOTAL BRUTO"],
            ignore_inline_percentages=True,
        )
        or find_amount_for_keywords(["SUBTOTAL"], lookahead=1, lookbehind=1)
        or find_amount_for_keywords(
            ["BASE"],
            ignore_inline_percentages=True,
            require_keyword_at_line_start=True,
        )
    )
    total_amount = (
        find_amount_for_keywords(
            ["TOTAL FACTURA"],
            forbid_if_contains=["BRUTO", "BASE", "IMPONIBLE", "I.V.A", "IVA", "REC.EQUIV"],
            require_single_amount=True,
            lookahead=0,
            lookbehind=1,
            search_previous_before_next=True,
            skip_previous_if_prior_label_contains=["I.V.A", "IVA", "BASE", "IMPONIBLE", "REC"],
        )
        or find_amount_for_keywords(
            ["TOTAL FACTURA"],
            forbid_if_contains=["BRUTO", "BASE", "IMPONIBLE", "I.V.A", "IVA", "REC.EQUIV"],
            require_single_amount=True,
            lookahead=4,
        )
        or find_amount_for_keywords(
            ["TOTAL IVA INCLUIDO", "TOTAL CON IVA", "NETO A PAGAR"],
            forbid_if_contains=["BRUTO", "BASE", "IMPONIBLE", "I.V.A", "IVA", "REC.EQUIV"],
            require_single_amount=True,
            lookahead=4,
        )
        or find_amount_for_keywords(
            ["TOTAL A PAGAR"],
            forbid_if_contains=["BRUTO", "BASE", "IMPONIBLE", "I.V.A", "IVA", "REC.EQUIV"],
            require_single_amount=True,
            lookahead=2,
            lookbehind=1,
        )
        or find_amount_for_keywords(
            ["TOTAL EUR"],
            forbid_if_contains=["BRUTO", "BASE", "IMPONIBLE", "I.V.A", "IVA", "REC.EQUIV"],
            require_single_amount=True,
            lookahead=3,
        )
        or find_amount_for_keywords(
            ["TOTAL"],
            forbid_if_contains=["BRUTO", "BASE", "IMPONIBLE", "I.V.A", "IVA", "REC.EQUIV"],
            require_single_amount=True,
            require_currency_on_keyword_line=True,
            lookahead=1,
        )
    )
    if total_amount is None:
        currency_matches = re.findall(
            r"(\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})|\d+[.,]\d{2})\s*(?:EUR|€)",
            text,
            flags=re.IGNORECASE,
        )
        if currency_matches:
            total_amount = _normalize_amount(currency_matches[-1])
    vat_amount = find_amount_for_keywords(["I.V.A", "IVA"])
    base_amount, vat_amount = reconcile_base_and_vat_from_total(base_amount, vat_amount, total_amount)
    return {"base": base_amount, "vat": vat_amount, "total": total_amount}


def _extract_explicit_withholding_amount_from_text(text: str) -> Optional[float]:
    if not text:
        return None
    normalized_text = _normalize_ocr_amount_text(text)
    lines = [line.strip() for line in normalized_text.splitlines() if line.strip()]
    amount_pattern = re.compile(r"\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})|\d+[.,]\d{2}")
    stop_tokens = ("portes", "subtotal", "base", "iva", "impuesto", "factura", "pendiente")

    for idx, line in enumerate(lines):
        lowered = line.lower()
        # Documents from landlords frequently print the tax label as
        # "I.R.P.F." instead of the compact "IRPF" spelling.
        compact_label = re.sub(r"[^a-z0-9áéíóúüñ]", "", lowered)
        is_irpf_label = "irpf" in compact_label
        is_withholding_label = any(
            keyword in lowered for keyword in ("retención", "retencion", "withholding")
        ) or is_irpf_label
        if not is_withholding_label:
            continue

        # Ignore the base and percentage column headers when the PDF emits
        # each table column on a separate line. The bare IRPF column is the
        # actual withheld amount.
        line_amounts = [_normalize_amount(raw) for raw in amount_pattern.findall(line)]
        line_amounts = [value for value in line_amounts if value is not None]
        is_base_header = is_irpf_label and "base" in compact_label
        # An invoice can state both the rate and amount in one line, e.g.
        # "IRPF (15%) -339,62". Only discard a percentage label when it
        # contains no monetary value, as happens in column headers.
        is_rate_header = (
            is_irpf_label
            and ("%" in line or "porcentaje" in compact_label)
            and not line_amounts
        )
        repeated_irpf_labels = compact_label.count("irpf") > 1
        # A rate label can be followed by the withheld amount on the next OCR
        # line ("IRPF (15%)" then "-339,62"). Keep scanning in that case.
        if is_base_header and not repeated_irpf_labels:
            continue

        if line_amounts:
            return _round_amount(line_amounts[-1])

        for offset in range(1, 3):
            if idx + offset >= len(lines):
                break
            candidate = lines[idx + offset].strip()
            if not candidate:
                continue
            candidate_lower = candidate.lower()
            if any(token in candidate_lower for token in stop_tokens):
                break
            matches = amount_pattern.findall(candidate)
            if not matches:
                continue
            parsed = _normalize_amount(matches[-1])
            if parsed is not None:
                return _round_amount(parsed)
            break
    return None


def _apply_withholding_to_payable_total(
    base_amount: Optional[float],
    vat_amount: Optional[float],
    total_amount: Optional[float],
    withholding_amount: Optional[float],
) -> Optional[float]:
    """Keep supplier invoice totals net when a later tax summary supplied gross."""
    if (
        base_amount is None
        or vat_amount is None
        or total_amount is None
        or not withholding_amount
    ):
        return total_amount
    gross_total = round(base_amount + vat_amount, 2)
    if abs(total_amount - gross_total) <= 0.02:
        return round(gross_total - withholding_amount, 2)
    return total_amount


def _extract_explicit_vat_exemption_amount_from_text(text: str) -> Optional[float]:
    """Return the taxable base when a document explicitly states it is VAT exempt."""
    if not text:
        return None
    normalized_text = _normalize_ocr_amount_text(text)
    lowered = normalized_text.lower()
    exemption_markers = ("exento", "exenta", "exentos", "exentas", "sin iva", "no sujeto")
    if not any(marker in lowered for marker in exemption_markers):
        return None

    # Do not treat a mixed invoice as fully exempt when it also declares a
    # standard positive VAT rate elsewhere in the document.
    positive_vat_rate = re.compile(
        r"(?:i\s*\.?\s*v\s*\.?\s*a\.?|iva)[^\d]{0,16}(?:4|10|21)(?:[.,]0+)?\s*%",
        flags=re.IGNORECASE,
    )
    if positive_vat_rate.search(normalized_text):
        return None

    amount_pattern = re.compile(r"\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})|\d+[.,]\d{2}")
    lines = [line.strip() for line in normalized_text.splitlines() if line.strip()]
    for idx, line in enumerate(lines):
        line_lower = line.lower()
        if not any(marker in line_lower for marker in exemption_markers):
            continue
        for offset in range(0, 3):
            if idx + offset >= len(lines):
                break
            amounts = [_normalize_amount(raw) for raw in amount_pattern.findall(lines[idx + offset])]
            amounts = [amount for amount in amounts if amount is not None]
            if amounts:
                return _round_amount(amounts[-1])
    return None


def _extract_payroll_amount_by_label(text: str, labels: List[str]) -> Optional[float]:
    if not text:
        return None
    normalized_text = _normalize_ocr_amount_text(text)
    amount_pattern = re.compile(r"\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})|\d+[.,]\d{2}")
    lines = [line.strip() for line in normalized_text.splitlines() if line.strip()]
    for idx, line in enumerate(lines):
        lowered = line.lower()
        if not any(label in lowered for label in labels):
            continue
        candidates = [line]
        if idx + 1 < len(lines):
            candidates.append(lines[idx + 1])
        for candidate in candidates:
            matches = amount_pattern.findall(candidate)
            if matches:
                parsed = _normalize_amount(matches[-1])
                if parsed is not None:
                    return _round_amount(parsed)
    return None


def _extract_payroll_employee_name_from_text(text: str) -> Optional[str]:
    if not text:
        return None
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for idx, line in enumerate(lines):
        lowered = line.lower()
        if any(token in lowered for token in ("trabajador/a", "trabajador", "empleado/a", "empleado")):
            for offset in range(1, 3):
                if idx + offset >= len(lines):
                    break
                candidate = lines[idx + offset].strip(" :-")
                if not candidate:
                    continue
                if any(ch.isdigit() for ch in candidate):
                    continue
                if len(candidate.split()) >= 2 and candidate.upper() == candidate:
                    return candidate.title()
                if len(candidate.split()) >= 2:
                    return candidate
    uppercase_candidates = [
        line.strip(" :-")
        for line in lines
        if len(line.split()) >= 2 and line.upper() == line and not any(ch.isdigit() for ch in line)
    ]
    return uppercase_candidates[0].title() if uppercase_candidates else None


def _extract_payroll_period_from_text(text: str) -> Optional[str]:
    if not text:
        return None
    normalized_text = _normalize_ocr_amount_text(text)
    for line in normalized_text.splitlines():
        lowered = line.lower()
        if "periodo" not in lowered and "mens" not in lowered:
            continue
        dates = re.findall(r"\d{1,2}[./-]\d{1,2}[./-]\d{2,4}", line)
        if dates:
            normalized = _normalize_date(dates[-1].replace(".", "/"))
            if normalized:
                return normalized[:7]
        month_match = re.search(
            r"\b(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)[a-z]*\s+(\d{2,4})\b",
            lowered,
        )
        if month_match:
            month_key, year = month_match.groups()
            months = {
                "ene": "01",
                "feb": "02",
                "mar": "03",
                "abr": "04",
                "may": "05",
                "jun": "06",
                "jul": "07",
                "ago": "08",
                "sep": "09",
                "oct": "10",
                "nov": "11",
                "dic": "12",
            }
            year = f"20{year}" if len(year) == 2 else year
            return f"{year}-{months[month_key]}"
    return None


def _extract_payroll_fields_from_text(text: str) -> Dict[str, Any]:
    if not text:
        return {}
    gross_amount = _extract_payroll_amount_by_label(
        text,
        ["t. devengado", "total devengado", "rem. total", "rem total"],
    )
    deductions_amount = _extract_payroll_amount_by_label(
        text,
        ["t. a deducir", "total a deducir", "deducciones"],
    )
    net_amount = _extract_payroll_amount_by_label(
        text,
        ["liquido a percibir", "líquido a percibir", "neto a percibir"],
    )
    employer_cost_amount = _extract_payroll_amount_by_label(
        text,
        ["coste empresa", "costo empresa"],
    )
    if deductions_amount is None and gross_amount is not None and net_amount is not None:
        deductions_amount = _round_amount(max(gross_amount - net_amount, 0))
    employee_name = _extract_payroll_employee_name_from_text(text)
    payroll_period = _extract_payroll_period_from_text(text)
    return {
        "employee_name": employee_name,
        "gross_amount": gross_amount,
        "payroll_total_deductions_amount": deductions_amount,
        "payroll_net_amount": net_amount,
        "payroll_employer_cost_amount": employer_cost_amount,
        "payroll_period": payroll_period,
    }


def _extract_invoice_date_from_text(text: str) -> Optional[str]:
    if not text:
        return None
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for idx, line in enumerate(lines):
        lowered = line.lower()
        if "vencimiento" in lowered or "fecha de pago" in lowered:
            continue
        if "factura" in lowered or "invoice" in lowered:
            if any(token in lowered for token in ("total", "impuesto", "importe", "base")):
                continue
            # Some ERP PDFs print the issue date on a nearby line and leave
            # only the FACTURA label in the text layer.
            for offset in (0, -1, 1, -2, 2, -3, 3):
                candidate_idx = idx + offset
                if candidate_idx < 0 or candidate_idx >= len(lines):
                    continue
                candidate_line = lines[candidate_idx].lower()
                if "vencimiento" in candidate_line or "pago" in candidate_line:
                    continue
                found = _extract_first_date(lines[candidate_idx])
                if found:
                    return found
        if "factura" in lowered and "fecha" in lowered:
            if "dias" in lowered:
                continue
            if idx > 0 and "vencimiento" in lines[idx - 1].lower():
                continue
            found = _extract_first_date(line)
            if found:
                return found
        if lowered.startswith("fecha") or " fecha " in lowered:
            found = _extract_first_date(line)
            if found:
                return found
            for offset in range(1, 21):
                if idx + offset < len(lines):
                    candidate = _extract_first_date(lines[idx + offset])
                    if candidate:
                        return candidate
    return None


def _extract_numbered_tax_summary_from_text(text: str) -> Dict[str, Any]:
    if not text:
        return {"found": False}
    normalized_text = _normalize_ocr_amount_text(text)
    upper_text = normalized_text.upper()
    required_tokens = ("TOTAL A PAGAR", "TOTAL BASE", "TOTAL I.V.A")
    if not all(token in upper_text for token in required_tokens):
        return {"found": False}

    lines = [line.strip() for line in normalized_text.splitlines() if line.strip()]
    def parse_block(block: List[str]) -> Dict[str, Any]:
        candidates: List[float] = []
        amount_pattern = re.compile(r"\d{1,3}(?:[.\s]\d{3})*[.,]\d{2}|\d+[.,]\d{2}")
        for line in block:
            for raw in amount_pattern.findall(line):
                parsed = _normalize_amount(raw)
                if parsed is None or parsed < 0:
                    continue
                rounded = round(parsed, 2)
                if rounded not in candidates:
                    candidates.append(rounded)
        if len(candidates) < 3:
            return {"found": False}

        total_triplet: Optional[Tuple[float, float, float]] = None
        for total_candidate in sorted(candidates, reverse=True):
            for base_candidate in sorted([v for v in candidates if v <= total_candidate], reverse=True):
                vat_candidate = round(total_candidate - base_candidate, 2)
                if vat_candidate < 0:
                    continue
                if any(abs(v - vat_candidate) <= 0.02 for v in candidates):
                    total_triplet = (base_candidate, vat_candidate, total_candidate)
                    break
            if total_triplet:
                break
        if not total_triplet:
            return {"found": False}

        total_base, total_vat, total_amount = total_triplet
        base_candidates = [value for value in candidates if 0 < value < total_base + 0.02]
        line_candidates: List[Dict[str, Any]] = []
        for base_candidate in base_candidates:
            for rate in (21.0, 10.0, 4.0, 0.0):
                vat_candidate = round(base_candidate * (rate / 100), 2)
                if rate == 0.0:
                    vat_match = abs(vat_candidate) <= 0.02
                else:
                    vat_match = any(abs(v - vat_candidate) <= 0.02 for v in candidates)
                if not vat_match:
                    continue
                line_candidates.append(
                    {
                        "base": _round_amount(base_candidate),
                        "vat_amount": _round_amount(vat_candidate),
                        "rate": rate,
                        "total": _round_amount(base_candidate + vat_candidate),
                    }
                )

        breakdown: List[Dict[str, Any]] = []
        seen = set()
        for line in line_candidates:
            key = (line["base"], line["vat_amount"], line["rate"])
            if key not in seen:
                breakdown.append(line)
                seen.add(key)

        resolved_breakdown: List[Dict[str, Any]] = []
        for size in range(1, min(3, len(breakdown)) + 1):
            for combo in combinations(breakdown, size):
                combo_base = round(sum(line["base"] or 0 for line in combo), 2)
                combo_vat = round(sum(line["vat_amount"] or 0 for line in combo), 2)
                if abs(combo_base - total_base) <= 0.02 and abs(combo_vat - total_vat) <= 0.02:
                    resolved_breakdown = list(combo)
                    break
            if resolved_breakdown:
                break

        return {
            "found": True,
            "base_amount": total_base,
            "vat_amount": total_vat,
            "total_amount": total_amount,
            "vat_rate": resolved_breakdown[0]["rate"] if len(resolved_breakdown) == 1 else None,
            "source": "regex_tax_summary",
            "breakdown": resolved_breakdown,
        }

    start_indices = []
    for idx, line in enumerate(lines):
        upper = line.upper()
        if (
            "TOTAL A PAGAR" in upper
            or "TOTAL BASE" in upper
            or "BASE 1" in upper
            or "BASE 2" in upper
            or "BASE 3" in upper
        ):
            start_indices.append(idx)
    if not start_indices:
        return {"found": False}

    summaries = []
    for start_idx in start_indices:
        block = lines[start_idx : start_idx + 30]
        summary = parse_block(block)
        if summary.get("found"):
            summaries.append(summary)
    if not summaries:
        return {"found": False}

    summaries.sort(
        key=lambda summary: (
            len(summary.get("breakdown") or []),
            summary.get("total_amount") or 0,
            summary.get("base_amount") or 0,
        ),
        reverse=True,
    )
    return summaries[0]


def _extract_vertical_tax_summary_from_text(text: str) -> Dict[str, Any]:
    if not text:
        return {"found": False}
    normalized_text = _normalize_ocr_amount_text(text)
    upper_text = normalized_text.upper()
    required_tokens = ("BASE IMPUESTO", "TASA", "IMPORTE IMPUESTO")
    if not all(token in upper_text for token in required_tokens):
        return {"found": False}

    lines = [line.strip() for line in normalized_text.splitlines() if line.strip()]
    totals_labels = (
        "TOTAL BASE IMPONIBLE",
        "TOTAL IVA",
        "TOTAL FACTURA",
        "NETO A PAGAR",
        "TOTAL A PAGAR",
        "FACTURA DE IMPUESTOS",
        "IMPORTE IVA",
    )

    header_idx = None
    for idx, line in enumerate(lines):
        if "BASE IMPUESTO" in line.upper():
            header_idx = idx
            break
    if header_idx is None:
        return {"found": False}

    breakdown: List[Dict[str, Any]] = []
    idx = header_idx
    amount_pattern = re.compile(r"^\d{1,3}(?:[.\s]\d{3})*[.,]\d{2}|\d+[.,]\d{2}$")
    while idx < len(lines):
        line = lines[idx].strip()
        upper = line.upper()
        if upper == "TOTAL" or any(label in upper for label in totals_labels):
            break
        rate_candidate = _normalize_rate(line)
        if rate_candidate in {0, 4, 10, 21}:
            amounts: List[float] = []
            for offset in range(1, 8):
                if idx + offset >= len(lines):
                    break
                next_line = lines[idx + offset].strip()
                next_upper = next_line.upper()
                if any(label in next_upper for label in totals_labels):
                    break
                if _normalize_rate(next_line) in {0, 4, 10, 21} and amounts:
                    break
                if amount_pattern.match(next_line):
                    parsed = _normalize_amount(next_line)
                    if parsed is not None:
                        amounts.append(_round_amount(parsed))
                if len(amounts) >= 2:
                    break
            if len(amounts) >= 2:
                base_value, vat_value = amounts[0], amounts[1]
                breakdown.append(
                    {
                        "base": _round_amount(base_value),
                        "vat_amount": _round_amount(vat_value),
                        "rate": float(rate_candidate),
                        "total": _round_amount(base_value + vat_value),
                    }
                )
                idx += 1
                continue
        idx += 1

    if not breakdown:
        return {"found": False}

    summary_values = _summarize_vat_breakdown(breakdown)
    if not summary_values:
        return {"found": False}
    base_sum, vat_sum, total_sum = summary_values

    total_amount = None
    for idx, line in enumerate(lines):
        upper = line.upper()
        if "TOTAL FACTURA" in upper or "NETO A PAGAR" in upper or "TOTAL A PAGAR" in upper:
            for offset in range(1, 6):
                if idx + offset >= len(lines):
                    break
                candidate_line = lines[idx + offset].strip()
                matches = re.findall(r"\d{1,3}(?:[.\s]\d{3})*[.,]\d{2}|\d+[.,]\d{2}", candidate_line)
                if matches:
                    parsed = _normalize_amount(matches[-1])
                    if parsed is not None:
                        total_amount = _round_amount(parsed)
                        break
            if total_amount is not None:
                break

    if total_amount is None:
        currency_matches = re.findall(
            r"(\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})|\d+[.,]\d{2})\s*(?:EUR|€)",
            normalized_text,
            flags=re.IGNORECASE,
        )
        for raw_amount in reversed(currency_matches):
            candidate = _normalize_amount(raw_amount)
            if candidate is not None and abs(candidate - total_sum) <= 0.05:
                total_amount = candidate
                break

    if total_amount is None:
        total_amount = total_sum

    if abs((base_sum + vat_sum) - total_amount) > 0.05:
        return {"found": False}

    return {
        "found": True,
        "base_amount": base_sum,
        "vat_amount": vat_sum,
        "total_amount": _round_amount(total_amount),
        "vat_rate": breakdown[0]["rate"] if len(breakdown) == 1 else None,
        "source": "regex_tax_summary",
        "breakdown": breakdown,
    }


def _extract_multiline_tax_summary_from_block(block: List[str]) -> Dict[str, Any]:
    if not block:
        return {"found": False}

    amount_pattern = re.compile(r"^\d{1,3}(?:[.\s]\d{3})*[.,]\d{2}|\d+[.,]\d{2}$")
    marker_pattern = re.compile(r"^\(\d+\)$")
    breakdown: List[Dict[str, Any]] = []

    idx = 0
    while idx + 3 < len(block):
        base_line = block[idx].strip()
        marker_line = block[idx + 1].strip()
        rate_line = block[idx + 2].strip()
        vat_line = block[idx + 3].strip()

        if (
            amount_pattern.match(base_line)
            and marker_pattern.match(marker_line)
            and amount_pattern.match(rate_line)
            and amount_pattern.match(vat_line)
        ):
            base_value = _normalize_amount(base_line)
            rate_value = _normalize_rate(rate_line)
            vat_value = _normalize_amount(vat_line)
            if (
                base_value is not None
                and vat_value is not None
                and rate_value in {0, 4, 10, 21}
            ):
                breakdown.append(
                    {
                        "base": _round_amount(base_value),
                        "vat_amount": _round_amount(vat_value),
                        "rate": float(rate_value),
                        "total": _round_amount(base_value + vat_value),
                    }
                )
                idx += 4
                continue
        idx += 1

    if not breakdown:
        return {"found": False}

    base_sum = _round_amount(sum(line["base"] or 0 for line in breakdown))
    vat_sum = _round_amount(sum(line["vat_amount"] or 0 for line in breakdown))
    total_sum = _round_amount(base_sum + vat_sum)

    candidates: List[float] = []
    for line in block:
        for raw in re.findall(r"\d{1,3}(?:[.\s]\d{3})*[.,]\d{2}|\d+[.,]\d{2}", line):
            parsed = _normalize_amount(raw)
            if parsed is None:
                continue
            rounded = round(parsed, 2)
            if rounded not in candidates:
                candidates.append(rounded)

    total_base = base_sum
    total_amount = total_sum
    if any(abs(candidate - base_sum) <= 0.02 for candidate in candidates):
        total_base = base_sum
    if any(abs(candidate - total_sum) <= 0.02 for candidate in candidates):
        total_amount = total_sum
    else:
        explicit_total = next(
            (candidate for candidate in sorted(candidates, reverse=True) if candidate > base_sum + 0.02),
            None,
        )
        if explicit_total is not None and abs((explicit_total - base_sum) - vat_sum) <= 0.05:
            total_amount = explicit_total

    total_vat = _round_amount(total_amount - total_base)
    if abs(total_vat - vat_sum) > 0.05:
        return {"found": False}

    return {
        "found": True,
        "base_amount": total_base,
        "vat_amount": total_vat,
        "total_amount": total_amount,
        "vat_rate": breakdown[0]["rate"] if len(breakdown) == 1 else None,
        "source": "regex_tax_summary",
        "breakdown": breakdown,
    }


def _extract_tax_summary_from_text(text: str) -> Dict[str, Any]:
    if not text:
        return {"found": False}
    text = _normalize_ocr_amount_text(text)
    vertical_summary = _extract_vertical_tax_summary_from_text(text)
    if vertical_summary.get("found"):
        return vertical_summary
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return {"found": False}
    start_idx = None
    for idx, line in enumerate(lines):
        if "IMPUESTOS" in line.upper():
            start_idx = idx
            break
    if start_idx is None:
        for idx, line in enumerate(lines):
            if "BASE IMPONIBLE" in line.upper():
                start_idx = idx
                break
    if start_idx is None:
        return _extract_numbered_tax_summary_from_text(text)
    block = lines[start_idx : start_idx + 20]
    multiline_summary = _extract_multiline_tax_summary_from_block(block)
    if multiline_summary.get("found"):
        return multiline_summary

    def find_amount_after_keywords(
        keywords: List[str],
        *,
        forbid_if_contains: Optional[List[str]] = None,
        prefer_last: bool = False,
        ignore_inline_percentages: bool = False,
    ) -> Tuple[Optional[float], Optional[str]]:
        for idx, line in enumerate(block):
            upper = line.upper()
            if any(keyword in upper for keyword in keywords):
                if forbid_if_contains and any(token in upper for token in forbid_if_contains):
                    continue
                candidates: List[str] = []
                for offset in range(0, 8):
                    if idx + offset >= len(block):
                        break
                    next_line = block[idx + offset]
                    if offset == 0 and ignore_inline_percentages and "%" in next_line:
                        continue
                    matches = _EU_AMOUNT_RE.findall(next_line)
                    if matches:
                        candidates.extend(matches)
                if candidates:
                    raw_value = candidates[-1] if prefer_last else candidates[0]
                    return parse_eu_amount(raw_value), raw_value
        return None, None

    rate_value = None
    rate_raw = None
    for line in block:
        if "IVA" in line.upper() or "%" in line:
            match = re.search(r"(\d{1,2}(?:[.,]\d{1,2})?)\s*%?", line)
            if match:
                candidate_rate = _normalize_rate(match.group(1))
                if candidate_rate in {0, 4, 10, 21}:
                    rate_value = candidate_rate
                    rate_raw = match.group(1)
                    break
    if rate_value is None:
        for line in block:
            stripped_line = line.strip()
            if stripped_line.startswith("(") and stripped_line.endswith(")"):
                continue
            candidate_line = stripped_line.strip("()")
            if "/" in candidate_line:
                continue
            if re.match(r"^\d{1,2}(?:[.,]\d{1,2})?$", candidate_line):
                candidate_rate = _normalize_rate(candidate_line)
                if candidate_rate in {0, 4, 10, 21}:
                    rate_value = candidate_rate
                    rate_raw = candidate_line
                    break

    base_value, base_raw = find_amount_after_keywords(
        ["BASE IMPONIBLE", "BASE IVA", "BASE I.V.A", "B.IMPON", "BASE"],
        forbid_if_contains=["TOTAL", "IVA", "I.V.A"],
        prefer_last=False,
        ignore_inline_percentages=True,
    )
    vat_value, vat_raw = find_amount_after_keywords(
        ["I.V.A", "IVA"],
        forbid_if_contains=["REC", "RECARGO"],
        prefer_last=False,
        ignore_inline_percentages=True,
    )
    total_value, total_raw = find_amount_after_keywords(
        ["TOTAL"],
        forbid_if_contains=["BRUTO", "IMPONIBLE", "I.V.A", "IVA", "REC"],
        prefer_last=True,
    )

    amount_candidates: List[Tuple[float, str]] = []
    for line in block:
        for raw in _EU_AMOUNT_RE.findall(line):
            parsed = parse_eu_amount(raw)
            if parsed is None:
                continue
            amount_candidates.append((parsed, raw))
    if rate_value is None and amount_candidates:
        for candidate, raw in amount_candidates:
            normalized_candidate = _normalize_rate(candidate)
            if normalized_candidate in {0, 4, 10, 21}:
                rate_value = normalized_candidate
                rate_raw = raw
                break
    if rate_value is not None:
        amount_candidates = [
            (value, raw)
            for (value, raw) in amount_candidates
            if abs(value - rate_value) > 0.1
        ]

    if rate_value is not None and amount_candidates:
        for base_candidate, base_candidate_raw in amount_candidates:
            if base_candidate <= 0:
                continue
            expected_vat = round(base_candidate * (rate_value / 100), 2)
            vat_match = None
            for vat_candidate, vat_candidate_raw in amount_candidates:
                if abs(vat_candidate - expected_vat) <= 0.05:
                    vat_match = (vat_candidate, vat_candidate_raw)
                    break
            if vat_match:
                computed_total = round(base_candidate + vat_match[0], 2)
                total_match = None
                for total_candidate, total_candidate_raw in amount_candidates:
                    if abs(total_candidate - computed_total) <= 0.05:
                        total_match = (total_candidate, total_candidate_raw)
                        break
                if base_value is None:
                    base_value, base_raw = base_candidate, base_candidate_raw
                if vat_value is None:
                    vat_value, vat_raw = vat_match
                if total_value is None:
                    total_value, total_raw = (
                        total_match if total_match else (computed_total, None)
                    )
                break

    if base_value is not None and rate_value is not None:
        expected_vat = round(base_value * (rate_value / 100), 2)
        if vat_value is None or abs(vat_value - expected_vat) > 0.05:
            vat_value = expected_vat
        expected_total = round(base_value + (vat_value or 0), 2)
        if total_value is None or abs(total_value - expected_total) > 0.05:
            total_value = expected_total
    if base_value is not None and total_value is not None:
        computed_vat = round(total_value - base_value, 2)
        if vat_value is None or abs(vat_value - computed_vat) > 0.05:
            vat_value = computed_vat
        if rate_value is None and base_value > 0:
            rate_value = round(100 * vat_value / base_value, 2)

    if base_value is not None and vat_value is None and rate_value is not None:
        vat_value = round(base_value * (rate_value / 100), 2)
    if base_value is not None and vat_value is not None and total_value is None:
        total_value = round(base_value + vat_value, 2)
    if base_value is not None and total_value is not None and vat_value is None:
        vat_value = round(total_value - base_value, 2)
    if rate_value is None and base_value and vat_value:
        if base_value > 0:
            rate_value = round(100 * vat_value / base_value, 2)

    found = any(value is not None for value in (base_value, vat_value, total_value))
    if found:
        return {
            "found": found,
            "base_amount": base_value,
            "vat_amount": vat_value,
            "total_amount": total_value,
            "vat_rate": rate_value,
            "base_raw": base_raw,
            "vat_raw": vat_raw,
            "total_raw": total_raw,
            "rate_raw": rate_raw,
            "source": "regex_tax_summary" if found else None,
        }
    numbered_summary = _extract_numbered_tax_summary_from_text(text)
    if numbered_summary.get("found"):
        return numbered_summary
    return {
        "found": found,
        "base_amount": base_value,
        "vat_amount": vat_value,
        "total_amount": total_value,
        "vat_rate": rate_value,
        "base_raw": base_raw,
        "vat_raw": vat_raw,
        "total_raw": total_raw,
        "rate_raw": rate_raw,
        "source": "regex_tax_summary" if found else None,
    }


def _apply_tax_summary_override(
    text: str,
    base_amount: Optional[float],
    vat_amount: Optional[float],
    total_amount: Optional[float],
    vat_rate: Optional[float],
    summary: Optional[Dict[str, Any]],
    withholding_amount: Optional[float] = None,
) -> Tuple[Optional[float], Optional[float], Optional[float], Optional[float], str]:
    source = "llm"
    if not summary or not summary.get("found"):
        return base_amount, vat_amount, total_amount, vat_rate, source

    summary_base = summary.get("base_amount")
    summary_vat = summary.get("vat_amount")
    summary_total = summary.get("total_amount")
    summary_rate = _pick_first_non_empty(summary.get("vat_rate"), vat_rate)
    summary_base_raw = summary.get("base_raw")
    summary_breakdown = summary.get("breakdown") if isinstance(summary.get("breakdown"), list) else []

    def within_tolerance(a: float, b: float) -> bool:
        return abs(a - b) <= 0.03

    if summary_base is not None and summary_total is not None:
        computed_vat = round(summary_total - summary_base, 2)
        if summary_vat is None or abs(summary_vat - computed_vat) > 0.05:
            summary_vat = computed_vat
    if summary_base is not None and summary_rate is not None and summary_vat is None:
        summary_vat = round(summary_base * (summary_rate / 100), 2)
    if summary_base is not None and summary_vat is not None and summary_total is None:
        summary_total = round(summary_base + summary_vat, 2)
    if summary_base is not None and summary_total is not None and summary_vat is None:
        summary_vat = round(summary_total - summary_base, 2)
    if summary_rate is None and summary_base and summary_vat and len(summary_breakdown) <= 1:
        summary_rate = round(100 * summary_vat / summary_base, 2)

    summary_math_ok = (
        summary_base is not None
        and summary_vat is not None
        and summary_total is not None
        and within_tolerance(summary_base + summary_vat, summary_total)
    )
    llm_math_ok = _validate_math(
        base_amount, vat_amount, total_amount, withholding_amount
    ).get("is_consistent")

    # A complete and coherent result read from the visual document is more
    # reliable than a flattened text table. Regex is a validator/fallback, not
    # the decision-maker for invoice amounts.
    if llm_math_ok:
        return base_amount, vat_amount, total_amount, vat_rate, source

    if (
        summary_base_raw
        and isinstance(summary_base_raw, str)
        and _EU_THOUSANDS_RE.match(summary_base_raw)
        and base_amount is not None
        and summary_base is not None
    ):
        diff = abs(summary_base - base_amount)
        if 150 <= diff <= 350:
            base_amount = summary_base
            vat_amount = summary_vat
            total_amount = summary_total
            vat_rate = summary_rate
            source = "regex_tax_summary"
            return base_amount, vat_amount, total_amount, vat_rate, source

    if summary_math_ok and (not llm_math_ok):
        base_amount = summary_base
        vat_amount = summary_vat
        total_amount = summary_total
        vat_rate = summary_rate
        source = "regex_tax_summary"
        return base_amount, vat_amount, total_amount, vat_rate, source

    if summary_math_ok:
        base_amount = summary_base
        vat_amount = summary_vat
        total_amount = summary_total
        vat_rate = summary_rate
        source = "regex_tax_summary"
        return base_amount, vat_amount, total_amount, vat_rate, source

    if summary_base is not None or summary_total is not None:
        base_amount = _pick_first_non_empty(summary_base, base_amount)
        vat_amount = _pick_first_non_empty(summary_vat, vat_amount)
        total_amount = _pick_first_non_empty(summary_total, total_amount)
        vat_rate = _pick_first_non_empty(summary_rate, vat_rate)
        source = "regex_tax_summary"
        return base_amount, vat_amount, total_amount, vat_rate, source

    return base_amount, vat_amount, total_amount, vat_rate, source


def _apply_text_amount_fallbacks(
    text: str,
    base_amount: Optional[float],
    vat_amount: Optional[float],
    total_amount: Optional[float],
    amount_source: Optional[str],
) -> Tuple[Optional[float], Optional[float], Optional[float], Optional[str]]:
    text_amounts = _extract_amounts_from_text(text)
    text_total = text_amounts.get("total")
    if text_total is not None and amount_source != "regex_tax_summary":
        text_total_math_ok = _validate_math(base_amount, vat_amount, text_total).get("is_consistent")
        current_math_ok = _validate_math(base_amount, vat_amount, total_amount).get("is_consistent")
        can_override_from_text = (
            total_amount is None
            or base_amount is None
            or vat_amount is None
            or bool(text_total_math_ok)
        )
        if can_override_from_text and (total_amount is None or text_total <= (total_amount + 0.02)):
            total_amount = text_total
            if amount_source == "llm":
                amount_source = "text_total"
        elif current_math_ok:
            text_total = None
    if text_amounts.get("base") is not None and base_amount is None:
        base_amount = text_amounts.get("base")
    if text_amounts.get("vat") is not None and vat_amount is None:
        vat_amount = text_amounts.get("vat")
    return base_amount, vat_amount, total_amount, amount_source


def _maybe_override_amounts_from_text(
    text: str,
    base_amount: Optional[float],
    vat_amount: Optional[float],
    total_amount: Optional[float],
    vat_rate: Optional[float],
) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    extracted = _extract_amounts_from_text(text)
    text_base = extracted.get("base")
    text_total = extracted.get("total")
    text_vat = extracted.get("vat")

    def is_significantly_different(a: Optional[float], b: Optional[float]) -> bool:
        if a is None or b is None:
            return False
        tolerance = max(0.05, b * 0.01)
        return abs(a - b) > tolerance

    math_ok = _validate_math(base_amount, vat_amount, total_amount).get("is_consistent")
    text_math_ok = _validate_math(text_base, text_vat, text_total).get("is_consistent")
    if text_math_ok and (not math_ok or is_significantly_different(base_amount, text_base) or is_significantly_different(total_amount, text_total)):
        base_amount = text_base
        vat_amount = text_vat
        total_amount = text_total
        return base_amount, vat_amount, total_amount

    if vat_rate is not None and text_base is not None and text_vat is not None and text_total is None:
        computed_total = round(text_base + text_vat, 2)
        if _validate_math(text_base, text_vat, computed_total):
            base_amount = text_base
            vat_amount = text_vat
            total_amount = computed_total
            return base_amount, vat_amount, total_amount

    if text_base is not None and (base_amount is None or is_significantly_different(base_amount, text_base)):
        base_amount = text_base
        if vat_rate is not None:
            vat_amount = round(base_amount * (vat_rate / 100), 2)
            total_amount = round(base_amount + vat_amount, 2)

    if text_total is not None and (total_amount is None or (not math_ok and is_significantly_different(total_amount, text_total))):
        total_amount = text_total
        if base_amount is not None and vat_amount is None:
            vat_amount = round(total_amount - base_amount, 2)

    if not math_ok and base_amount is not None and vat_amount is not None and total_amount is not None:
        difference = round((base_amount + vat_amount) - total_amount, 2)
        if abs(difference) > 0.05:
            total_amount = round(base_amount + vat_amount, 2)

    return base_amount, vat_amount, total_amount


def _summarize_vat_breakdown(lines: List[Dict[str, Any]]) -> Optional[Tuple[float, float, float]]:
    if not lines:
        return None
    base_total = 0.0
    vat_total = 0.0
    total_total = 0.0
    for line in lines:
        base_total += float(line.get("base") or 0)
        vat_total += float(line.get("vat_amount") or 0)
        total_total += float(line.get("total") or 0)
    return (_round_amount(base_total), _round_amount(vat_total), _round_amount(total_total))


def _reconcile_vat_breakdown(
    vat_breakdown: List[Dict[str, Any]],
    base_amount: Optional[float],
    vat_amount: Optional[float],
    total_amount: Optional[float],
    vat_rate: Optional[float],
    source: Optional[str],
) -> List[Dict[str, Any]]:
    if not vat_breakdown:
        return []
    if base_amount is None or vat_amount is None or total_amount is None:
        return vat_breakdown
    summary = _summarize_vat_breakdown(vat_breakdown)
    if not summary:
        return vat_breakdown
    base_sum, vat_sum, total_sum = summary
    if (
        abs((base_sum or 0) - base_amount) <= 0.05
        and abs((vat_sum or 0) - vat_amount) <= 0.05
        and abs((total_sum or 0) - total_amount) <= 0.05
    ):
        return vat_breakdown
    if source == "regex_tax_summary" and vat_rate is not None:
        return [
            {
                "rate": float(vat_rate),
                "base": _round_amount(base_amount),
                "vat_amount": _round_amount(vat_amount),
                "total": _round_amount(total_amount),
            }
        ]
    return vat_breakdown


def _extract_vat_breakdown_from_text(text: str) -> List[Dict[str, Any]]:
    if not text:
        return []
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return []
    breakdown: List[Dict[str, Any]] = []
    context_indices = set()
    keywords = [
        "base iva",
        "base i.v.a",
        "base i.v.a.",
        "iva",
        "i.v.a",
        "i.v.a.",
        "%iva",
        "% iva",
    ]
    for idx, line in enumerate(lines):
        lowered = line.lower()
        if any(keyword in lowered for keyword in keywords):
            context_indices.update({idx, idx + 1, idx + 2})
    if not context_indices:
        context_indices = set(range(min(6, len(lines))))

    number_pattern = re.compile(r"\d{1,6}[\\.,]\\d{2}")
    rate_pattern = re.compile(r"\b(\d{1,2}(?:[\\.,]\d{1,2})?)\s*%?")

    for idx, line in enumerate(lines):
        if idx not in context_indices:
            continue
        if "cliente" in line.lower() or "facturado a" in line.lower():
            continue
        numbers = number_pattern.findall(line)
        if len(numbers) < 2:
            continue
        floats = [_normalize_amount(value) for value in numbers]
        floats = [value for value in floats if value is not None]
        if len(floats) < 2:
            continue
        rates = []
        for match in rate_pattern.findall(line):
            rate_value = _normalize_rate(match)
            if rate_value is not None and rate_value in {0, 4, 10, 21}:
                rates.append(rate_value)
        rate = rates[0] if rates else None
        if rate is None:
            for value in floats:
                if value in {0, 4, 10, 21}:
                    rate = value
                    break
        if rate is None:
            continue
        if len(floats) >= 3:
            base_value = floats[0]
            vat_value = floats[2] if floats[1] == rate else floats[1]
        else:
            base_value = floats[0]
            vat_value = round(base_value * (rate / 100), 2)
        total_value = round(base_value + vat_value, 2)
        breakdown.append(
            {
                "rate": float(rate),
                "base": _round_amount(base_value),
                "vat_amount": _round_amount(vat_value),
                "total": _round_amount(total_value),
            }
        )
    return breakdown


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
        normalized_name = _normalize_entity_name(name)
        if normalized == normalized_name:
            return True
        if normalized_name:
            shorter, longer = sorted(
                [normalized, normalized_name],
                key=len,
            )
            if len(shorter) >= 5 and shorter in longer:
                return True
        if (
            normalized_name
            and len(normalized) >= 10
            and len(normalized_name) >= 10
            and SequenceMatcher(None, normalized, normalized_name).ratio() >= 0.96
        ):
            return True
    return False


def looks_like_person(name: Optional[str]) -> bool:
    if not name:
        return False
    if any(token in name.lower() for token in ("www.", "@", ".com", ".es", ".net", "http")):
        return False
    cleaned = re.sub(r"[^\w\s]", " ", name).strip()
    if not cleaned:
        return False
    tokens = [token for token in cleaned.split() if token.isalpha()]
    # Spanish invoices often prefix an autonomous professional's name with
    # abbreviations such as "Mª". They are honorifics, not name components.
    honorific_tokens = {"m", "mª", "ma", "d", "dona", "don", "sr", "sra", "srta"}
    tokens = [token for token in tokens if token.lower() not in honorific_tokens]
    if not tokens:
        return False
    if has_legal_form(name):
        return False
    stop_tokens = {
        "www",
        "es",
        "com",
        "net",
        "valencia",
        "madrid",
        "barcelona",
        "girona",
        "lugo",
        "españa",
        "spain",
    }
    lowered_tokens = [token.lower() for token in tokens]
    if any(token in stop_tokens for token in lowered_tokens):
        return False
    if len(set(lowered_tokens)) < len(lowered_tokens):
        return False
    long_tokens = [token for token in tokens if len(token) >= 3]
    if len(tokens) in {2, 3} and all(len(token) > 1 for token in tokens) and len(long_tokens) >= 2:
        return True
    return False


def has_legal_form(name: Optional[str]) -> bool:
    if not name:
        return False
    patterns = [
        r"(^|[^A-Z0-9])S\.?\s*L\.?\s*U\.?($|[^A-Z0-9])",
        r"(^|[^A-Z0-9])S\.?\s*A\.?\s*U\.?($|[^A-Z0-9])",
        r"(^|[^A-Z0-9])S\.?\s*L\.?\s*L\.?($|[^A-Z0-9])",
        r"(^|[^A-Z0-9])S\.?\s*L\.?\s*P\.?($|[^A-Z0-9])",
        r"(^|[^A-Z0-9])S\.?\s*C\.?\s*P\.?($|[^A-Z0-9])",
        r"(^|[^A-Z0-9])S\.?\s*R\.?\s*L\.?($|[^A-Z0-9])",
        r"(^|[^A-Z0-9])S\.?\s*A\.?($|[^A-Z0-9])",
        r"(^|[^A-Z0-9])S\.?\s*L\.?($|[^A-Z0-9])",
        r"(^|[^A-Z0-9])S\.?\s*C\.?($|[^A-Z0-9])",
        r"\bSCOOP\b",
        r"\bCOOP\b",
        r"\bCOOPERATIVA\b",
        r"\bAIE\b",
        r"\bUTE\b",
        r"\bCB\b",
        r"\bLIMITED\b",
        r"\bLTD\b",
        r"\bINC\b",
        r"\bGMBH\b",
        r"\bSARL\b",
        r"\bBV\b",
        r"\bNV\b",
        r"\bSAS\b",
        r"\bSRL\b",
    ]
    upper = str(name).upper()
    return any(re.search(pattern, upper) for pattern in patterns)


def _trim_company_name(value: Optional[str]) -> str:
    if not value:
        return ""
    value = _strip_inline_tax_id(value)
    value = re.sub(r"\s{2,}", " ", value).strip(" -–—|,;:·")
    legal_form_pattern = (
        r"(?:"
        r"S\.?\s*L\.?\s*U\.?|"
        r"S\.?\s*A\.?\s*U\.?|"
        r"S\.?\s*L\.?\s*L\.?|"
        r"S\.?\s*L\.?\s*P\.?|"
        r"S\.?\s*C\.?\s*P\.?|"
        r"S\.?\s*R\.?\s*L\.?|"
        r"S\.?\s*A\.?|"
        r"LIMITED|LTD|INC|GMBH|SARL|BV|NV|SAS|COOPERATIVA|COOP"
        r")"
    )
    segment_candidates: List[str] = []
    segments = [value] + [part.strip() for part in re.split(r"[,;|]", value) if part.strip()]
    for segment in segments:
        match = re.search(rf"(.+?\b{legal_form_pattern}\b)", segment, flags=re.IGNORECASE)
        if not match:
            continue
        candidate = re.sub(r"\s{2,}", " ", match.group(1)).strip(" -–—|,;:·")
        if candidate:
            segment_candidates.append(candidate)
    if segment_candidates:
        preferred = [
            candidate for candidate in segment_candidates if not _looks_like_legal_or_footer_text(candidate)
        ]
        value = min(preferred or segment_candidates, key=len)
    value = re.sub(r"\s{2,}", " ", value).strip(" -–—|,;:·")
    return value


def _looks_like_legal_or_footer_text(value: Optional[str]) -> bool:
    if not value:
        return False
    lowered = value.lower().strip()
    if not lowered:
        return False
    blocked_phrases = [
        "reglamento (ue)",
        "protección de datos",
        "proteccion de datos",
        "obligaciones legales",
        "agencia española de protección de datos",
        "agencia espanola de proteccion de datos",
        "derechos de acceso",
        "quedarán incorporados",
        "quedaran incorporados",
        "serán tratados",
        "seran tratados",
        "datos proporcionados",
        "impreso por sap business one",
        "impreso por",
    ]
    if any(phrase in lowered for phrase in blocked_phrases):
        return True
    words = [token for token in re.split(r"\s+", lowered) if token]
    if len(words) > 18:
        return True
    if len(lowered) > 140:
        return True
    if len(words) > 8 and lowered[0].islower():
        return True
    return False


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
        "impreso por",
        "datos del emisor",
        "datos de emisor",
        "datos del proveedor",
        "datos del cliente",
        "datos de cliente",
    ]
    return any(keyword in lowered for keyword in forbidden)


def _looks_like_address_line(value: Optional[str]) -> bool:
    if not value:
        return False
    normalized = re.sub(r"\s{2,}", " ", value).strip()
    if not normalized:
        return False
    lowered = normalized.lower()
    street_tokens = [
        "calle ",
        "c/",
        "c/ ",
        "cl. ",
        "cl ",
        "avda",
        "avenida",
        "paseo",
        "plaza",
        "camino",
        "carretera",
        "ronda",
        "barrio",
        "poligono",
        "polígono",
        "urbanizacion",
        "urbanización",
    ]
    location_tokens = [
        "codigo postal",
        "código postal",
        "población",
        "poblacion",
        "provincia",
        "valencia",
        "madrid",
        "barcelona",
        "girona",
        "lugo",
        "españa",
        "spain",
    ]
    starts_with_street = any(lowered.startswith(token) for token in street_tokens)
    has_postal_code = bool(re.search(r"\b\d{5}\b", normalized))
    many_digits = sum(char.isdigit() for char in normalized) >= 3
    has_location_token = any(token in lowered for token in location_tokens)
    address_number_pattern = bool(re.search(r"(n[ºo]\.?|num(?:ero)?\.?)\s*\d+", lowered))
    return (
        starts_with_street
        or address_number_pattern
        or (has_postal_code and (many_digits or has_location_token))
    )


def _is_valid_client(
    candidate: Optional[str],
    company_names,
    text: Optional[str] = None,
) -> bool:
    if candidate is None:
        return False
    value = str(candidate).strip()
    if not value:
        return False
    if contains_forbidden_keyword(value):
        return False
    if _looks_like_metadata(value):
        return False
    if _looks_like_address_line(value):
        return False
    if _is_same_entity(value, company_names):
        return False
    has_form = has_legal_form(value)
    inline_tax = _has_tax_id(value) or _has_iban(value) or "iban" in value.lower()
    has_tax = bool(text and _supplier_has_near_tax_id_or_iban(text, value))
    if looks_like_person(value):
        return inline_tax or has_tax
    if not has_form and not (inline_tax or has_tax):
        # Allow non-legal-form clients only if tax id/IBAN is present.
        return False
    return True


def _is_valid_supplier(
    candidate: Optional[str],
    company_names,
    text: Optional[str] = None,
    require_tax_id: bool = True,
) -> bool:
    if candidate is None:
        return False
    value = str(candidate).strip()
    if not value:
        return False
    value = _trim_company_name(value)
    if not value:
        return False
    if contains_forbidden_keyword(value):
        return False
    if _looks_like_address_line(value):
        return False
    if _looks_like_legal_or_footer_text(value):
        return False
    has_form = has_legal_form(value)
    meaningful_tokens = re.findall(r"[A-Za-zÀ-ÿ]{2,}", value)
    if len(meaningful_tokens) < 2 and not has_form:
        return False
    inline_tax = _has_tax_id(value) or _has_iban(value) or "iban" in value.lower()
    has_tax = bool(text and _supplier_has_near_tax_id_or_iban(text, value))
    is_person = looks_like_person(value)
    if not has_form:
        # For non-legal entities, only accept:
        # - an inline tax id/IBAN, or
        # - an autonomous person's full name with nearby tax id.
        if not inline_tax and not (is_person and has_tax):
            return False
    if _is_same_entity(value, company_names):
        return False
    if _appears_in_customer_context(text, value):
        return False
    if require_tax_id and text and not has_form and not (has_tax or inline_tax):
        return False
    return True


def _is_valid_visual_supplier(
    candidate: Optional[str],
    company_names,
    evidence: Optional[str],
) -> bool:
    """Allow an issuer visible only in a document header or logo.

    Some supplier names are not embedded in a PDF text layer. They may still
    be unambiguous in the visual header (for example, a registered brand). We
    only accept this narrower path when the model returns literal evidence and
    the candidate cannot be the registered customer company.
    """
    if candidate is None or not evidence:
        return False
    value = _trim_company_name(str(candidate).strip())
    if not value or contains_forbidden_keyword(value):
        return False
    if _looks_like_address_line(value) or _looks_like_legal_or_footer_text(value):
        return False
    if _is_same_entity(value, company_names):
        return False
    normalized_value = _normalize_entity_name(value)
    normalized_evidence = _normalize_entity_name(str(evidence))
    meaningful_tokens = re.findall(r"[A-Za-zÀ-ÿ]{3,}", value)
    return bool(
        meaningful_tokens
        and normalized_value
        and normalized_value in normalized_evidence
    )


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
        "codigo",
        "código",
        "cod.",
        "referencia",
        "direccion",
        "dirección",
        "población",
        "poblacion",
        "e-mail",
        "email",
        "operación",
        "operacion",
        "farmacia",
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


def _has_iban(line: str) -> bool:
    if not line:
        return False
    patterns = [
        r"\bES\d{22}\b",
        r"\b[A-Z]{2}\d{2}[A-Z0-9]{11,30}\b",
    ]
    return any(re.search(pattern, line, re.IGNORECASE) for pattern in patterns)


def _supplier_has_near_tax_id_or_iban(text: str, supplier: str, window: int = 4) -> bool:
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
                if _has_tax_id(candidate) or _has_iban(candidate) or "iban" in candidate.lower():
                    return True
            for candidate in lines:
                normalized_candidate = _normalize_entity_name(candidate)
                if (
                    normalized_supplier
                    and normalized_candidate
                    and (
                        normalized_supplier in normalized_candidate
                        or normalized_candidate in normalized_supplier
                    )
                    and (_has_tax_id(candidate) or _has_iban(candidate) or "iban" in candidate.lower())
                ):
                    return True
            return False
    return False


def _appears_in_customer_context(
    text: Optional[str],
    candidate: Optional[str],
    lookahead: int = 6,
) -> bool:
    if not text or not candidate:
        return False
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    normalized_candidate = _normalize_entity_name(_trim_company_name(candidate))
    if not normalized_candidate:
        return False

    client_keywords = [
        "cliente",
        "datos cliente",
        "datos fiscales",
        "destinatario factura",
        "facturado a",
        "direccion de facturacion",
        "dirección de facturación",
        "datos de facturacion",
        "datos de facturación",
        "receptor",
        "enviado a",
        "bill to",
        "ship to",
    ]
    stop_keywords = [
        "factura nº",
        "factura no",
        "nº de factura",
        "artículo",
        "articulo",
        "material/",
        "concepto",
        "b.impon",
        "b. impon",
        "forma de pago",
    ]

    for idx, line in enumerate(lines):
        lowered = line.lower()
        if not any(keyword in lowered for keyword in client_keywords):
            continue
        end = min(len(lines), idx + lookahead + 1)
        for candidate_idx in range(idx, end):
            if candidate_idx > idx:
                candidate_lower = lines[candidate_idx].lower()
                if any(stop in candidate_lower for stop in stop_keywords):
                    break
            normalized_line = _normalize_entity_name(lines[candidate_idx])
            if not normalized_line:
                continue
            if (
                normalized_candidate == normalized_line
                or normalized_candidate in normalized_line
            ):
                return True
    return False


def _strip_inline_tax_id(value: str) -> str:
    if not value:
        return value
    cleaned = value
    cleaned = re.sub(
        r"(?:n[ºo]\s*)?(?:nif/cif/nie/vat|c\.?\s*i\.?\s*f\.?|n\.?\s*i\.?\s*f\.?|d\.?\s*n\.?\s*i\.?|v\.?\s*a\.?\s*t\.?|i\.?\s*v\.?\s*a\.?)\s*[:#-]?\s*",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )
    cleaned = re.sub(r"(cif|nif|dni|vat|iva)\s*[:#-]?\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b[A-HJ-NP-SUVW]\s?-?\d{7}\s?-?[0-9A-J]\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b\d{8}\s?-?[A-Z]\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b[A-Z]{2}\s?-?\d{6,12}\b", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+-\s+", " ", cleaned)
    cleaned = re.sub(r"\s{2,}", " ", cleaned)
    return cleaned.strip(" -–—|")


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
        "en nombre de",
        "issued by",
        "seller",
    ]
    client_keywords = [
        "cliente",
        "enviado a",
        "destinatario",
        "facturado a",
        "direccion de facturacion",
        "dirección de facturación",
        "datos de facturacion",
        "datos de facturación",
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
    client_context_lines = {
        idx for idx, line in enumerate(lines) if _appears_in_customer_context(text, line)
    }
    line_counts = {}
    for line in lines:
        key = _normalize_entity_name(line)
        if key:
            line_counts[key] = line_counts.get(key, 0) + 1

    candidates: List[Tuple[str, int]] = []
    for idx, line in enumerate(lines):
        lowered = line.lower()
        if idx in client_context_lines:
            continue
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
        if not _contains_legal_form(line) and not _has_tax_id(line) and not _has_iban(line):
            for offset in (1, 2):
                if idx + offset < len(lines):
                    neighbor = lines[idx + offset]
                    if _has_tax_id(neighbor) or _has_iban(neighbor):
                        score += 25
                        break
        if line_counts.get(_normalize_entity_name(line), 0) > 1:
            score += 10
        if any(keyword in lowered for keyword in supplier_keywords):
            score += 25

        if score <= 0:
            continue
        candidates.append((line, score))

    return candidates


def _supplier_candidate_priority(
    candidate: Optional[str],
    company_names,
    text: Optional[str] = None,
) -> Optional[int]:
    if candidate is None:
        return None
    value = _trim_company_name(str(candidate).strip())
    if not _is_valid_supplier(value, company_names, text, require_tax_id=True):
        return None
    if has_legal_form(value):
        return 0
    if looks_like_person(value):
        return 1
    return None


def _pick_best_supplier_candidate(
    candidates: List[Tuple[str, int, int]],
    company_names,
    text: Optional[str] = None,
) -> Optional[str]:
    ranked: List[Tuple[int, int, int, str]] = []
    seen = set()
    for candidate, score, position in candidates:
        cleaned = _trim_company_name(candidate)
        if not cleaned:
            continue
        normalized = _normalize_entity_name(cleaned)
        if not normalized or normalized in seen:
            continue
        priority = _supplier_candidate_priority(cleaned, company_names, text)
        if priority is None:
            continue
        ranked.append((priority, -score, position, cleaned))
        seen.add(normalized)
    if not ranked:
        return None
    ranked.sort()
    return ranked[0][3]


def _extract_client_candidates(text: str, company_names=None) -> List[Tuple[str, int]]:
    if not text:
        return []
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return []

    client_keywords = [
        "cliente",
        "facturado a",
        "direccion de facturacion",
        "dirección de facturación",
        "datos de facturacion",
        "datos de facturación",
        "destinatario",
        "receptor",
        "enviado a",
        "bill to",
        "ship to",
    ]
    supplier_keywords = [
        "expedido por",
        "emisor",
        "proveedor",
        "facturado por",
        "en nombre de",
        "issued by",
        "seller",
    ]
    operational_keywords = [
        "transporte",
        "envío",
        "envio",
        "logística",
        "logistica",
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
        if any(keyword in lowered for keyword in supplier_keywords):
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
            score += 10
        if _contains_legal_form(line):
            score += 70
        if _has_tax_id(line):
            score += 35
        if not _contains_legal_form(line) and not _has_tax_id(line) and not _has_iban(line):
            for offset in (1, 2):
                if idx + offset < len(lines):
                    neighbor = lines[idx + offset]
                    if _has_tax_id(neighbor) or _has_iban(neighbor):
                        score += 20
                        break
        if line_counts.get(_normalize_entity_name(line), 0) > 1:
            score += 8
        if any(keyword in lowered for keyword in client_keywords):
            score += 35

        if score <= 0:
            continue
        candidates.append((line, score))

    return candidates


def _select_best_supplier(text: str, company_names=None) -> Optional[str]:
    if not text:
        return None
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    supplier_keywords = [
        "nombre empresa",
        "expedido por",
        "emisor",
        "proveedor",
        "facturado por",
        "en nombre de",
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
        "en nombre de",
        "iban",
        "datos bancarios",
        "datos fiscales",
    ]
    candidate_pool: List[Tuple[str, int, int]] = []

    for line in lines:
        lowered = line.lower()
        if "en nombre de" in lowered:
            match = re.split(r"en nombre de", line, flags=re.IGNORECASE)
            if len(match) > 1:
                candidate = match[1].strip(" :-")
                candidate_pool.append((candidate, 120, 0))

    for idx, line in enumerate(lines):
        lowered = line.lower()
        if "nombre empresa" in lowered:
            for offset in (1, 2):
                if idx + offset < len(lines):
                    candidate = lines[idx + offset].strip()
                    candidate_pool.append((candidate, 160 - offset, idx + offset))
        if any(keyword in lowered for keyword in supplier_keywords):
            for keyword in supplier_keywords:
                if keyword in lowered:
                    parts = re.split(keyword, line, flags=re.IGNORECASE)
                    if len(parts) > 1:
                        candidate = parts[1].strip(" :-")
                        candidate_pool.append((candidate, 130, idx))
            for offset in (1, 2):
                if idx + offset < len(lines):
                    candidate = lines[idx + offset].strip()
                    candidate_pool.append((candidate, 125 - offset, idx + offset))

    for idx, line in enumerate(lines):
        lowered = line.lower()
        if any(anchor in lowered for anchor in anchor_keywords):
            parts = line.split(":", 1)
            if len(parts) > 1:
                candidate_pool.append((parts[1].strip(), 110, idx))
            for offset in (1, 2):
                if idx + offset < len(lines):
                    candidate = lines[idx + offset].strip()
                    candidate_pool.append((candidate, 105 - offset, idx + offset))

    for idx, (candidate, score) in enumerate(_extract_supplier_candidates(text, company_names)):
        candidate_pool.append((candidate, score, idx + 1000))

    legal_form_fragment = re.compile(
        r"^(?:S\.?\s*L\.?\s*U\.?|S\.?\s*A\.?\s*U\.?|S\.?\s*L\.?\s*L\.?|S\.?\s*L\.?\s*P\.?|S\.?\s*C\.?\s*P\.?|S\.?\s*R\.?\s*L\.?|S\.?\s*A\.?|LIMITED|LTD|INC|GMBH|SARL|BV|NV|SAS|COOPERATIVA|COOP)\b",
        flags=re.IGNORECASE,
    )
    for idx in range(len(lines) - 1):
        next_line = lines[idx + 1].strip()
        if not legal_form_fragment.search(next_line):
            continue
        combined = f"{lines[idx]} {next_line}".strip()
        candidate_pool.append((combined, 140, idx))

    return _pick_best_supplier_candidate(candidate_pool, company_names, text)


def _extract_supplier_from_text(text: str, company_names=None) -> Optional[str]:
    return _select_best_supplier(text, company_names)


def _select_best_client(text: str, company_names=None) -> Optional[str]:
    if not text:
        return None
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    client_keywords = [
        "cliente",
        "facturado a",
        "direccion de facturacion",
        "dirección de facturación",
        "datos de facturacion",
        "datos de facturación",
        "destinatario",
        "receptor",
        "enviado a",
        "bill to",
        "ship to",
    ]
    supplier_keywords = [
        "expedido por",
        "emisor",
        "proveedor",
        "facturado por",
        "en nombre de",
        "issued by",
        "seller",
    ]

    for idx, line in enumerate(lines):
        lowered = line.lower()
        if any(keyword in lowered for keyword in supplier_keywords):
            continue
        if any(keyword in lowered for keyword in client_keywords):
            parts = re.split(
                r"cliente|facturado a|direccion de facturacion|dirección de facturación|datos de facturacion|datos de facturación|destinatario|receptor|enviado a|bill to|ship to",
                line,
                flags=re.IGNORECASE,
            )
            if len(parts) > 1:
                candidate = parts[1].strip(" :-")
                if _is_valid_client(candidate, company_names, text):
                    return _strip_inline_tax_id(candidate)
            for offset in (1, 2):
                if idx + offset < len(lines):
                    candidate = lines[idx + offset].strip()
                    if _is_valid_client(candidate, company_names, text):
                        return _strip_inline_tax_id(candidate)

    candidates = _extract_client_candidates(text, company_names)
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[1], reverse=True)
    best, score = candidates[0]
    if _is_valid_client(best, company_names, text):
        return _strip_inline_tax_id(best)
    return None


def _extract_client_from_text(text: str, company_names=None) -> Optional[str]:
    return _select_best_client(text, company_names)


def _match_known_supplier(text: str, known_suppliers: Optional[List[str]], company_names=None) -> Optional[str]:
    if not text or not known_suppliers:
        return None
    normalized_text = _normalize_entity_name(text)
    if not normalized_text:
        return None
    for name in known_suppliers:
        if not name:
            continue
        cleaned = str(name).strip()
        if not cleaned:
            continue
        if _normalize_entity_name(cleaned) in normalized_text:
            if _is_valid_supplier(cleaned, company_names, text, require_tax_id=True):
                return cleaned
    return None


def _validate_math(
    base_amount: Optional[float],
    vat_amount: Optional[float],
    total_amount: Optional[float],
    withholding_amount: Optional[float] = None,
) -> Dict[str, Any]:
    if base_amount is None or vat_amount is None or total_amount is None:
        return {"is_consistent": None, "difference": None}
    withholding = max(float(withholding_amount or 0), 0)
    difference = round((base_amount + vat_amount - withholding) - total_amount, 2)
    # Invoice totals are exact monetary values. A percentage tolerance could
    # silently accept a wrong total on high-value invoices; only permit normal
    # rounding differences.
    tolerance = 0.05
    return {
        "is_consistent": abs(difference) <= tolerance,
        "difference": difference,
    }


def _round_amount(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return round(float(value), 2)


def _confidence_score_for_source(source: Optional[str]) -> Optional[float]:
    if not source:
        return None
    mapping = {
        "regex_tax_summary": 0.98,
        "llm": 0.85,
        "text_total": 0.90,
        "breakdown": 0.70,
        "fallback": 0.60,
    }
    return mapping.get(source)


def _review_reasons_for_invoice(
    *,
    document_type: str,
    supplier: Optional[str],
    base_amount: Optional[float],
    vat_amount: Optional[float],
    total_amount: Optional[float],
    withholding_amount: Optional[float],
    breakdown_warning: bool,
) -> List[str]:
    reasons = []
    if document_type != "income" and not supplier:
        reasons.append("No se ha identificado el proveedor con evidencia suficiente.")
    if base_amount is None or vat_amount is None or total_amount is None:
        reasons.append("Faltan importes fiscales necesarios para validar la factura.")
    elif not _validate_math(
        base_amount, vat_amount, total_amount, withholding_amount
    ).get("is_consistent"):
        reasons.append("Base, IVA, retención y total no cuadran entre sí.")
    if breakdown_warning:
        reasons.append("El desglose de IVA requiere comprobación manual.")
    return reasons


def _group_standard_vat_rate(rate: Optional[float], tolerance: float = 0.25) -> Optional[float]:
    if rate is None:
        return None
    for standard in (0.0, 4.0, 10.0, 21.0):
        if abs(rate - standard) <= tolerance:
            return standard
    return rate


def _adjust_breakdown_to_targets(
    lines: List[Dict[str, Any]],
    base_target: Optional[float],
    vat_target: Optional[float],
    total_target: Optional[float],
) -> List[Dict[str, Any]]:
    if not lines:
        return []
    base_sum = sum(line.get("base") or 0 for line in lines)
    vat_sum = sum(line.get("vat_amount") or 0 for line in lines)
    total_sum = base_sum + vat_sum
    if total_target is not None:
        if base_target is None and vat_target is None and total_sum > 0:
            factor = total_target / total_sum
            base_target = round(base_sum * factor, 2)
            vat_target = round(vat_sum * factor, 2)
        elif base_target is not None and vat_target is None:
            vat_target = round(total_target - base_target, 2)
        elif vat_target is not None and base_target is None:
            base_target = round(total_target - vat_target, 2)
    if base_target is None or vat_target is None:
        return lines

    base_factor = base_target / base_sum if base_sum > 0 else 1
    vat_factor = vat_target / vat_sum if vat_sum > 0 else 1
    adjusted: List[Dict[str, Any]] = []
    for line in lines:
        base = _round_amount((line.get("base") or 0) * base_factor)
        vat = _round_amount((line.get("vat_amount") or 0) * vat_factor)
        rate_calc = round((vat / base) * 100, 2) if base > 0 else None
        rate_final = _group_standard_vat_rate(rate_calc)
        adjusted.append(
            {
                "base": base,
                "vat_amount": vat,
                "rate": rate_final,
                "total": _round_amount(base + vat),
            }
        )

    # Fix rounding deltas on last line.
    base_delta = _round_amount(base_target - sum(line["base"] or 0 for line in adjusted))
    vat_delta = _round_amount(vat_target - sum(line["vat_amount"] or 0 for line in adjusted))
    if adjusted:
        last = adjusted[-1]
        last_base = _round_amount((last.get("base") or 0) + base_delta)
        last_vat = _round_amount((last.get("vat_amount") or 0) + vat_delta)
        last_rate = round((last_vat / last_base) * 100, 2) if last_base > 0 else None
        adjusted[-1] = {
            "base": last_base,
            "vat_amount": last_vat,
            "rate": _group_standard_vat_rate(last_rate),
            "total": _round_amount(last_base + last_vat),
        }
    return adjusted


def normalize_and_validate_amounts(extracted: Dict[str, Any]) -> Dict[str, Any]:
    result = dict(extracted or {})
    totals_payload = result.get("totals") if isinstance(result.get("totals"), dict) else {}
    incoming_source = result.get("amount_source") or "llm"

    base_amount = _normalize_amount(
        _pick_first_non_empty(
            result.get("base_amount"),
            totals_payload.get("base"),
            totals_payload.get("base_amount"),
        )
    )
    vat_amount = _normalize_amount(
        _pick_first_non_empty(
            result.get("vat_amount"),
            totals_payload.get("vat"),
            totals_payload.get("vat_amount"),
        )
    )
    total_amount = _normalize_amount(
        _pick_first_non_empty(
            result.get("total_amount"),
            totals_payload.get("total"),
            totals_payload.get("total_amount"),
        )
    )
    vat_rate = _normalize_rate(result.get("vat_rate"))
    withholding_amount = _normalize_amount(result.get("withholding_amount")) or 0.0

    raw_breakdown = result.get("vat_breakdown") or []
    if isinstance(raw_breakdown, str):
        try:
            raw_breakdown = json.loads(raw_breakdown)
        except json.JSONDecodeError:
            raw_breakdown = []
    if isinstance(raw_breakdown, dict):
        raw_breakdown = [raw_breakdown]

    normalized_breakdown: List[Dict[str, Any]] = []
    breakdown_warning = False
    for line in raw_breakdown:
        if not isinstance(line, dict):
            continue
        base_line = _normalize_amount(
            _pick_first_non_empty(line.get("base"), line.get("base_amount"))
        )
        vat_line = _normalize_amount(
            _pick_first_non_empty(line.get("vat_amount"), line.get("vat"))
        )
        if base_line is None or vat_line is None:
            continue
        rate_calc = round((vat_line / base_line) * 100, 2) if base_line > 0 else None
        rate_final = _group_standard_vat_rate(rate_calc)
        line_total = round(base_line + vat_line, 2)
        normalized_breakdown.append(
            {
                "base": _round_amount(base_line),
                "vat_amount": _round_amount(vat_line),
                "rate": rate_final,
                "total": _round_amount(line_total),
            }
        )

    amount_source = incoming_source or "llm"
    if normalized_breakdown:
        base_sum = round(sum(line["base"] or 0 for line in normalized_breakdown), 2)
        vat_sum = round(sum(line["vat_amount"] or 0 for line in normalized_breakdown), 2)
        total_sum = round(base_sum + vat_sum, 2)
        supported_vat_rates = {0.0, 4.0, 10.0, 21.0}
        breakdown_has_supported_rates = all(
            line.get("rate") in supported_vat_rates for line in normalized_breakdown
        )

        def totals_match(a: Optional[float], b: Optional[float]) -> bool:
            if a is None or b is None:
                return False
            return abs(a - b) <= 0.02

        llm_has_totals = base_amount is not None and vat_amount is not None and total_amount is not None
        llm_consistent = (
            llm_has_totals
            and abs((base_amount + vat_amount - withholding_amount) - total_amount) <= 0.02
        )
        breakdown_consistent = abs((base_sum + vat_sum) - total_sum) <= 0.02

        # If totals come from text/summary, do NOT let breakdown override a lower explicit total.
        if incoming_source in {"regex_tax_summary", "text_total"} and total_amount is not None:
            gross_total = total_amount + withholding_amount
            if total_sum > gross_total + 0.02:
                if len(normalized_breakdown) > 1 and breakdown_consistent:
                    # Multi-line VAT summaries are more reliable than a partial explicit
                    # summary that only captured the first tax band.
                    base_amount = base_sum
                    vat_amount = vat_sum
                    total_amount = round(total_sum - withholding_amount, 2)
                    amount_source = "breakdown"
                else:
                    # Keep explicit total, adjust breakdown to match it (likely OCR noise).
                    breakdown_warning = True
                    normalized_breakdown = _adjust_breakdown_to_targets(
                        normalized_breakdown,
                        base_amount,
                        vat_amount,
                        total_amount + withholding_amount,
                    )
                    base_amount = round(sum(line["base"] or 0 for line in normalized_breakdown), 2)
                    vat_amount = round(sum(line["vat_amount"] or 0 for line in normalized_breakdown), 2)
                    amount_source = incoming_source
            elif llm_consistent:
                amount_source = incoming_source
            else:
                # If a coherent breakdown matches the explicit total, repair the
                # incomplete summary using the breakdown instead of dropping the
                # whole extraction as inconsistent.
                if (
                    breakdown_consistent
                    and abs(total_sum - gross_total) <= 0.02
                    and (base_amount is None or vat_amount is None or not llm_consistent)
                ):
                    base_amount = base_sum
                    vat_amount = vat_sum
                    total_amount = round(total_sum - withholding_amount, 2)
                    amount_source = "breakdown"
                else:
                    # Keep explicit total even if base/vat are incomplete.
                    amount_source = incoming_source
        elif breakdown_consistent:
            llm_matches_breakdown = (
                totals_match(base_amount, base_sum)
                and totals_match(vat_amount, vat_sum)
                and totals_match(total_amount + withholding_amount, total_sum)
            )
            breakdown_looks_partial_or_non_tax = (
                llm_consistent
                and (
                    not breakdown_has_supported_rates
                    or (
                        total_amount is not None
                        and total_sum < (total_amount + withholding_amount) - 0.02
                    )
                )
            )
            # Prefer the line-level breakdown whenever it is coherent and the
            # aggregate totals do not match it. This keeps multi-IVA invoices
            # internally consistent without relying on a single aggregate guess.
            if breakdown_looks_partial_or_non_tax:
                normalized_breakdown = []
                amount_source = amount_source or "llm"
            elif not llm_matches_breakdown:
                base_amount = base_sum
                vat_amount = vat_sum
                total_amount = round(total_sum - withholding_amount, 2)
                amount_source = "breakdown"
            elif llm_consistent:
                amount_source = amount_source or "llm"
            else:
                amount_source = "breakdown"
        else:
            # Keep LLM totals if they exist; otherwise fall back later.
            amount_source = amount_source or "llm"

    analysis_status = result.get("analysis_status") or "ok"
    mismatch = (
        base_amount is not None
        and vat_amount is not None
        and total_amount is not None
        and abs((base_amount + vat_amount - withholding_amount) - total_amount) > 0.02
    )
    if total_amount is None or mismatch or base_amount is None or vat_amount is None:
        analysis_status = "partial"
        if mismatch and not breakdown_warning:
            base_amount = None
            vat_amount = None
            total_amount = None
            vat_rate = None
            normalized_breakdown = []
        elif (base_amount is None or vat_amount is None) and not breakdown_warning:
            base_amount = None
            vat_amount = None
            vat_rate = None
            normalized_breakdown = []
        if total_amount is None:
            amount_source = "fallback"

    if normalized_breakdown:
        if len(normalized_breakdown) == 1:
            vat_rate = normalized_breakdown[0].get("rate")
        else:
            vat_rate = None

    result.update(
        {
            "base_amount": _round_amount(base_amount),
            "vat_amount": _round_amount(vat_amount),
            "total_amount": _round_amount(total_amount),
            "vat_rate": vat_rate,
            "vat_breakdown": normalized_breakdown,
            "withholding_amount": _round_amount(withholding_amount),
            "analysis_status": analysis_status,
            "amount_source": amount_source,
            "breakdown_warning": breakdown_warning,
        }
    )
    return result


def _recover_from_tax_summary_if_needed(
    analysis_status: str,
    base_amount: Optional[float],
    vat_amount: Optional[float],
    total_amount: Optional[float],
    vat_rate: Optional[float],
    vat_breakdown: List[Dict[str, Any]],
    amount_source: Optional[str],
    tax_summary: Optional[Dict[str, Any]],
) -> Dict[str, Any]:
    if not isinstance(tax_summary, dict) or not tax_summary.get("found"):
        return {
            "analysis_status": analysis_status,
            "base_amount": base_amount,
            "vat_amount": vat_amount,
            "total_amount": total_amount,
            "vat_rate": vat_rate,
            "vat_breakdown": vat_breakdown,
            "amount_source": amount_source,
            "breakdown_warning": False,
        }

    if (
        base_amount is not None
        and vat_amount is not None
        and total_amount is not None
        and abs((base_amount + vat_amount) - total_amount) <= 0.02
    ):
        return {
            "analysis_status": analysis_status,
            "base_amount": base_amount,
            "vat_amount": vat_amount,
            "total_amount": total_amount,
            "vat_rate": vat_rate,
            "vat_breakdown": vat_breakdown,
            "amount_source": amount_source,
            "breakdown_warning": False,
        }

    rescued = normalize_and_validate_amounts(
        {
            "analysis_status": "ok",
            "base_amount": tax_summary.get("base_amount"),
            "vat_amount": tax_summary.get("vat_amount"),
            "total_amount": tax_summary.get("total_amount"),
            "vat_rate": tax_summary.get("vat_rate"),
            "vat_breakdown": tax_summary.get("breakdown") or vat_breakdown,
            "amount_source": "regex_tax_summary",
        }
    )
    if (
        rescued.get("base_amount") is not None
        and rescued.get("vat_amount") is not None
        and rescued.get("total_amount") is not None
    ):
        return rescued
    return {
        "analysis_status": analysis_status,
        "base_amount": base_amount,
        "vat_amount": vat_amount,
        "total_amount": total_amount,
        "vat_rate": vat_rate,
        "vat_breakdown": vat_breakdown,
        "amount_source": amount_source,
        "breakdown_warning": False,
    }


def _force_tax_summary_result_if_available(
    analysis_status: str,
    base_amount: Optional[float],
    vat_amount: Optional[float],
    total_amount: Optional[float],
    vat_rate: Optional[float],
    vat_breakdown: List[Dict[str, Any]],
    amount_source: Optional[str],
    breakdown_warning: bool,
    tax_summary: Optional[Dict[str, Any]],
    withholding_amount: Optional[float] = None,
) -> Dict[str, Any]:
    if not isinstance(tax_summary, dict) or not tax_summary.get("found"):
        return {
            "analysis_status": analysis_status,
            "base_amount": base_amount,
            "vat_amount": vat_amount,
            "total_amount": total_amount,
            "vat_rate": vat_rate,
            "vat_breakdown": vat_breakdown,
            "amount_source": amount_source,
            "breakdown_warning": breakdown_warning,
        }

    summary_breakdown = tax_summary.get("breakdown") or []
    summary_base = _normalize_amount(tax_summary.get("base_amount"))
    summary_vat = _normalize_amount(tax_summary.get("vat_amount"))
    summary_total = _normalize_amount(tax_summary.get("total_amount"))
    summary_rate = _normalize_rate(tax_summary.get("vat_rate"))

    if summary_breakdown and (
        summary_base is None or summary_vat is None or summary_total is None
    ):
        summary_values = _summarize_vat_breakdown(summary_breakdown)
        if summary_values:
            summary_base, summary_vat, summary_total = summary_values

    summary_is_complete = (
        summary_base is not None
        and summary_vat is not None
        and summary_total is not None
        and abs((summary_base + summary_vat) - summary_total) <= 0.02
    )
    current_is_complete = (
        base_amount is not None
        and vat_amount is not None
        and total_amount is not None
        and abs((base_amount + vat_amount - max(float(withholding_amount or 0), 0)) - total_amount)
        <= 0.02
    )

    should_force_summary = summary_is_complete and (
        not current_is_complete
        or amount_source == "fallback"
        or base_amount is None
        or vat_amount is None
        or total_amount is None
    )

    if not should_force_summary:
        return {
            "analysis_status": analysis_status,
            "base_amount": base_amount,
            "vat_amount": vat_amount,
            "total_amount": total_amount,
            "vat_rate": vat_rate,
            "vat_breakdown": vat_breakdown,
            "amount_source": amount_source,
            "breakdown_warning": breakdown_warning,
        }

    normalized_breakdown = summary_breakdown if summary_breakdown else vat_breakdown
    final_rate = summary_rate if len(normalized_breakdown) <= 1 else None
    return {
        "analysis_status": "ok",
        "base_amount": summary_base,
        "vat_amount": summary_vat,
        "total_amount": summary_total,
        "vat_rate": final_rate,
        "vat_breakdown": normalized_breakdown,
        "amount_source": "regex_tax_summary",
        "breakdown_warning": False,
    }


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


def _is_low_quality_ocr(text: str, min_chars: int = 200) -> bool:
    if not text:
        return True
    stripped = text.strip()
    if not stripped:
        return True
    total = len(stripped)
    alnum = sum(1 for char in stripped if char.isalnum())
    letters = sum(1 for char in stripped if char.isalpha())
    if alnum < min_chars:
        return True
    if alnum == 0:
        return True
    if letters / alnum < 0.3:
        return True
    garbage = sum(
        1
        for char in stripped
        if not (char.isalnum() or char.isspace() or char in ".,:-/%()")
    )
    if garbage / max(total, 1) > 0.3:
        return True
    tokens = re.findall(r"[A-Za-zÀ-ÿ0-9]{2,}", stripped)
    if len(set(tokens)) < 10:
        return True
    return False


def _has_amount_hints(text: str) -> bool:
    if not text:
        return False
    lowered = text.lower()
    keywords = ["total", "base", "imponible", "iva", "vat", "subtotal"]
    amount_pattern = r"\d{1,3}(?:[\.\s]\d{3})*(?:[,\.\u00b7]\d{2})"
    percent_pattern = r"\d{1,2}\s?%"
    if any(keyword in lowered for keyword in keywords):
        for line in text.splitlines():
            line_lower = line.lower()
            if any(keyword in line_lower for keyword in keywords):
                if re.search(amount_pattern, line) or re.search(percent_pattern, line):
                    return True
    if re.search(amount_pattern, text) and re.search(percent_pattern, text):
        return True
    return False


def _extract_pdf_text(file_path: str) -> str:
    if fitz is None:
        logger.warning("PyMuPDF no disponible. Texto PDF no extraido.")
        return ""
    with fitz.open(file_path) as doc:
        parts = []
        for page in doc:
            parts.append(page.get_text("text"))
        return "\n".join(parts).strip()


def _extract_pdf_text_from_bytes(data: bytes) -> str:
    if fitz is None:
        logger.warning("PyMuPDF no disponible. Texto PDF no extraido.")
        return ""
    with fitz.open(stream=data, filetype="pdf") as doc:
        parts = []
        for page in doc:
            parts.append(page.get_text("text"))
        return "\n".join(parts).strip()


def _compact_native_pdf_page_text(text: str) -> str:
    """Preserve native PDF reading order while removing layout-only whitespace."""
    compact_lines: List[str] = []
    previous_line = None
    for raw_line in _normalize_ocr_amount_text(text or "").splitlines():
        line = re.sub(r"[ \t]+", " ", raw_line).strip()
        if not line or line == previous_line:
            continue
        compact_lines.append(line)
        previous_line = line
    return "\n".join(compact_lines)


def _fast_text_layout_token_from_word(word: Any, page_number: int) -> Optional[Dict[str, Any]]:
    """Normalize PyMuPDF ``words`` output without retaining it beyond analysis."""
    try:
        if isinstance(word, dict):
            text = str(word.get("text") or "").strip()
            x0, y0, x1, y1 = (
                float(word["x0"]),
                float(word["y0"]),
                float(word["x1"]),
                float(word["y1"]),
            )
            block = int(word.get("block", 0))
            line = int(word.get("line", 0))
            word_index = int(word.get("word", 0))
        else:
            x0, y0, x1, y1, text, block, line, word_index = word[:8]
            text = str(text or "").strip()
            x0, y0, x1, y1 = float(x0), float(y0), float(x1), float(y1)
            block, line, word_index = int(block), int(line), int(word_index)
    except (IndexError, KeyError, TypeError, ValueError):
        return None
    if not text or x1 <= x0 or y1 <= y0:
        return None
    return {
        "page": page_number,
        "text": text,
        "x0": x0,
        "y0": y0,
        "x1": x1,
        "y1": y1,
        "block": block,
        "line": line,
        "word": word_index,
    }


def _fast_text_layout_median(values: List[float]) -> float:
    if not values:
        return 1.0
    ordered = sorted(values)
    middle = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[middle]
    return (ordered[middle - 1] + ordered[middle]) / 2


def _fast_text_layout_token_height(token: Dict[str, Any]) -> float:
    return max(float(token["y1"]) - float(token["y0"]), 0.1)


def _fast_text_layout_vertical_overlap(first: Dict[str, Any], second: Dict[str, Any]) -> float:
    return max(0.0, min(first["y1"], second["y1"]) - max(first["y0"], second["y0"]))


def _fast_text_layout_row_accepts_token(row: Dict[str, Any], token: Dict[str, Any]) -> bool:
    """Use source line metadata first, then local token-height geometry."""
    if (token["block"], token["line"]) in row["source_lines"]:
        return True
    row_height = max(float(row["y1"]) - float(row["y0"]), 0.1)
    token_height = _fast_text_layout_token_height(token)
    overlap = _fast_text_layout_vertical_overlap(row, token)
    if overlap / min(row_height, token_height) >= 0.6:
        return True
    row_center = (float(row["y0"]) + float(row["y1"])) / 2
    token_center = (float(token["y0"]) + float(token["y1"])) / 2
    local_height = max(row["median_token_height"], token_height)
    return abs(row_center - token_center) <= local_height * 0.45


def _finalize_fast_text_layout_row(row: Dict[str, Any], index: int) -> Dict[str, Any]:
    tokens = sorted(row["tokens"], key=lambda token: (token["x0"], token["word"]))
    row["tokens"] = tokens
    row["index"] = index
    row["x0"] = min(token["x0"] for token in tokens)
    row["y0"] = min(token["y0"] for token in tokens)
    row["x1"] = max(token["x1"] for token in tokens)
    row["y1"] = max(token["y1"] for token in tokens)
    row["median_token_height"] = _fast_text_layout_median(
        [_fast_text_layout_token_height(token) for token in tokens]
    )
    row["text"] = " ".join(token["text"] for token in tokens)
    row["normalized_text"] = _normalize_fast_text_parser_line(row["text"])
    return row


def _build_fast_text_document_layout(raw_pages_words: List[List[Any]]) -> Dict[str, Any]:
    """Build minimal visual rows from PyMuPDF words for the V14 verifier.

    The representation is deliberately ephemeral. It preserves only the local
    geometry required to prove a model-proposed value, not a document model or
    reconstructed table suitable for persistence.
    """
    pages: List[Dict[str, Any]] = []
    for page_number, raw_words in enumerate(raw_pages_words or [], start=1):
        tokens = [
            token
            for word in raw_words or []
            if (token := _fast_text_layout_token_from_word(word, page_number)) is not None
        ]
        tokens.sort(key=lambda token: (token["y0"], token["x0"], token["block"], token["line"], token["word"]))
        rows: List[Dict[str, Any]] = []
        for token in tokens:
            compatible = [row for row in rows if _fast_text_layout_row_accepts_token(row, token)]
            if compatible:
                row = min(
                    compatible,
                    key=lambda candidate: abs(
                        ((candidate["y0"] + candidate["y1"]) / 2)
                        - ((token["y0"] + token["y1"]) / 2)
                    ),
                )
                row["tokens"].append(token)
                row["source_lines"].add((token["block"], token["line"]))
                row["median_token_height"] = _fast_text_layout_median(
                    [_fast_text_layout_token_height(item) for item in row["tokens"]]
                )
                row["x0"] = min(row["x0"], token["x0"])
                row["y0"] = min(row["y0"], token["y0"])
                row["x1"] = max(row["x1"], token["x1"])
                row["y1"] = max(row["y1"], token["y1"])
                continue
            rows.append(
                {
                    "page": page_number,
                    "tokens": [token],
                    "source_lines": {(token["block"], token["line"])},
                    "x0": token["x0"],
                    "y0": token["y0"],
                    "x1": token["x1"],
                    "y1": token["y1"],
                    "median_token_height": _fast_text_layout_token_height(token),
                }
            )
        rows.sort(key=lambda row: (row["y0"], row["x0"]))
        pages.append(
            {
                "page": page_number,
                "rows": [
                    _finalize_fast_text_layout_row(row, index)
                    for index, row in enumerate(rows)
                ],
            }
        )
    return {"pages": pages}


def prepare_invoice_v2_fast_text(
    file_bytes: bytes,
    *,
    filename: Optional[str] = None,
    mime_type: Optional[str] = None,
) -> Dict[str, Any]:
    """Build an in-memory, page-labelled representation for the V2 shadow path.

    This function deliberately does not extract accounting fields. Its output is
    only a faithful compact rendering of native PDF text and operational size
    metadata. Callers must not persist the returned ``text`` value.
    """
    filename = filename or "documento"
    is_pdf = (mime_type or "").lower() == "application/pdf" or filename.lower().endswith(
        ".pdf"
    )
    base_result: Dict[str, Any] = {
        "eligible": False,
        "reason": "not_pdf",
        "text_representation_version": _FAST_TEXT_REPRESENTATION_VERSION,
        "page_count": 0,
        "native_text_chars": 0,
        "sent_text_chars": 0,
        "document_text_complete": False,
        "size_reduction_ratio": None,
        "text": "",
        # V14 keeps its layout private to this in-memory object. Callers only
        # persist the explicit operational fields above.
        "_document_layout": None,
        "layout_build_ms": None,
    }
    if not is_pdf:
        return base_result
    if fitz is None:
        base_result["reason"] = "pymupdf_unavailable"
        return base_result
    try:
        with fitz.open(stream=file_bytes, filetype="pdf") as document:
            page_count = len(document)
            raw_pages = []
            raw_pages_words = []
            layout_started = time.monotonic()
            for page in document:
                raw_pages.append(page.get_text("text", sort=True) or "")
                raw_pages_words.append(page.get_text("words", sort=True) or [])
            document_layout = _build_fast_text_document_layout(raw_pages_words)
            layout_build_ms = round((time.monotonic() - layout_started) * 1000)
    except Exception:
        logger.info("Fast path V2 no elegible (%s): native_text_unavailable", filename)
        base_result["reason"] = "native_text_unavailable"
        return base_result

    native_text = "\n".join(raw_pages)
    native_text_chars = len(native_text)
    base_result["page_count"] = page_count
    base_result["native_text_chars"] = native_text_chars
    if not _is_text_significant(native_text, PDF_TEXT_THRESHOLD):
        base_result["reason"] = "scanned_pdf"
        return base_result

    page_blocks = []
    for page_number, raw_page in enumerate(raw_pages, start=1):
        compact_page = _compact_native_pdf_page_text(raw_page)
        if compact_page:
            page_blocks.append(f"[PÁGINA {page_number}]\n{compact_page}")
    compact_text = "\n\n".join(page_blocks).strip()
    if not _is_text_significant(compact_text, PDF_TEXT_THRESHOLD):
        base_result["reason"] = "native_text_insufficient"
        return base_result

    document_text_complete = True
    max_chars = _get_invoice_v2_fast_text_max_chars()
    if len(compact_text) > max_chars:
        # Keep document headers and final tax/total blocks instead of silently
        # dropping the end of a long digital invoice.
        marker = "\n\n[CONTENIDO INTERMEDIO OMITIDO POR LÍMITE DE TAMAÑO]\n\n"
        head_size = max(1, int((max_chars - len(marker)) * 0.6))
        tail_size = max(1, max_chars - len(marker) - head_size)
        compact_text = compact_text[:head_size] + marker + compact_text[-tail_size:]
        document_text_complete = False

    sent_text_chars = len(compact_text)
    base_result.update(
        {
            "eligible": True,
            "reason": "native_text_sufficient",
            "sent_text_chars": sent_text_chars,
            "document_text_complete": document_text_complete,
            "size_reduction_ratio": round(
                max(0.0, 1 - (sent_text_chars / max(native_text_chars, 1))), 4
            ),
            "text": compact_text,
            "_document_layout": document_layout,
            "layout_build_ms": layout_build_ms,
        }
    )
    return base_result


def _render_document_pages_for_vision(
    data: bytes,
    *,
    is_pdf: bool,
    mime_type: Optional[str] = None,
) -> List[str]:
    """Return compact document pages for the invoice model to inspect visually."""
    if not data:
        return []
    if not is_pdf:
        image_mime = mime_type if mime_type in {"image/jpeg", "image/png"} else "image/png"
        return [
            f"data:{image_mime};base64,"
            + base64.b64encode(data).decode("ascii")
        ]
    if fitz is None:
        return []
    try:
        images = []
        with fitz.open(stream=data, filetype="pdf") as doc:
            for page_number in range(min(len(doc), INVOICE_VISION_MAX_PAGES)):
                page = doc[page_number]
                max_dimension = max(page.rect.width, page.rect.height, 1)
                scale = min(2.0, INVOICE_VISION_MAX_DIMENSION / max_dimension)
                pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                image_data = pixmap.tobytes("png")
                images.append(
                    "data:image/png;base64," + base64.b64encode(image_data).decode("ascii")
                )
        return images
    except Exception as exc:
        logger.warning("No se pudo preparar el documento para vision: %s", exc)
        return []


def _render_first_pdf_page_for_vision(data: bytes) -> Optional[str]:
    """Backward-compatible helper for callers that only need page one."""
    pages = _render_document_pages_for_vision(data, is_pdf=True)
    return pages[0] if pages else None


def _get_ocr_reader():
    global _ocr_reader
    if not _ocr_is_enabled():
        logger.info("OCR desactivado por configuracion de memoria.")
        return None
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
    if runtime_env == "production":
        download_enabled = False
    model_contents = []
    if os.path.isdir(model_dir):
        try:
            model_contents = os.listdir(model_dir)
        except OSError:
            model_contents = []
    if not os.path.isdir(model_dir) or not model_contents:
        if runtime_env == "production":
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
    if fitz is None:
        logger.warning("PyMuPDF no disponible. OCR PDF omitido.")
        return ""
    reader = _get_ocr_reader()
    if reader is None:
        return ""
    try:
        import numpy as np
    except ImportError as exc:
        logger.warning("NumPy no disponible para OCR: %s", exc)
        return ""

    parts = []
    max_pages, max_zoom, max_dim_limit = _production_ocr_limits()
    with fitz.open(file_path) as doc:
        for idx, page in enumerate(doc):
            if idx >= max_pages:
                break
            base_scale = max_zoom
            max_dim = max(page.rect.width, page.rect.height, 1)
            if max_dim * base_scale > max_dim_limit:
                base_scale = max_dim_limit / max_dim
            matrix = fitz.Matrix(base_scale, base_scale)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.height, pix.width, pix.n
            )
            if pix.n == 4:
                image = image[:, :, :3]
            lines = reader.readtext(image, detail=0)
            if lines:
                parts.append("\n".join(lines))
            del image, pix
            gc.collect()
    return "\n".join(parts).strip()


def _extract_pdf_text_ocr_from_bytes(data: bytes) -> str:
    if fitz is None:
        logger.warning("PyMuPDF no disponible. OCR PDF omitido.")
        return ""
    reader = _get_ocr_reader()
    if reader is None:
        return ""
    try:
        import numpy as np
    except ImportError as exc:
        logger.warning("NumPy no disponible para OCR: %s", exc)
        return ""

    parts = []
    start_time = time.time()
    max_pages, max_zoom, max_dim_limit = _production_ocr_limits()
    with fitz.open(stream=data, filetype="pdf") as doc:
        for idx, page in enumerate(doc):
            if idx >= max_pages:
                break
            if time.time() - start_time > OCR_MAX_SECONDS:
                break
            base_scale = max_zoom
            max_dim = max(page.rect.width, page.rect.height, 1)
            if max_dim * base_scale > max_dim_limit:
                base_scale = max_dim_limit / max_dim
            matrix = fitz.Matrix(base_scale, base_scale)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.height, pix.width, pix.n
            )
            if pix.n == 4:
                image = image[:, :, :3]
            lines = reader.readtext(image, detail=0)
            if lines:
                parts.append("\n".join(lines))
            if idx == 0:
                preview_text = "\n".join(parts).strip()
                if _is_low_quality_ocr(preview_text) and not _has_amount_hints(preview_text):
                    del image, pix
                    gc.collect()
                    return preview_text
            del image, pix
            gc.collect()
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
    _, _, max_dim_limit = _production_ocr_limits()
    height, width = image.shape[:2]
    max_dim = max(width, height, 1)
    if max_dim > max_dim_limit:
        scale = max_dim_limit / max_dim
        image = cv2.resize(
            image, (int(width * scale), int(height * scale)), interpolation=cv2.INTER_AREA
        )
    lines = reader.readtext(image, detail=0)
    if not lines:
        return ""
    return "\n".join(lines).strip()


def _run_with_timeout(func, timeout: int, *args, **kwargs):
    with ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(func, *args, **kwargs)
        try:
            return future.result(timeout=timeout), False
        except FuturesTimeoutError:
            return None, True


def _nullable(kind: str) -> Dict[str, Any]:
    return {"type": [kind, "null"]}


def _strict_object(properties: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "type": "object",
        "properties": properties,
        "required": list(properties),
        "additionalProperties": False,
    }


_EVIDENCE_SCHEMA = _strict_object(
    {"page": _nullable("integer"), "evidence": _nullable("string"), "confidence": _nullable("number")}
)
_PARTY_SCHEMA = _strict_object(
    {
        "legal_name": _nullable("string"), "commercial_name": _nullable("string"),
        "tax_id": _nullable("string"), "address": _nullable("string"),
        "country": _nullable("string"), "billing_address": _nullable("string"),
        "shipping_address": _nullable("string"),
    }
)
_INVOICE_SCHEMA = _strict_object(
    {
        "invoice_number": _nullable("string"), "issue_date": _nullable("string"),
        "currency": _nullable("string"), "order_reference": _nullable("string"),
        "delivery_note": _nullable("string"),
    }
)
_ITEM_SCHEMA = _strict_object(
    {
        "reference": _nullable("string"), "description": _nullable("string"), "quantity": _nullable("number"),
        "unit": _nullable("string"), "gross_unit_price": _nullable("number"),
        "discount_percentage": _nullable("number"), "discount_amount": _nullable("number"),
        "net_unit_price": _nullable("number"), "taxable_amount": _nullable("number"),
        "vat_rate": _nullable("number"), "vat_amount": _nullable("number"), "line_total": _nullable("number"),
        "is_free": {"type": "boolean"}, "is_promotional": {"type": "boolean"},
    }
)
_TAX_SCHEMA = _strict_object(
    {"taxable_base": _nullable("number"), "vat_rate": _nullable("number"), "vat_amount": _nullable("number")}
)
_INSTALLMENT_SCHEMA = _strict_object(
    {
        "due_date": _nullable("string"), "amount": _nullable("number"),
        "actual_payment_date": _nullable("string"), "payment_evidence": _nullable("string"),
    }
)
_TOTALS_SCHEMA = _strict_object(
    {
        "subtotal": _nullable("number"), "global_discount": _nullable("number"), "shipping": _nullable("number"),
        "taxable_base": _nullable("number"), "vat_amount": _nullable("number"), "withholding": _nullable("number"),
        "other_taxes": _nullable("number"), "total": _nullable("number"),
    }
)
INVOICE_EXTRACTION_SCHEMA = _strict_object(
    {
        "document_type": _nullable("string"), "supplier": _PARTY_SCHEMA, "customer": _PARTY_SCHEMA,
        "invoice": _INVOICE_SCHEMA, "items": {"type": "array", "items": _ITEM_SCHEMA},
        "taxes": {"type": "array", "items": _TAX_SCHEMA}, "totals": _TOTALS_SCHEMA,
        "payment": _strict_object({"method": _nullable("string"), "status": _nullable("string"), "iban": _nullable("string"), "bic": _nullable("string")}),
        "installments": {"type": "array", "items": _INSTALLMENT_SCHEMA},
        "observations": {"type": "array", "items": {"type": "string"}},
        "field_evidence": _strict_object({
            "supplier": _EVIDENCE_SCHEMA, "customer": _EVIDENCE_SCHEMA, "invoice_number": _EVIDENCE_SCHEMA,
            "issue_date": _EVIDENCE_SCHEMA, "totals": _EVIDENCE_SCHEMA, "taxes": _EVIDENCE_SCHEMA,
            "installments": _EVIDENCE_SCHEMA,
        }),
        "validation": _strict_object({"is_consistent": _nullable("boolean"), "issues": {"type": "array", "items": {"type": "string"}}}),
    }
)

# V2 shadow intentionally asks for only the accounting contract needed to
# compare a digital invoice with V1. Evidence is used for validation in-memory
# and is never persisted with a shadow run.
_FAST_TEXT_PARTY_SCHEMA = _strict_object(
    {"legal_name": _nullable("string"), "tax_id": _nullable("string")}
)
_FAST_TEXT_INVOICE_SCHEMA = _strict_object(
    {
        "invoice_number": _nullable("string"),
        "issue_date": _nullable("string"),
        "currency": _nullable("string"),
    }
)
_FAST_TEXT_TAX_SCHEMA = _strict_object(
    {
        "taxable_base": _nullable("number"),
        "vat_rate": _nullable("number"),
        "vat_amount": _nullable("number"),
    }
)
_FAST_TEXT_TOTALS_SCHEMA = _strict_object(
    {
        "taxable_base": _nullable("number"),
        "vat_amount": _nullable("number"),
        "withholding": _nullable("number"),
        "other_taxes": _nullable("number"),
        "total": _nullable("number"),
    }
)
INVOICE_FAST_TEXT_SCHEMA = _strict_object(
    {
        "supplier": _FAST_TEXT_PARTY_SCHEMA,
        "customer": _FAST_TEXT_PARTY_SCHEMA,
        "invoice": _FAST_TEXT_INVOICE_SCHEMA,
        "due_dates": {"type": "array", "items": {"type": "string"}},
        "taxes": {"type": "array", "items": _FAST_TEXT_TAX_SCHEMA},
        "totals": _FAST_TEXT_TOTALS_SCHEMA,
        "field_evidence": _strict_object(
            {
                "supplier": _EVIDENCE_SCHEMA,
                "invoice_number": _EVIDENCE_SCHEMA,
                "issue_date": _EVIDENCE_SCHEMA,
                "taxes": _EVIDENCE_SCHEMA,
                "totals": _EVIDENCE_SCHEMA,
                "withholding": _EVIDENCE_SCHEMA,
                "due_dates": _EVIDENCE_SCHEMA,
            }
        ),
    }
)


def _money_decimal(value: Any) -> Optional[Decimal]:
    amount = _normalize_amount(value)
    if amount is None:
        return None
    try:
        return Decimal(str(amount)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return None


def _validate_structured_invoice(data: Dict[str, Any], document_type: str) -> List[str]:
    """Validate model output without silently changing any reported value."""
    issues: List[str] = []
    supplier = data.get("supplier") or {}
    invoice = data.get("invoice") or {}
    totals = data.get("totals") or {}
    taxes = data.get("taxes") or []
    installments = data.get("installments") or []
    items = data.get("items") or []
    if document_type != "expense_payroll" and not supplier.get("legal_name"):
        issues.append("Falta el proveedor/emisor.")
    if document_type != "expense_payroll" and not invoice.get("invoice_number"):
        issues.append("Falta el número de factura.")
    base, vat, withholding, other, total = (
        _money_decimal(totals.get("taxable_base")), _money_decimal(totals.get("vat_amount")),
        _money_decimal(totals.get("withholding")) or Decimal("0.00"),
        _money_decimal(totals.get("other_taxes")) or Decimal("0.00"), _money_decimal(totals.get("total")),
    )
    if total is None:
        issues.append("Falta el total de la factura.")
    elif base is not None and vat is not None:
        expected = (base + vat + other - abs(withholding)).quantize(Decimal("0.01"))
        if abs(expected - total) > Decimal("0.01"):
            issues.append("La ecuación base + IVA + otros impuestos - retención no cuadra con el total.")
    if taxes and base is not None and vat is not None:
        tax_base = sum((_money_decimal(line.get("taxable_base")) or Decimal("0.00") for line in taxes), Decimal("0.00"))
        tax_vat = sum((_money_decimal(line.get("vat_amount")) or Decimal("0.00") for line in taxes), Decimal("0.00"))
        if abs(tax_base - base) > Decimal("0.01") or abs(tax_vat - vat) > Decimal("0.01"):
            issues.append("El desglose de impuestos no coincide con los totales.")
    taxable_lines = []
    for item in items:
        quantity = _money_decimal(item.get("quantity"))
        gross_price = _money_decimal(item.get("gross_unit_price"))
        net_price = _money_decimal(item.get("net_unit_price"))
        taxable = _money_decimal(item.get("taxable_amount"))
        discount_percentage = _money_decimal(item.get("discount_percentage"))
        if quantity is not None and gross_price is not None and net_price is not None and discount_percentage is not None:
            expected_net = (gross_price * (Decimal("1") - discount_percentage / Decimal("100"))).quantize(Decimal("0.01"))
            if abs(expected_net - net_price) > Decimal("0.01"):
                issues.append("El descuento de una línea no cuadra con su precio neto.")
                break
        if taxable is not None:
            taxable_lines.append(taxable)
    if taxable_lines and base is not None and abs(sum(taxable_lines, Decimal("0.00")) - base) > Decimal("0.01"):
        issues.append("La suma de líneas no coincide con la base imponible.")
    amounts = [_money_decimal(item.get("amount")) for item in installments]
    if amounts and all(amount is not None for amount in amounts) and total is not None:
        if abs(sum(amounts, Decimal("0.00")) - total) > Decimal("0.01"):
            issues.append("La suma de vencimientos no coincide con el total.")
    for evidence in (data.get("field_evidence") or {}).values():
        if isinstance(evidence, dict) and evidence.get("confidence") is not None and evidence["confidence"] < 0.75:
            issues.append("Existe un campo importante con baja confianza.")
            break
    return issues


def _requires_invoice_audit(validation_issues: List[str]) -> bool:
    """Reserve the expensive second pass for missing or inconsistent accounting totals."""
    critical_issues = {
        "Falta el total de la factura.",
        "La ecuación base + IVA + otros impuestos - retención no cuadra con el total.",
        "El desglose de impuestos no coincide con los totales.",
    }
    return any(issue in critical_issues for issue in validation_issues)


def _response_input_for_invoice(file_bytes: bytes, filename: str, mime_type: str, extracted_text: str, prompt: str) -> List[Dict[str, Any]]:
    content: List[Dict[str, Any]] = [{"type": "input_text", "text": prompt}]
    data_url = f"data:{mime_type or 'application/octet-stream'};base64," + base64.b64encode(file_bytes).decode("ascii")
    if mime_type == "application/pdf" or filename.lower().endswith(".pdf"):
        content.append({"type": "input_file", "filename": filename, "file_data": data_url, "detail": "high"})
    elif mime_type in {"image/jpeg", "image/png"}:
        content.append({"type": "input_image", "image_url": data_url, "detail": "high"})
    if extracted_text:
        content.append({"type": "input_text", "text": "Texto nativo/apoyo (puede estar desordenado):\n" + extracted_text[:30000]})
    return [{"role": "user", "content": content}]


def _response_value(value: Any, key: str) -> Any:
    return value.get(key) if isinstance(value, dict) else getattr(value, key, None)


def _response_has_refusal(response: Any) -> bool:
    for item in _response_value(response, "output") or []:
        for content in _response_value(item, "content") or []:
            if _response_value(content, "type") == "refusal":
                return True
    return False


def _response_usage_values(response: Any) -> Dict[str, Optional[int]]:
    usage = _response_value(response, "usage") or {}
    output_details = _response_value(usage, "output_tokens_details") or {}
    return {
        "input_tokens": _response_value(usage, "input_tokens"),
        "output_tokens": _response_value(usage, "output_tokens"),
        "reasoning_tokens": _response_value(output_details, "reasoning_tokens"),
        "total_tokens": _response_value(usage, "total_tokens"),
    }


def _log_invoice_response_usage(
    response: Any,
    *,
    status: Any,
    max_output_tokens: int,
    reasoning_effort: str,
    audit: bool,
) -> None:
    usage = _response_usage_values(response)
    logger.info(
        "OpenAI Responses invoice usage: response_status=%s input_tokens=%s output_tokens=%s reasoning_tokens=%s total_tokens=%s max_output_tokens=%s reasoning_effort=%s audit=%s",
        status,
        usage["input_tokens"],
        usage["output_tokens"],
        usage["reasoning_tokens"],
        usage["total_tokens"],
        max_output_tokens,
        reasoning_effort,
        audit,
    )


def _record_invoice_response_telemetry(
    telemetry: Optional[Dict[str, Any]],
    *,
    response: Any = None,
    model: Optional[str] = None,
    elapsed_ms: int,
    audit: bool,
) -> None:
    """Accumulate only operational metadata from an existing Responses call."""
    if telemetry is None:
        return

    telemetry["openai_ms"] = int(telemetry.get("openai_ms") or 0) + max(int(elapsed_ms), 0)
    if model:
        telemetry["openai_model"] = str(
            _response_value(response, "model") or model
        )[:255]
    if audit:
        # The current pipeline calls this limited audit as its second review.
        telemetry["audit_used"] = True
        telemetry["second_review_used"] = True
    if response is None:
        return

    for field, value in _response_usage_values(response).items():
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            telemetry[field] = int(telemetry.get(field) or 0) + max(int(value), 0)


def _invoice_analysis_telemetry_result(
    result: Dict[str, Any], telemetry: Dict[str, Any], return_telemetry: bool
):
    if return_telemetry:
        return result, telemetry
    return result


def _call_invoice_responses(
    client,
    *,
    file_bytes: bytes,
    filename: str,
    mime_type: str,
    extracted_text: str,
    prompt: str,
    audit_issues: Optional[List[str]] = None,
    telemetry: Optional[Dict[str, Any]] = None,
    queue_managed_rate_limits: bool = False,
    response_input: Optional[List[Dict[str, Any]]] = None,
    response_schema: Optional[Dict[str, Any]] = None,
    schema_name: str = "invoice_extraction",
    route: str = "current_full_document",
) -> Dict[str, Any]:
    audit = bool(audit_issues)
    if audit_issues:
        prompt += "\n\nREVISIÓN LIMITADA: corrige solo estas discrepancias: " + " | ".join(audit_issues)
    started = time.monotonic()
    model = _get_invoice_model()
    max_output_tokens = _get_invoice_max_output_tokens()
    timeout_seconds = _get_invoice_timeout_seconds()
    reasoning_effort = _get_invoice_reasoning_effort(audit)
    # Only persistent jobs own rate-limit retries. The direct fallback keeps
    # the SDK policy because it has no durable queue for a later retry.
    request_client = (
        client.with_options(max_retries=0)
        if queue_managed_rate_limits and hasattr(client, "with_options")
        else client
    )
    logger.info(
        "OpenAI invoice request retry policy: route=%s max_retries=%s",
        route,
        getattr(request_client, "max_retries", None),
    )
    request_input = response_input or _response_input_for_invoice(
        file_bytes, filename, mime_type, extracted_text, prompt
    )
    try:
        response = request_client.responses.create(
            model=model,
            reasoning={"effort": reasoning_effort},
            max_output_tokens=max_output_tokens,
            store=False,
            input=request_input,
            text={
                "format": {
                    "type": "json_schema",
                    "name": schema_name,
                    "strict": True,
                    "schema": response_schema or INVOICE_EXTRACTION_SCHEMA,
                }
            },
            timeout=timeout_seconds,
        )
    except Exception as exc:
        elapsed_ms = round((time.monotonic() - started) * 1000)
        _record_invoice_response_telemetry(
            telemetry,
            model=model,
            elapsed_ms=elapsed_ms,
            audit=audit,
        )
        rate_limit_error = getattr(openai, "RateLimitError", None) if openai is not None else None
        is_rate_limited = rate_limit_error is not None and isinstance(exc, rate_limit_error)
        metadata = _safe_openai_error_metadata(exc)
        if is_rate_limited:
            rate_limit_kind, retryable = _classify_rate_limit(metadata)
            metadata["rate_limit_kind"] = rate_limit_kind
            metadata["retryable"] = retryable
            status = "rate_limited"
        else:
            status = "timeout" if openai is not None and isinstance(exc, openai.APITimeoutError) else "api_error"
        logger.warning(
            "OpenAI Responses invoice API error: route=%s status=%s error_class=%s http_status=%s error_code=%s error_type=%s error_param=%s request_id=%s retry_after_seconds=%s rate_limit_kind=%s retryable=%s requests_limit=%s requests_remaining=%s requests_reset=%s tokens_limit=%s tokens_remaining=%s tokens_reset=%s audit=%s openai_response_elapsed_ms=%s timeout_seconds=%s reasoning_effort=%s",
            route,
            status,
            metadata.get("error_class"),
            metadata.get("http_status"),
            metadata.get("error_code"),
            metadata.get("error_type"),
            metadata.get("error_param"),
            metadata.get("request_id"),
            metadata.get("retry_after_seconds"),
            metadata.get("rate_limit_kind"),
            metadata.get("retryable"),
            metadata.get("requests_limit"),
            metadata.get("requests_remaining"),
            metadata.get("requests_reset"),
            metadata.get("tokens_limit"),
            metadata.get("tokens_remaining"),
            metadata.get("tokens_reset"),
            audit,
            elapsed_ms,
            timeout_seconds,
            reasoning_effort,
        )
        detail = metadata.get("error_code") or metadata.get("error_type") or type(exc).__name__
        raise InvoiceAnalysisResponseError(status, detail, metadata=metadata) from exc
    status = _response_value(response, "status")
    elapsed_ms = round((time.monotonic() - started) * 1000)
    _record_invoice_response_telemetry(
        telemetry,
        response=response,
        model=model,
        elapsed_ms=elapsed_ms,
        audit=audit,
    )
    _log_invoice_response_usage(
        response,
        status=status,
        max_output_tokens=max_output_tokens,
        reasoning_effort=reasoning_effort,
        audit=audit,
    )
    request_id = _response_value(response, "_request_id")
    if status != "completed":
        error = _response_value(response, "error") or {}
        incomplete = _response_value(response, "incomplete_details") or {}
        detail = _response_value(error, "code") or _response_value(incomplete, "reason") or status or "unknown"
        logger.warning(
            "OpenAI Responses invoice response_status=%s detail=%s refusal=false max_output_tokens=%s request_id=%s",
            status,
            detail,
            max_output_tokens,
            request_id,
        )
        raise InvoiceAnalysisResponseError(str(status or "unknown"), str(detail))
    if _response_has_refusal(response):
        logger.warning("OpenAI Responses invoice response_status=completed response_refusal=true")
        raise InvoiceAnalysisResponseError("refusal", "The model refused the document")
    raw_text = getattr(response, "output_text", "") or ""
    if not raw_text.strip():
        logger.warning("OpenAI Responses invoice response_status=completed response_refusal=false empty_output=true")
        raise InvoiceAnalysisResponseError("empty_output", "No structured output")
    parsing_started = time.monotonic()
    data = _extract_json(raw_text)
    if telemetry is not None:
        telemetry["parsing_ms"] = int(telemetry.get("parsing_ms") or 0) + round(
            (time.monotonic() - parsing_started) * 1000
        )
    if not data:
        logger.warning("OpenAI Responses invoice response_status=completed response_refusal=false valid_json=false")
        raise InvoiceAnalysisResponseError("invalid_structured_output", "JSON Schema output could not be parsed")
    logger.info(
        "OpenAI Responses invoice: route=%s model=%s endpoint=responses response_status=completed response_refusal=false direct_document=%s audit=%s openai_response_elapsed_ms=%s valid_json=true request_id=%s max_retries=%s timeout_seconds=%s reasoning_effort=%s",
        route,
        getattr(response, "model", model),
        bool(file_bytes),
        audit,
        elapsed_ms,
        request_id,
        getattr(request_client, "max_retries", None),
        timeout_seconds,
        reasoning_effort,
    )
    return data


def _invoice_analysis_failure(
    status: str,
    detail: Optional[str] = None,
    *,
    metadata: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    return {
        "analysis_status": "failed",
        "analysis_error": {
            "status": status,
            "detail": detail or None,
            "metadata": metadata or None,
        },
        "supplier": None,
        "provider_name": None,
        "client_name": None,
        "invoice_date": None,
        "payment_dates": [],
        "payment_date": None,
        "base_amount": None,
        "vat_rate": None,
        "vat_amount": None,
        "total_amount": None,
        "extraction_source": None,
        "confidence_score": None,
        "analysis_text": "",
        "validation": {"is_consistent": None, "difference": None},
    }


def analyze_invoice_v1(
    file_path: Optional[str] = None,
    file_bytes: Optional[bytes] = None,
    filename: Optional[str] = None,
    mime_type: Optional[str] = None,
    document_type: str = "expense",
    company_names: Optional[list] = None,
    known_suppliers: Optional[List[str]] = None,
    return_telemetry: bool = False,
    ocr_semaphore: Any = None,
    queue_managed_rate_limits: bool = False,
) -> Union[Dict[str, Any], Tuple[Dict[str, Any], Dict[str, Any]]]:
    analysis_started = time.monotonic()
    telemetry: Dict[str, Any] = {
        "preprocessing_ms": None,
        "ocr_ms": None,
        "openai_ms": 0,
        "openai_model": None,
        # Usage can be absent on an API response; retain that as unknown rather
        # than recording a false zero-cost analysis.
        "input_tokens": None,
        "output_tokens": None,
        "reasoning_tokens": None,
        "total_tokens": None,
        "ocr_used": False,
        "audit_used": False,
        "second_review_used": False,
        "processing_type": None,
        "ocr_concurrency_wait_ms": 0,
    }
    try:
        _get_invoice_model()
    except RuntimeError as exc:
        logger.warning("Invoice analysis configuration error: %s", exc)
        logger.info("Invoice analysis total: status=configuration_error audit=false ocr=false invoice_total_elapsed_ms=%s", round((time.monotonic() - analysis_started) * 1000))
        return _invoice_analysis_telemetry_result(
            _invoice_analysis_failure("configuration_error", str(exc)),
            telemetry,
            return_telemetry,
        )
    client = _get_client()
    company_names = company_names or []
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
        # The original PDF is sent directly to Responses. OCR is deliberately
        # deferred until an incomplete first pass needs audit support.
        extracted_text = embedded_text if is_significant else ""
        ocr_text = ""
        pdf_kind = "scanned" if is_scanned else "original"
        logger.info("PDF tratado como escaneado (%s): %s", filename, is_scanned)
        logger.info("Longitud texto extraido (%s): %s", filename, text_length)
        logger.info("Texto significativo (%s): %s", filename, is_significant)
        logger.info("OCR diferido (%s): %s", filename, is_scanned)
        telemetry["processing_type"] = "pdf_embedded_text" if is_significant else "pdf_direct_document"
    elif is_image:
        # Images are native multimodal inputs. OCR remains a later fallback.
        pdf_kind = "image"
        logger.info("OCR diferido para imagen (%s).", filename)
        telemetry["processing_type"] = "image_direct"
    else:
        logger.warning("Tipo de archivo no soportado (%s). Texto vacio enviado.", filename)
        telemetry["processing_type"] = "unknown"

    logger.info(
        "OCR usado (%s): %s | Longitud texto final: %s",
        filename,
        used_ocr,
        len(extracted_text.strip()),
    )
    # PDF text can contain HTML entities (for example, ``&#x20;`` for a
    # space). Normalize once before every extractor and the model see it.
    extracted_text = _normalize_ocr_amount_text(extracted_text)
    embedded_text = _normalize_ocr_amount_text(embedded_text)

    analysis_status = "ok"
    is_income = document_type == "income"
    is_rent_expense = document_type == "expense_rent"
    is_payroll_expense = document_type == "expense_payroll"
    is_other_expense = document_type == "expense_other"
    # The visual document is the primary extraction source. PDF text order is
    # frequently unrelated to its visual layout and may omit logos, headers or
    # columns where the issuer and fiscal totals are shown.
    needs_visual_document_check = (is_pdf or is_image) and analysis_status == "ok"
    logger.info(
        "Analisis documental multimodal (%s): visual=%s",
        filename,
        needs_visual_document_check,
    )
    if is_income:
        prompt = (
            "Analiza el siguiente texto extraido de una factura emitida (ingreso). "
            "Devuelve SOLO JSON valido con estas claves: "
            "client, invoice_date, payment_terms_days, payment_dates, totals, vat_breakdown. "
            "totals es un objeto con {base, vat, total} (pueden ser null). "
            "vat_breakdown es una lista de lineas IVA con {base, vat_amount} y opcional {rate}. "
            "Si hay varias lineas IVA, NO rellenes un vat_rate unico (deja rate en cada linea o null). "
            "payment_terms_days es el numero de dias si aparece una condicion tipo "
            "\"RECIBO X DIAS FECHA FACTURA\". "
            "Si hay payment_terms_days y invoice_date, devuelve payment_dates con invoice_date + X dias. "
            "payment_dates debe ser una lista (YYYY-MM-DD) y puede estar vacia. "
            "Usa null si no puedes inferir un dato con seguridad. "
            "No incluyas texto adicional fuera del JSON."
        )
    elif is_rent_expense:
        prompt = (
            "Analiza el siguiente texto extraido de un documento de gasto de alquiler o arrendamiento. "
            "Devuelve SOLO JSON valido con estas claves: "
            "supplier, invoice_date, payment_terms_days, payment_dates, withholding_amount, totals, vat_breakdown. "
            "totals es un objeto con {base, vat, total} (pueden ser null). "
            "vat_breakdown es una lista de lineas IVA con {base, vat_amount} y opcional {rate}. "
            "withholding_amount es la retencion a Hacienda si aparece; devuelve siempre su importe "
            "en positivo, aunque la linea IRPF lo muestre con signo menos, y nunca devuelvas el porcentaje. "
            "El supplier debe ser la razon social del arrendador o emisor, nunca el cliente/receptor. "
            "Prioriza base imponible, IVA, total, retencion y fecha de pago si aparecen. "
            "payment_terms_days es el numero de dias si aparece una condicion tipo "
            "\"RECIBO X DIAS FECHA FACTURA\". "
            "Si hay payment_terms_days y invoice_date, devuelve payment_dates con invoice_date + X dias. "
            "payment_dates debe ser una lista de fechas (YYYY-MM-DD) y puede estar vacia. "
            "Usa null si no puedes inferir un dato con seguridad. "
            "No incluyas texto adicional fuera del JSON."
        )
    elif is_payroll_expense:
        prompt = (
            "Analiza el siguiente texto extraido de una nomina o documento de coste laboral. "
            "Devuelve SOLO JSON valido con estas claves: "
            "employee_name, employer_name, payroll_period, gross_amount, payroll_total_deductions_amount, payroll_net_amount, payroll_employer_cost_amount, withholding_amount, invoice_date, payment_dates. "
            "employee_name es el trabajador de la nomina. employer_name es la empresa pagadora si aparece. "
            "gross_amount es el total devengado. payroll_total_deductions_amount es el total a deducir. "
            "payroll_net_amount es el liquido a percibir. payroll_employer_cost_amount es el coste empresa si aparece. "
            "withholding_amount es la retencion IRPF si aparece de forma explicita; devuelve el importe "
            "en positivo aunque el documento lo muestre con signo menos, nunca el porcentaje. Si no aparece, usa 0 o null. "
            "No uses claves de IVA salvo que el documento realmente las tenga; en nominas normales no hay IVA. "
            "Prioriza trabajador, periodo, fecha, bruto, deducciones, liquido y retencion. "
            "payment_dates debe ser una lista de fechas (YYYY-MM-DD) y puede estar vacia. "
            "Usa null si no puedes inferir un dato con seguridad. "
            "No incluyas texto adicional fuera del JSON."
        )
    elif is_other_expense:
        prompt = (
            "Analiza el siguiente texto extraido de un documento de gasto. "
            "Devuelve SOLO JSON valido con estas claves: "
            "supplier, invoice_date, payment_terms_days, payment_dates, withholding_amount, totals, vat_breakdown. "
            "totals es un objeto con {base, vat, total} (pueden ser null). "
            "vat_breakdown es una lista de lineas IVA con {base, vat_amount} y opcional {rate}. "
            "withholding_amount es la retencion a Hacienda si aparece; devuelve siempre su importe "
            "en positivo, aunque la linea IRPF lo muestre con signo menos, y nunca devuelvas el porcentaje. "
            "El supplier debe ser la razon social del emisor si aparece y no debe ser el cliente/receptor. "
            "Prioriza fecha, base, IVA, total y retencion. "
            "payment_dates debe ser una lista de fechas (YYYY-MM-DD) y puede estar vacia. "
            "Usa null si no puedes inferir un dato con seguridad. "
            "No incluyas texto adicional fuera del JSON."
        )
    else:
        prompt = (
            "Analiza el siguiente texto extraido de una factura recibida (gasto). "
            "Devuelve SOLO JSON valido con estas claves: "
            "supplier, invoice_date, payment_terms_days, payment_dates, withholding_amount, totals, vat_breakdown. "
            "totals es un objeto con {base, vat, total} (pueden ser null). "
            "vat_breakdown es una lista de lineas IVA con {base, vat_amount} y opcional {rate}. "
            "withholding_amount es la retencion a Hacienda si aparece; devuelve siempre su importe "
            "en positivo, aunque la linea IRPF lo muestre con signo menos, y nunca devuelvas el porcentaje. "
            "Si hay varias lineas IVA, NO rellenes un vat_rate unico (deja rate en cada linea o null). "
            "El supplier debe ser la razon social del emisor (forma juridica si aparece) "
            "y no debe ser el cliente/receptor. "
            "payment_terms_days es el numero de dias si aparece una condicion tipo "
            "\"RECIBO X DIAS FECHA FACTURA\". "
            "Si hay payment_terms_days y invoice_date, devuelve payment_dates con invoice_date + X dias. "
            "payment_dates debe ser una lista de fechas (YYYY-MM-DD) y puede estar vacia. "
            "Usa null si no puedes inferir un dato con seguridad. "
            "No incluyas texto adicional fuera del JSON."
        )

    prompt += (
        "\n\nREGLAS DE EXTRACCION: La imagen adjunta es el documento original y es la "
        "fuente principal; el texto extraido solo sirve como apoyo porque puede estar "
        "desordenado. Lee el encabezado, las tablas y el pie visualmente. No confundas "
        "al cliente, destinatario o dirección de entrega con el emisor/proveedor. No "
        "inventes ningún IVA, retención, fecha ni importe: usa null si no aparece de forma "
        "inequívoca. Para una factura recibida, total = base + IVA - retención. La retención "
        "siempre se devuelve como importe positivo. Comprueba internamente esa ecuación "
        "antes de responder. Incluye evidence como objeto opcional con textos breves "
        "literales del documento para supplier, invoice_date y totals."
    )
    prompt += (
        "\n\nINSTRUCCIONES TRANSVERSALES: Extrae la factura usando el documento original como fuente primaria y devuelve el schema solicitado. "
        "No inventes valores. El emisor no puede ser el cliente ni la dirección de entrega. "
        "Un vencimiento es una fecha prevista; NUNCA es un pago real. actual_payment_date debe ser null "
        "sin evidencia explícita de cobro/pago. payment.status no puede ser paid solo por existir o vencer una fecha. "
        "Las retenciones se devuelven siempre en valor absoluto positivo."
    )
    logger.info(
        "Invoice preprocessing: document_preprocessing_elapsed_ms=%s ocr=%s",
        round((time.monotonic() - analysis_started) * 1000),
        used_ocr,
    )
    telemetry["preprocessing_ms"] = round((time.monotonic() - analysis_started) * 1000)
    try:
        response_data = _call_invoice_responses(
            client,
            file_bytes=file_bytes,
            filename=filename,
            mime_type=mime_type,
            extracted_text=extracted_text,
            prompt=prompt,
            telemetry=telemetry,
            queue_managed_rate_limits=queue_managed_rate_limits,
        )
    except InvoiceAnalysisResponseError as exc:
        logger.warning("Invoice analysis stopped before parsing: status=%s detail=%s", exc.status, exc.detail)
        logger.info("Invoice analysis total: status=%s audit=false ocr=false invoice_total_elapsed_ms=%s", exc.status, round((time.monotonic() - analysis_started) * 1000))
        return _invoice_analysis_telemetry_result(
            _invoice_analysis_failure(
                exc.status, exc.detail, metadata=exc.metadata
            ),
            telemetry,
            return_telemetry,
        )
    except RuntimeError as exc:
        # Configuration errors must be explicit, not disguised as ambiguity.
        logger.warning("Invoice analysis configuration error: %s", exc)
        logger.info("Invoice analysis total: status=configuration_error audit=false ocr=false invoice_total_elapsed_ms=%s", round((time.monotonic() - analysis_started) * 1000))
        return _invoice_analysis_telemetry_result(
            _invoice_analysis_failure("configuration_error", str(exc)),
            telemetry,
            return_telemetry,
        )

    structured_data = response_data
    validation_started = time.monotonic()
    validation_issues = _validate_structured_invoice(structured_data, document_type)
    validation_elapsed_ms = round((time.monotonic() - validation_started) * 1000)
    audit_performed = _requires_invoice_audit(validation_issues)
    logger.info(
        "Invoice validation: validation_elapsed_ms=%s issues=%s audit=%s issues_detail=%s",
        validation_elapsed_ms,
        len(validation_issues),
        audit_performed,
        validation_issues,
    )
    audit_elapsed_ms = 0
    if audit_performed and (is_pdf or is_image) and not extracted_text:
        ocr_function = _extract_pdf_text_ocr_from_bytes if is_pdf else _extract_image_text_ocr_from_bytes
        ocr_slot_acquired = False
        ocr_slot_started = time.monotonic()
        if ocr_semaphore is not None:
            # OCR is the memory-heavy fallback. The shared worker semaphore
            # serializes only this operation, not the model analysis itself.
            ocr_semaphore.acquire()
            ocr_slot_acquired = True
        telemetry["ocr_concurrency_wait_ms"] = round(
            (time.monotonic() - ocr_slot_started) * 1000
        )
        ocr_started = time.monotonic()
        try:
            ocr_text, ocr_timed_out = _run_with_timeout(
                ocr_function, OCR_TIMEOUT_SECONDS, file_bytes
            )
        finally:
            if ocr_slot_acquired:
                ocr_semaphore.release()
        telemetry["ocr_ms"] = round((time.monotonic() - ocr_started) * 1000)
        if not ocr_timed_out and ocr_text:
            extracted_text = _normalize_ocr_amount_text(ocr_text)
            used_ocr = True
            telemetry["processing_type"] = f"{file_kind}_ocr_fallback"
            logger.info("OCR usado como fallback de auditoría (%s).", filename)
    if audit_performed:
        audit_started = time.monotonic()
        try:
            audited_data = _call_invoice_responses(
                client,
                file_bytes=file_bytes,
                filename=filename,
                mime_type=mime_type,
                extracted_text=extracted_text,
                prompt=prompt,
                audit_issues=validation_issues,
                telemetry=telemetry,
                queue_managed_rate_limits=queue_managed_rate_limits,
            )
        except InvoiceAnalysisResponseError as exc:
            logger.warning("Invoice audit failed safely: status=%s", exc.status)
            if exc.status == "rate_limited":
                logger.info(
                    "Invoice analysis total: status=%s audit=true ocr=%s invoice_total_elapsed_ms=%s",
                    exc.status,
                    used_ocr,
                    round((time.monotonic() - analysis_started) * 1000),
                )
                return _invoice_analysis_telemetry_result(
                    _invoice_analysis_failure(
                        exc.status, exc.detail, metadata=exc.metadata
                    ),
                    telemetry,
                    return_telemetry,
                )
            audited_data = None
        if audited_data:
            structured_data = audited_data
            audit_validation_started = time.monotonic()
            validation_issues = _validate_structured_invoice(structured_data, document_type)
            validation_elapsed_ms += round((time.monotonic() - audit_validation_started) * 1000)
        audit_elapsed_ms = round((time.monotonic() - audit_started) * 1000)
    logger.info(
        "Invoice audit: audit=%s audit_elapsed_ms=%s ocr=%s validation_elapsed_ms=%s",
        audit_performed,
        audit_elapsed_ms,
        used_ocr,
        validation_elapsed_ms,
    )
    telemetry["ocr_used"] = used_ocr
    telemetry["audit_used"] = audit_performed
    telemetry["second_review_used"] = audit_performed
    logger.info(
        "Validacion estructurada de factura (%s): valid=%s incidencias=%s segunda_revision=%s",
        filename,
        not validation_issues,
        len(validation_issues),
        audit_performed,
    )

    # Flatten the strict contract to the legacy response expected by the UI.
    supplier = structured_data.get("supplier") or {}
    customer = structured_data.get("customer") or {}
    invoice = structured_data.get("invoice") or {}
    totals = structured_data.get("totals") or {}
    evidence = structured_data.get("field_evidence") or {}
    installments = structured_data.get("installments") or []
    data = {
        "supplier": supplier.get("legal_name") or supplier.get("commercial_name"),
        "client": customer.get("legal_name") or customer.get("commercial_name"),
        "invoice_date": invoice.get("issue_date"),
        "invoice_number": invoice.get("invoice_number"),
        "payment_dates": [item.get("due_date") for item in installments if item.get("due_date")],
        "withholding_amount": totals.get("withholding"),
        "totals": {"base": totals.get("taxable_base"), "vat": totals.get("vat_amount"), "total": totals.get("total")},
        "vat_breakdown": [{"base": item.get("taxable_base"), "vat_amount": item.get("vat_amount"), "rate": item.get("vat_rate")} for item in structured_data.get("taxes") or []],
        "evidence": {key: (value or {}).get("evidence") for key, value in evidence.items() if isinstance(value, dict)},
    }
    raw_text = json.dumps(structured_data, ensure_ascii=False)
    if not data:
        logger.warning("Responses no devolvió una extracción válida (%s).", filename)

    provider_name = (
        data.get("supplier")
        or data.get("proveedor")
        or data.get("provider_name")
        or data.get("provider")
    )
    evidence_payload = data.get("evidence") if isinstance(data.get("evidence"), dict) else {}
    supplier_evidence = _pick_first_non_empty(
        evidence_payload.get("supplier"),
        evidence_payload.get("provider"),
        evidence_payload.get("emisor"),
    )
    client_name = (
        data.get("client")
        or data.get("cliente")
        or data.get("customer")
        or data.get("client_name")
    )
    model_issue_date = (
        data.get("invoice_date") or data.get("fecha_factura") or data.get("fecha")
    )
    issue_date_evidence = evidence_payload.get("issue_date")
    invoice_date, issue_date_correction, issue_date_review_reason = (
        _resolve_issue_date_from_evidence(model_issue_date, issue_date_evidence)
    )
    text_invoice_date = _extract_invoice_date_from_text(extracted_text)
    if text_invoice_date and invoice_date is None:
        invoice_date = text_invoice_date
    raw_payment_dates = (
        data.get("payment_dates")
        or data.get("fechas_pago")
        or data.get("fechas_vencimiento")
        or data.get("vencimientos")
    )
    single_payment_date_raw = (
        data.get("payment_date")
        or data.get("fecha_pago")
        or data.get("fecha_vencimiento")
        or data.get("vencimiento")
    )
    totals_payload = data.get("totals") if isinstance(data.get("totals"), dict) else {}
    base_amount = _normalize_amount(
        _pick_first_non_empty(
            totals_payload.get("base"),
            totals_payload.get("base_amount"),
            data.get("base_amount"),
            data.get("base_imponible"),
            data.get("base"),
        )
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
        _pick_first_non_empty(
            totals_payload.get("vat"),
            totals_payload.get("vat_amount"),
            data.get("vat_amount"),
            data.get("importe_iva"),
            data.get("iva_importe"),
        )
    )
    total_amount = _normalize_amount(
        _pick_first_non_empty(
            totals_payload.get("total"),
            totals_payload.get("total_amount"),
            data.get("total_amount"),
            data.get("total_factura"),
            data.get("total"),
        )
    )
    payroll_fields = _extract_payroll_fields_from_text(extracted_text) if is_payroll_expense else {}
    employee_name = (
        data.get("employee_name")
        or data.get("worker_name")
        or data.get("employee")
        or payroll_fields.get("employee_name")
    )
    payroll_period = (
        data.get("payroll_period")
        or data.get("period")
        or payroll_fields.get("payroll_period")
    )
    payroll_net_amount = _normalize_amount(
        _pick_first_non_empty(
            data.get("payroll_net_amount"),
            data.get("net_amount"),
            data.get("liquido_a_percibir"),
            data.get("liquido"),
            payroll_fields.get("payroll_net_amount"),
        )
    )
    payroll_total_deductions_amount = _normalize_amount(
        _pick_first_non_empty(
            data.get("payroll_total_deductions_amount"),
            data.get("total_deductions_amount"),
            data.get("deductions_amount"),
            payroll_fields.get("payroll_total_deductions_amount"),
        )
    )
    payroll_employer_cost_amount = _normalize_amount(
        _pick_first_non_empty(
            data.get("payroll_employer_cost_amount"),
            data.get("employer_cost_amount"),
            data.get("coste_empresa"),
            payroll_fields.get("payroll_employer_cost_amount"),
        )
    )
    gross_amount = _normalize_amount(
        _pick_first_non_empty(
            data.get("gross_amount"),
            data.get("total_devengado"),
            totals_payload.get("gross"),
            payroll_fields.get("gross_amount"),
            base_amount,
            total_amount,
        )
    )
    vat_breakdown = (
        data.get("vat_breakdown")
        or data.get("iva_breakdown")
        or data.get("vat_lines")
        or data.get("iva_lines")
        or []
    )
    withholding_amount = _normalize_amount(
        _pick_first_non_empty(
            data.get("withholding_amount"),
            data.get("retencion"),
            data.get("retencion_irpf"),
            data.get("retention_amount"),
        )
    )
    explicit_withholding_amount = _extract_explicit_withholding_amount_from_text(extracted_text)
    if explicit_withholding_amount is not None:
        withholding_amount = abs(explicit_withholding_amount)
    elif withholding_amount is not None:
        withholding_amount = abs(withholding_amount)
    amount_source = "llm" if data else "fallback"

    if company_names is None:
        company_names = []

    if document_type != "income":
        supplier_source_text = embedded_text if pdf_kind == "original" else extracted_text
        provider_name = provider_name.strip() if isinstance(provider_name, str) else provider_name
        if isinstance(provider_name, str):
            provider_name = _strip_inline_tax_id(provider_name)
        if provider_name is not None and not _is_valid_supplier(
            provider_name, company_names, supplier_source_text, require_tax_id=False
        ):
            if not _is_valid_visual_supplier(
                provider_name,
                company_names,
                supplier_evidence,
            ):
                provider_name = None
        if provider_name is None and analysis_status == "ok":
            learned_supplier = _match_known_supplier(
                supplier_source_text, known_suppliers, company_names
            )
            if learned_supplier:
                provider_name = learned_supplier
        if provider_name is None and analysis_status == "ok":
            heuristic_supplier = _extract_supplier_from_text(
                supplier_source_text,
                company_names,
            )
            if heuristic_supplier is not None and not _is_valid_supplier(
                heuristic_supplier, company_names, supplier_source_text, require_tax_id=True
            ):
                heuristic_supplier = None
            provider_name = heuristic_supplier
        if is_payroll_expense:
            employer_name = (
                data.get("employer_name")
                or data.get("company_name")
                or data.get("supplier")
                or provider_name
            )
            provider_name = employer_name.strip() if isinstance(employer_name, str) else employer_name
    else:
        client_source_text = embedded_text if pdf_kind == "original" else extracted_text
        client_name = client_name.strip() if isinstance(client_name, str) else client_name
        if isinstance(client_name, str):
            client_name = _strip_inline_tax_id(client_name)
        if client_name is not None and not _is_valid_client(
            client_name, company_names, client_source_text
        ):
            client_name = None
        if client_name is None and analysis_status == "ok":
            heuristic_client = _extract_client_from_text(
                client_source_text,
                company_names,
            )
            if heuristic_client is not None and not _is_valid_client(
                heuristic_client, company_names, client_source_text
            ):
                heuristic_client = None
            client_name = heuristic_client

    payment_dates, payment_terms_days = _resolve_payment_schedule(
        extracted_text,
        invoice_date,
        raw_payment_dates,
        single_payment_date_raw,
        _pick_first_non_empty(data.get("payment_terms_days"), data.get("payment_terms")),
        issue_date_corrected_from_evidence=issue_date_correction is not None,
    )
    payment_date = payment_dates[0] if payment_dates else None

    if is_payroll_expense:
        if not payment_dates and invoice_date:
            payment_dates = [invoice_date]
            payment_date = invoice_date
        if invoice_date is None and payroll_period:
            invoice_date = f"{payroll_period}-01"
        if gross_amount is not None:
            base_amount = gross_amount
            total_amount = gross_amount
        elif base_amount is not None:
            gross_amount = base_amount
            total_amount = base_amount
        vat_rate = 0
        vat_amount = 0
        vat_breakdown = []
        if (
            payroll_total_deductions_amount is None
            and gross_amount is not None
            and payroll_net_amount is not None
        ):
            payroll_total_deductions_amount = _round_amount(
                max(gross_amount - payroll_net_amount, 0)
            )
        if (
            payroll_net_amount is None
            and gross_amount is not None
            and payroll_total_deductions_amount is not None
        ):
            payroll_net_amount = _round_amount(
                max(gross_amount - payroll_total_deductions_amount, 0)
            )

    tax_summary = _extract_tax_summary_from_text(extracted_text)
    base_amount, vat_amount, total_amount, vat_rate, summary_source = _apply_tax_summary_override(
        extracted_text,
        base_amount,
        vat_amount,
        total_amount,
        vat_rate,
        tax_summary,
        withholding_amount,
    )
    if summary_source != "llm":
        amount_source = summary_source
    summary_breakdown = tax_summary.get("breakdown") if isinstance(tax_summary, dict) else None
    if summary_breakdown and (not vat_breakdown or summary_source != "llm"):
        vat_breakdown = summary_breakdown

    # Fallback to explicit amounts in text (e.g., "Total factura") if present.
    base_amount, vat_amount, total_amount, amount_source = _apply_text_amount_fallbacks(
        extracted_text,
        base_amount,
        vat_amount,
        total_amount,
        amount_source,
    )

    normalized = normalize_and_validate_amounts(
        {
            "analysis_status": analysis_status,
            "base_amount": base_amount,
            "vat_amount": vat_amount,
            "total_amount": total_amount,
            "vat_rate": vat_rate,
            "vat_breakdown": vat_breakdown,
            "withholding_amount": withholding_amount,
            "totals": totals_payload,
            "amount_source": amount_source,
        }
    )

    analysis_status = normalized.get("analysis_status") or analysis_status
    base_amount = normalized.get("base_amount")
    vat_amount = normalized.get("vat_amount")
    total_amount = normalized.get("total_amount")
    vat_rate = normalized.get("vat_rate")
    vat_breakdown = normalized.get("vat_breakdown") or []
    breakdown_warning = normalized.get("breakdown_warning")
    amount_source = normalized.get("amount_source") or amount_source or ("llm" if data else "fallback")

    rescued = _recover_from_tax_summary_if_needed(
        analysis_status,
        base_amount,
        vat_amount,
        total_amount,
        vat_rate,
        vat_breakdown,
        amount_source,
        tax_summary,
    )
    analysis_status = rescued.get("analysis_status") or analysis_status
    base_amount = rescued.get("base_amount")
    vat_amount = rescued.get("vat_amount")
    total_amount = rescued.get("total_amount")
    vat_rate = rescued.get("vat_rate")
    vat_breakdown = rescued.get("vat_breakdown") or []
    amount_source = rescued.get("amount_source") or amount_source
    if rescued.get("breakdown_warning") is not None:
        breakdown_warning = rescued.get("breakdown_warning")

    forced_summary = _force_tax_summary_result_if_available(
        analysis_status,
        base_amount,
        vat_amount,
        total_amount,
        vat_rate,
        vat_breakdown,
        amount_source,
        bool(breakdown_warning),
        tax_summary,
        withholding_amount,
    )
    analysis_status = forced_summary.get("analysis_status") or analysis_status
    base_amount = forced_summary.get("base_amount")
    vat_amount = forced_summary.get("vat_amount")
    total_amount = forced_summary.get("total_amount")
    vat_rate = forced_summary.get("vat_rate")
    vat_breakdown = forced_summary.get("vat_breakdown") or []
    amount_source = forced_summary.get("amount_source") or amount_source
    if forced_summary.get("breakdown_warning") is not None:
        breakdown_warning = forced_summary.get("breakdown_warning")

    total_amount = _apply_withholding_to_payable_total(
        base_amount,
        vat_amount,
        total_amount,
        withholding_amount,
    )

    explicit_exempt_base = _extract_explicit_vat_exemption_amount_from_text(extracted_text)
    if explicit_exempt_base is not None:
        # A stated exemption is stronger evidence than a model guess or the
        # default UI rate. Keep the payable amount separate from any possible
        # withholding and never manufacture a VAT quota for this document.
        base_amount = explicit_exempt_base
        total_amount = round(explicit_exempt_base - (withholding_amount or 0), 2)
        vat_rate = 0
        vat_amount = 0.0
        vat_breakdown = [
            {
                "base": _round_amount(base_amount),
                "vat_amount": 0.0,
                "rate": 0.0,
                "total": _round_amount(base_amount),
            }
        ]
        amount_source = "explicit_vat_exemption"
        analysis_status = "ok"
        breakdown_warning = False

    validation = _validate_math(base_amount, vat_amount, total_amount, withholding_amount)
    is_rectificativa = bool((base_amount is not None and base_amount < 0) or (total_amount is not None and total_amount < 0))
    review_reasons = _review_reasons_for_invoice(
        document_type=document_type,
        supplier=provider_name,
        base_amount=base_amount,
        vat_amount=vat_amount,
        total_amount=total_amount,
        withholding_amount=withholding_amount,
        breakdown_warning=bool(breakdown_warning),
    )
    if issue_date_review_reason:
        review_reasons.append(issue_date_review_reason)
    review_reasons = list(dict.fromkeys(validation_issues + review_reasons))
    if review_reasons and analysis_status == "ok":
        analysis_status = "needs_review"

    confidence_score = _confidence_score_for_source(amount_source)
    if confidence_score is not None and review_reasons:
        confidence_score = min(confidence_score, 0.6)
    logger.info("Fuente importes (%s): %s", filename, amount_source)
    logger.info(
        "Valores detectados (%s): proveedor=%s cliente=%s fecha=%s pago=%s base=%s iva_rate=%s iva_importe=%s retencion=%s total=%s",
        filename,
        provider_name,
        client_name,
        invoice_date,
        payment_date,
        base_amount,
        vat_rate,
        vat_amount,
        withholding_amount,
        total_amount,
    )

    result = {
        "analysis_status": analysis_status,
        "review_reasons": review_reasons,
        "supplier": provider_name,
        "provider_name": provider_name,
        "employee_name": employee_name.strip() if isinstance(employee_name, str) else employee_name,
        "payroll_period": payroll_period,
        "payroll_net_amount": payroll_net_amount,
        "payroll_total_deductions_amount": payroll_total_deductions_amount,
        "payroll_employer_cost_amount": payroll_employer_cost_amount,
        "client_name": client_name,
        "client": client_name,
        "invoice_date": invoice_date,
        "date_correction": issue_date_correction,
        "payment_terms_days": payment_terms_days,
        "payment_dates": payment_dates,
        "payment_date": payment_date,
        "base_amount": base_amount,
        "vat_rate": vat_rate,
        "vat_amount": vat_amount,
        "total_amount": total_amount,
        "vat_breakdown": vat_breakdown,
        "is_rectificativa": is_rectificativa,
        "withholding_amount": withholding_amount,
        "breakdown_warning": breakdown_warning,
        "extraction_source": amount_source,
        "confidence_score": confidence_score,
        "analysis_text": raw_text[:500],
        "validation": validation,
        "structured_extraction": structured_data,
        "payment_status": (structured_data.get("payment") or {}).get("status"),
    }
    logger.info(
        "Invoice analysis total: status=%s audit=%s ocr=%s invoice_total_elapsed_ms=%s",
        analysis_status,
        audit_performed,
        used_ocr,
        round((time.monotonic() - analysis_started) * 1000),
    )
    return _invoice_analysis_telemetry_result(result, telemetry, return_telemetry)


def _response_input_for_invoice_fast_text(prompt: str, compact_text: str) -> List[Dict[str, Any]]:
    return [
        {
            "role": "user",
            "content": [
                {"type": "input_text", "text": prompt},
                {
                    "type": "input_text",
                    "text": "DOCUMENTO DIGITAL (texto nativo por página):\n" + compact_text,
                },
            ],
        }
    ]


def _normalize_fast_text_tax_id(value: Any) -> Optional[str]:
    if value is None:
        return None
    normalized = re.sub(r"[^A-Za-z0-9]", "", str(value)).upper()
    return normalized or None


def _is_plausible_fast_text_tax_id(value: Optional[str]) -> bool:
    if value is None:
        return True
    return (
        8 <= len(value) <= 16
        and any(character.isalpha() for character in value)
        and any(character.isdigit() for character in value)
    )


def _normalize_fast_text_tax_lines(raw_taxes: Any) -> List[Dict[str, Optional[float]]]:
    lines: List[Dict[str, Optional[float]]] = []
    for line in raw_taxes if isinstance(raw_taxes, list) else []:
        if not isinstance(line, dict):
            continue
        base_amount = _money_decimal(line.get("taxable_base"))
        vat_amount = _money_decimal(line.get("vat_amount"))
        lines.append(
            {
                "base": _round_amount(float(base_amount)) if base_amount is not None else None,
                "rate": _normalize_rate(line.get("vat_rate")),
                "vat_amount": _round_amount(float(vat_amount)) if vat_amount is not None else None,
            }
        )
    return lines


def _fast_text_evidence_present(structured_data: Dict[str, Any], field: str) -> bool:
    evidence = (structured_data.get("field_evidence") or {}).get(field) or {}
    return bool(isinstance(evidence, dict) and str(evidence.get("evidence") or "").strip())


def _normalize_fast_text_company_context(
    company_context: Optional[Dict[str, Any]],
) -> Dict[str, Optional[str]]:
    """Normalize trusted Ledged company data used only to disambiguate parties."""
    context = company_context if isinstance(company_context, dict) else {}
    name = re.sub(r"\s+", " ", str(context.get("company_name") or "")).strip()
    tax_id = _normalize_fast_text_tax_id(context.get("company_tax_id"))
    return {
        "company_name": name[:255] or None,
        "company_tax_id": tax_id,
    }


def _fast_text_invoice_prompt(company_context: Optional[Dict[str, Any]]) -> str:
    """Build V2 instructions without treating Ledged company data as invoice evidence."""
    context = _normalize_fast_text_company_context(company_context)
    company_name = context.get("company_name") or "no disponible"
    company_tax_id = context.get("company_tax_id") or "no disponible"
    return (
        "Extrae exclusivamente los campos contables de esta factura digital usando el texto nativo "
        "por páginas. Devuelve el schema exacto. El proveedor/emisor es la contraparte que emite "
        "la factura y cobra; el cliente/receptor/comprador es quien la recibe y registra. "
        "CONTEXTO DE LEDGED PARA DISTINGUIR ROLES (no es evidencia documental ni obliga a "
        "inventar campos): la empresa que registra esta factura es "
        f"{json.dumps(company_name, ensure_ascii=False)}, con identificador fiscal "
        f"{json.dumps(company_tax_id, ensure_ascii=False)}. Si ese nombre o identificador "
        "aparece en el documento, normalmente identifica al destinatario/receptor/comprador y "
        "nunca debe clasificarse como proveedor/emisor. Si aparecen ambos identificadores, usa "
        "las etiquetas, la posición y el contexto documental para diferenciarlos. Si el documento "
        "es ambiguo, devuelve null para la parte que no puedas atribuir con seguridad en vez de "
        "asignar a la empresa registrada como proveedor. No inventes importes, IVA, retenciones, "
        "fechas o NIF. La retención se devuelve siempre como importe absoluto positivo. Verifica "
        "internamente: total = base + IVA + otros impuestos - retención. Incluye evidencia literal "
        "breve y el número de página para proveedor, número, fecha, líneas fiscales y totales. "
        "El número de factura debe ser exclusivamente el identificador etiquetado como número de "
        "factura. Nunca uses como número de factura un pedido, referencia de pedido o cliente, "
        "albarán, delivery note u otra referencia documental. Si no hay una etiqueta inequívoca de "
        "factura, devuelve null en lugar de escoger otra referencia."
    )


# V11 parses document identifiers through one typed grammar. Labels and source
# text are normalized only in memory; neither the compact document text nor its
# surrounding fragments are persisted by the shadow benchmark.
_FAST_TEXT_INVOICE_PARSER_REVISION = "v11"
_FAST_TEXT_REPRESENTATION_VERSION = "native_compact_v1"
_FAST_TEXT_SECONDARY_IDENTIFIER_TYPES = {
    "order_reference",
    "delivery_note",
    "customer_reference",
    "product_reference",
    "unknown",
}
_FAST_TEXT_NUMBER_MARKER_NORMALIZATION_PATTERN = re.compile(
    r"(?<!INVOICE )(?<!ORDER )(?<!DOCUMENT )(?<!CUSTOMER )"
    r"\bN\s*(?:[º°]\.?|O\.?|\.\s*(?:[º°O])|U(?:M(?:ERO)?)\.?)"
    r"(?=\s|:|#|$|(?:FACTURA|PEDIDO|ALBARAN|DOCUMENTO|DOC)\b)",
    flags=re.IGNORECASE,
)
_FAST_TEXT_IDENTIFIER_LABEL_GRAMMAR = (
    # Invoice labels. A bare FACTURA form remains strong only after the
    # extractor verifies one immediate structured identifier.
    ("invoice_number", "invoice_number_prefix", "strong", r"NUMERO\s+(?:DE\s+)?FACTURA"),
    ("invoice_number", "invoice_number_short_prefix", "strong", r"NUMERO\s+(?:DE\s+)?FACT\.?"),
    ("invoice_number", "invoice_number_suffix", "strong", r"FACTURA\s+NUMERO"),
    ("invoice_number", "invoice_id", "strong", r"ID\s+(?:DE\s+)?FACTURA"),
    ("invoice_number", "invoice_number_english", "strong", r"INVOICE\s+(?:NUMBER|NO\.?|#)"),
    ("invoice_number", "invoice_number_colon", "strong", r"FACTURA\s*:"),
    ("invoice_number", "invoice_context", "strong", r"FACTURA(?=\s|:|$)"),
    # Secondary document identifiers must remain typed so they cannot compete
    # with an explicit invoice label during reconciliation.
    ("order_reference", "order_prefix", "strong", r"NUMERO\s+(?:DE\s+)?PEDIDO"),
    ("order_reference", "order_suffix", "strong", r"PEDIDO\s+NUMERO"),
    ("order_reference", "order_english", "strong", r"ORDER\s+(?:NUMBER|NO\.?)"),
    ("order_reference", "order_context", "contextual", r"PEDIDO(?=\s|:|$)"),
    ("delivery_note", "delivery_prefix", "strong", r"NUMERO\s+(?:DE\s+)?ALBARAN"),
    ("delivery_note", "delivery_suffix", "strong", r"ALBARAN\s+NUMERO"),
    ("delivery_note", "delivery_english", "strong", r"DELIVERY\s+NOTE(?:\s+(?:NUMBER|NO\.?))?"),
    ("delivery_note", "delivery_context", "contextual", r"ALBARAN(?=\s|:|$)"),
    ("customer_reference", "customer_reference", "strong", r"CUSTOMER\s+REFERENCE"),
    ("customer_reference", "customer_number_prefix", "strong", r"CUSTOMER\s+(?:NUMBER|NO\.?)"),
    ("customer_reference", "customer_number_suffix", "strong", r"NUMERO\s+(?:DE\s+)?CLIENTE"),
    ("customer_reference", "reference_order", "strong", r"ORDER\s+REFERENCE"),
    ("customer_reference", "document_reference", "strong", r"DOCUMENT\s+REFERENCE"),
    ("customer_reference", "reference_spanish", "strong", r"REFERENCIA\s+(?:DE\s+)?(?:PEDIDO|CLIENTE)"),
    ("customer_reference", "reference_context", "contextual", r"(?:REFERENCIA|REF\.?)\s*(?=\s|:|$)"),
    ("product_reference", "product_reference", "strong", r"(?:PRODUCT|PRODUCTO|ARTICULO)\s+(?:REFERENCE|REFERENCIA|NUMBER|NUMERO)"),
    ("product_reference", "product_context", "contextual", r"(?:PRODUCTO|ARTICULO)(?=\s|:|$)"),
    # Generic document labels are intentionally not invoice labels.  They are
    # promoted only by the guarded rule in _inspect_fast_text_invoice_number_evidence.
    ("generic_document_number", "generic_document_prefix", "strong", r"NUMERO\s+(?:DEL\s+)?(?:DOCUMENTO|DOC\.?)"),
    ("generic_document_number", "generic_document_english", "strong", r"DOCUMENT\s+(?:NUMBER|NO\.?)"),
)
_FAST_TEXT_IDENTIFIER_COMPONENT_PATTERN = re.compile(
    r"(?<![A-Z0-9])([A-Z0-9]+(?:[._/\-][A-Z0-9]+)*)(?![A-Z0-9])"
)
_FAST_TEXT_LABEL_VALUE_SEPARATOR_PATTERN = re.compile(r"^[\s:#\-]*$")
_FAST_TEXT_RELATIVE_PAYMENT_TERMS_PATTERN = re.compile(
    r"(?<![0-9/])(?:RECIBO\s+)?([1-9][0-9]{0,2})\s+D[IÍ]AS\s+(?:A\s+)?FECHA\s+FACTURA\b",
    flags=re.IGNORECASE,
)
# Native invoices commonly render their issue date with a two-digit year.  The
# established Ledged date normalizer owns the century policy; this pattern only
# finds an explicitly invoice-labelled candidate for that normalizer.
_FAST_TEXT_INVOICE_DATE_PATTERN = re.compile(
    r"(?<!\d)(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})(?!\d)"
)
_FAST_TEXT_DATE_EXCLUSION_PATTERN = re.compile(
    r"\b(?:VENCIMIENTO|PAGO|ENTREGA|PEDIDO|ALBARAN|EXPEDICION|ENVIO|DUE|PAYMENT|DELIVERY|ORDER|SHIP(?:PING)?)\b"
)
_FAST_TEXT_INVOICE_DATE_STRONG_PATTERN = re.compile(
    r"\b(?:FECHA\s+(?:DE\s+)?FACTURA|FACTURA\b.*\bFECHA|INVOICE\s+DATE)\b"
)
_FAST_TEXT_INVOICE_DATE_CONTEXTUAL_PATTERN = re.compile(r"\bDE\s+FECHA\b")
_FAST_TEXT_DOCUMENT_INVOICE_PATTERN = re.compile(r"\b(?:FACTURA|INVOICE)\b")
_FAST_TEXT_DOCUMENT_NON_INVOICE_PATTERN = re.compile(
    r"\b(?:PROFORMA|PEDIDO|ALBARAN|DELIVERY|ORDER|PRESUPUESTO|QUOTE)\b"
)
_FAST_TEXT_PROVIDER_CONTEXT_PATTERN = re.compile(
    r"\b(?:PROVEEDOR|EMISOR(?:A)?|VENDEDOR|SUPPLIER|ISSUER|REMITENTE|"
    r"DATOS\s+DEL\s+EMISOR|FROM)\b"
)
_FAST_TEXT_RECIPIENT_CONTEXT_PATTERN = re.compile(
    r"\b(?:CLIENTE|RECEPTOR|DESTINATARIO|COMPRADOR|BILL\s+TO|SHIP\s+TO|"
    r"DIRECCION\s+DE\s+ENTREGA|DELIVERY|TO)\b"
)
_FAST_TEXT_TAX_ID_CONTEXT_PATTERN = re.compile(
    r"\b(?:N\s*\.?\s*I\s*\.?\s*F\s*\.?|C\s*\.?\s*I\s*\.?\s*F\s*\.?|"
    r"VAT|TAX\s*ID|IDENTIFICACION\s+FISCAL)\b"
)
_FAST_TEXT_TAX_ID_TOKEN_PATTERN = re.compile(
    r"(?<![A-Z0-9])(?:[A-Z]{2}[\s.-]?)?(?:[A-Z]\s*\d(?:[\s.-]?\d){6}[\s.-]?[A-Z0-9]|"
    r"\d(?:[\s.-]?\d){7}[\s.-]?[A-Z])(?![A-Z0-9])"
)
_FAST_TEXT_SPANISH_TAX_ID_PATTERN = re.compile(
    r"^(?:ES)?(?:[A-HJ-NP-SUVW]\d{7}[0-9A-J]|\d{8}[A-Z]|[XYZ]\d{7}[A-Z])$"
)
_FAST_TEXT_MONEY_TOKEN_PATTERN = re.compile(
    r"(?<![A-Z0-9])(?:EUR|USD|GBP|US\$|€|£)?\s*[-+]?"
    r"(?:\d{1,3}(?:[.,\s]\d{3})+|\d+)(?:[.,]\d{2})?\s*"
    r"(?:EUR|USD|GBP|US\$|€|£)?(?![A-Z0-9])",
    flags=re.IGNORECASE,
)
_FAST_TEXT_MONEY_CONTEXT_PATTERNS = {
    "base_amount": re.compile(r"\b(?:BASE\s+IMPONIBLE|BASE\b|SUBTOTAL|TAXABLE\s+BASE)\b"),
    "vat_amount": re.compile(r"\b(?:TOTAL\s+IVA|IVA|VAT|IMPUESTO)\b"),
    "withholding_amount": re.compile(r"\b(?:RETENCION|IRPF|WITHHOLDING)\b"),
    "other_taxes_amount": re.compile(
        r"\b(?:OTROS\s+IMPUESTOS|OTHER\s+TAXES|RECARGO(?:\s+DE\s+EQUIVALENCIA)?)\b"
    ),
    "total_amount": re.compile(
        r"\b(?:TOTAL\s+(?:FACTURA|A\s+PAGAR|GENERAL)|IMPORTE\s+TOTAL|TOTAL\s*[:#-])\b"
    ),
}
_FAST_TEXT_CURRENCY_MARKERS = {
    "EUR": re.compile(r"(?:€|\bEUR\b)", flags=re.IGNORECASE),
    "USD": re.compile(r"(?:US\$|\bUSD\b)", flags=re.IGNORECASE),
    "GBP": re.compile(r"(?:£|\bGBP\b)", flags=re.IGNORECASE),
}


def _normalize_fast_text_identifier(value: Any) -> Optional[str]:
    if value is None:
        return None
    normalized = unicodedata.normalize("NFKD", str(value))
    normalized = "".join(character for character in normalized if not unicodedata.combining(character))
    normalized = re.sub(r"[^A-Za-z0-9]", "", normalized).upper()
    return normalized or None


def _normalize_fast_text_parser_line(value: Any) -> str:
    """Normalize typography and label markers without retaining source text."""
    normalized = unicodedata.normalize("NFKD", str(value or ""))
    normalized = "".join(character for character in normalized if not unicodedata.combining(character))
    normalized = normalized.replace("\u00a0", " ")
    normalized = re.sub(r"[‐‑‒–—―]", "-", normalized)
    normalized = _FAST_TEXT_NUMBER_MARKER_NORMALIZATION_PATTERN.sub("NUMERO ", normalized)
    return re.sub(r"[ \t]+", " ", normalized).strip().upper()


def _fast_text_parser_lines(text: str) -> List[Tuple[int, str]]:
    """Return the normalized lines from V2's in-memory document representation."""
    return [
        (index, normalized)
        for index, raw_line in enumerate((text or "").splitlines())
        if (normalized := _normalize_fast_text_parser_line(raw_line))
    ]


def _fast_text_document_representation(text: str) -> Dict[str, Any]:
    """Create the shared V2 text projection used by the model verifier.

    ``text`` is exactly the page-labelled compact string sent to Responses API.
    This helper only creates in-memory normalized lines and token locations for
    deterministic matching; it neither extracts the PDF again nor persists text.
    """
    lines = _fast_text_parser_lines(text)
    tokens: List[Dict[str, Any]] = []
    for line_index, (_, line) in enumerate(lines):
        for match in re.finditer(r"[A-Z0-9]+", line):
            tokens.append(
                {
                    "line_index": line_index,
                    "start": match.start(),
                    "end": match.end(),
                    "value": match.group(0),
                }
            )
    return {
        "version": _FAST_TEXT_REPRESENTATION_VERSION,
        "lines": lines,
        "tokens": tokens,
    }


def _fast_text_identifier_components(value: str) -> List[re.Match]:
    return list(_FAST_TEXT_IDENTIFIER_COMPONENT_PATTERN.finditer(value or ""))


def _should_join_fast_text_identifier_components(first: str, second: str) -> bool:
    """Allow only bounded series/year splits, never two unrelated document IDs."""
    if not any(character.isdigit() for character in second):
        return False
    if first.isalpha():
        return len(first) <= 4
    return bool(
        re.fullmatch(r"(?:\d{2,4}[A-Z]{1,6}|[A-Z]{1,6}\d{2,4})", first)
        and second.isdigit()
    )


def _extract_fast_text_identifier_value(
    value: str, *, require_single_identifier: bool = False
) -> Optional[str]:
    """Extract one compact or compound identifier without consuming description text."""
    matches = _fast_text_identifier_components(value)
    if not matches or (
        matches[0].start() != 0 and not (value or "")[: matches[0].start()].strip(" :#-") == ""
    ):
        return None
    components = [match.group(1) for match in matches]
    if not components:
        return None
    first = components[0]
    candidate_parts = [first]
    next_index = 1
    if len(components) > 1 and _should_join_fast_text_identifier_components(
        first, components[1]
    ):
        candidate_parts.append(components[1])
        next_index = 2
    candidate = " ".join(candidate_parts)
    if (
        not any(character.isdigit() for character in candidate)
        or _normalize_fast_text_date(candidate) is not None
    ):
        return None
    if require_single_identifier:
        remaining = components[next_index:]
        if any(any(character.isdigit() for character in item) for item in remaining):
            return None
    return candidate


def _extract_fast_text_typed_identifiers(text: str) -> List[Dict[str, Any]]:
    """Extract typed candidates using the shared V11 document-label grammar."""
    lines = _fast_text_parser_lines(text)
    candidates: List[Dict[str, Any]] = []
    seen = set()
    for candidate_type, label_type, strength, expression in _FAST_TEXT_IDENTIFIER_LABEL_GRAMMAR:
        label_pattern = re.compile(
            r"(?<![A-Z0-9])(?:" + expression + r")(?![A-Z0-9])"
        )
        for line_position, (_, line) in enumerate(lines):
            for match in label_pattern.finditer(line):
                remainder = line[match.end() :]
                requires_single_identifier = (
                    strength == "contextual"
                    or label_type in {"invoice_number_colon", "invoice_context"}
                )
                candidate = _extract_fast_text_identifier_value(
                    remainder, require_single_identifier=requires_single_identifier
                )
                if (
                    candidate is None
                    and _FAST_TEXT_LABEL_VALUE_SEPARATOR_PATTERN.fullmatch(remainder or "")
                    and line_position + 1 < len(lines)
                ):
                    candidate = _extract_fast_text_identifier_value(
                        lines[line_position + 1][1],
                        require_single_identifier=requires_single_identifier,
                    )
                normalized = _normalize_fast_text_identifier(candidate)
                if not candidate or not normalized:
                    continue
                key = (candidate_type, normalized)
                if key in seen:
                    continue
                seen.add(key)
                candidates.append(
                    {
                        "visible_value": candidate,
                        "normalized_value": normalized,
                        "type": candidate_type,
                        "candidate_type": candidate_type,
                        "label_type": label_type,
                        "evidence_strength": strength,
                        "position": line_position,
                    }
                )
    return candidates


def _fast_text_document_has_invoice_context(text: str) -> bool:
    """Require a document-level invoice anchor before promoting a generic number."""
    for position, (_, line) in enumerate(_fast_text_parser_lines(text)):
        if position > 12:
            break
        if re.search(r"\b(?:FACTURA|INVOICE)\b", line) and not re.search(
            r"\b(?:PROFORMA|PEDIDO|ALBARAN|DELIVERY|ORDER)\b", line
        ):
            return True
    return False


def _fast_text_has_guarded_generic_invoice_context(
    text: str, generic_candidates: List[Dict[str, Any]]
) -> bool:
    """Require a nearby standalone invoice heading before promoting a generic ID.

    A generic document number is weaker than an explicit invoice label.  A
    document merely mentioning ``FACTURA`` elsewhere must not turn it into an
    invoice number: the number must be in the first header lines immediately
    following a standalone invoice heading.
    """
    if not generic_candidates:
        return False
    header_positions = {
        position
        for position, (_, line) in enumerate(_fast_text_parser_lines(text))
        if re.fullmatch(r"(?:FACTURA|INVOICE)\s*[:#-]?", line)
    }
    return any(
        candidate.get("position") in {header_position + 1, header_position + 2}
        for candidate in generic_candidates
        for header_position in header_positions
    )


def _limited_invoice_parser_candidates(
    candidates: List[Dict[str, Any]], candidate_types: set[str], limit: int = 3
) -> List[Dict[str, str]]:
    """Expose only bounded structured identifiers, never surrounding document text."""
    return [
        {
            "type": candidate["type"],
            "normalized_value": str(candidate["normalized_value"])[:128],
            "label_type": str(candidate["label_type"])[:64],
            "strength": str(candidate["evidence_strength"])[:24],
        }
        for candidate in candidates
        if candidate["type"] in candidate_types
    ][:limit]


def _fast_text_tokens_are_contiguous(
    representation: Dict[str, Any], previous: Dict[str, Any], current: Dict[str, Any]
) -> bool:
    """Allow typography-only gaps within a single document identifier."""
    line_delta = current["line_index"] - previous["line_index"]
    lines = representation["lines"]
    if line_delta == 0:
        gap = lines[previous["line_index"]][1][previous["end"] : current["start"]]
    elif line_delta == 1:
        gap = (
            lines[previous["line_index"]][1][previous["end"] :]
            + "\n"
            + lines[current["line_index"]][1][: current["start"]]
        )
    else:
        return False
    return not any(character.isalnum() for character in gap)


def _fast_text_model_match_has_complete_boundaries(
    representation: Dict[str, Any], start_index: int, end_index: int
) -> bool:
    """Reject prefix/suffix fragments while allowing label words around an ID."""
    tokens = representation["tokens"]
    first = tokens[start_index]
    last = tokens[end_index]
    if start_index > 0:
        previous = tokens[start_index - 1]
        if _fast_text_tokens_are_contiguous(representation, previous, first) and any(
            character.isdigit() for character in previous["value"]
        ):
            return False
    if end_index + 1 < len(tokens):
        following = tokens[end_index + 1]
        if _fast_text_tokens_are_contiguous(representation, last, following) and any(
            character.isdigit() for character in following["value"]
        ):
            return False
    return True


def _find_fast_text_model_value_matches(
    text: str,
    normalized_model_value: Optional[str],
    representation: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    """Find full IDs by canonical token sequence in V2's shared document text.

    The sequence retains every alphanumeric component and may only bridge
    typography-only gaps. It therefore accepts ``SF - 012516`` for
    ``SF012516`` but never accepts ``0625`` as ``2025IR0625``.
    """
    if not normalized_model_value:
        return []
    representation = representation or _fast_text_document_representation(text)
    tokens = representation["tokens"]
    matches: List[Dict[str, Any]] = []
    for start_index, first in enumerate(tokens):
        sequence = ""
        for end_index in range(start_index, min(len(tokens), start_index + 16)):
            current = tokens[end_index]
            if end_index > start_index and not _fast_text_tokens_are_contiguous(
                representation, tokens[end_index - 1], current
            ):
                break
            sequence += current["value"]
            if not normalized_model_value.startswith(sequence):
                break
            if sequence != normalized_model_value:
                continue
            if not _fast_text_model_match_has_complete_boundaries(
                representation, start_index, end_index
            ):
                break
            matches.append(
                {
                    "start_line": first["line_index"],
                    "start": first["start"],
                    "end_line": current["line_index"],
                    "end": current["end"],
                    "method": "exact" if start_index == end_index else "canonical_sequence",
                }
            )
            break
    return matches


def _fast_text_label_is_locally_associated_with_match(
    lines: List[Tuple[int, str]],
    label_match: re.Match,
    match: Dict[str, Any],
    line_index: int,
    *,
    allow_reversed_layout: bool,
) -> bool:
    """Require an adjacent label/ID relationship with no intervening token.

    PyMuPDF can invert nearby header fragments in its reading order.  Accept the
    invoice label immediately before *or* after the complete matched ID, but
    never bridge an unrelated word, a non-adjacent line or a second identifier.
    """
    if match["start_line"] == line_index and label_match.end() <= match["start"]:
        gap = lines[line_index][1][label_match.end() : match["start"]]
    elif match["start_line"] == line_index + 1:
        gap = (
            lines[line_index][1][label_match.end() :]
            + "\n"
            + lines[line_index + 1][1][: match["start"]]
        )
    elif (
        allow_reversed_layout
        and match["end_line"] == line_index
        and match["end"] <= label_match.start()
    ):
        gap = lines[line_index][1][match["end"] : label_match.start()]
    elif allow_reversed_layout and match["end_line"] + 1 == line_index:
        gap = (
            lines[match["end_line"]][1][match["end"] :]
            + "\n"
            + lines[line_index][1][: label_match.start()]
        )
    else:
        return False
    return not any(character.isalnum() for character in gap)


def _find_fast_text_model_value_contexts(
    text: str,
    normalized_model_value: Optional[str],
    matches: List[Dict[str, Any]],
    representation: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    """Associate complete canonical matches with the nearest typed document label."""
    if not normalized_model_value or not matches:
        return []
    lines = (representation or _fast_text_document_representation(text))["lines"]
    contexts: List[Dict[str, Any]] = []
    seen = set()
    for candidate_type, label_type, strength, expression in _FAST_TEXT_IDENTIFIER_LABEL_GRAMMAR:
        label_pattern = re.compile(r"(?<![A-Z0-9])(?:" + expression + r")(?![A-Z0-9])")
        for line_index, (_, line) in enumerate(lines):
            for label_match in label_pattern.finditer(line):
                for match in matches:
                    if not _fast_text_label_is_locally_associated_with_match(
                        lines,
                        label_match,
                        match,
                        line_index,
                        allow_reversed_layout=(
                            candidate_type == "invoice_number"
                            and strength == "strong"
                            and label_type != "invoice_context"
                        ),
                    ):
                        continue
                    key = (candidate_type, label_type, line_index, match["start_line"], match["start"])
                    if key in seen:
                        continue
                    seen.add(key)
                    contexts.append(
                        {
                            "type": candidate_type,
                            "candidate_type": candidate_type,
                            "normalized_value": normalized_model_value,
                            "label_type": label_type,
                            "evidence_strength": strength,
                            "position": line_index,
                        }
                    )
    return contexts


def _inspect_fast_text_model_invoice_number_evidence(
    text: str, normalized_model_value: Optional[str]
) -> Dict[str, Any]:
    """Verify V11's model value against the exact compact model input text."""
    representation = _fast_text_document_representation(text)
    matches = _find_fast_text_model_value_matches(
        text, normalized_model_value, representation
    )
    contexts = _find_fast_text_model_value_contexts(
        text, normalized_model_value, matches, representation
    )
    invoice_contexts = [context for context in contexts if context["type"] == "invoice_number"]
    generic_contexts = [
        context for context in contexts if context["type"] == "generic_document_number"
    ]
    secondary_contexts = [
        context for context in contexts if context["type"] in _FAST_TEXT_SECONDARY_IDENTIFIER_TYPES
    ]
    guarded_generic_contexts = (
        generic_contexts
        if _fast_text_has_guarded_generic_invoice_context(text, generic_contexts)
        else []
    )
    if len(matches) > 1:
        selected = None
        status = "ambiguous_match"
    elif invoice_contexts and secondary_contexts:
        selected = None
        status = "conflicting_identifier_context"
    elif invoice_contexts:
        selected = invoice_contexts[0]
        status = "strong_invoice_label"
    elif guarded_generic_contexts:
        selected = guarded_generic_contexts[0]
        status = "guarded_generic_document_number"
    elif secondary_contexts:
        selected = secondary_contexts[0]
        status = "secondary_identifier"
    elif matches:
        selected = None
        status = "unlabeled"
    else:
        selected = None
        status = "not_found"
    return {
        "text_representation_version": _FAST_TEXT_REPRESENTATION_VERSION,
        "model_value_found": bool(matches),
        "model_value_found_in_native_text": bool(matches),
        "model_value_match_count": len(matches),
        "model_value_match_method": (
            "exact"
            if any(match["method"] == "exact" for match in matches)
            else "canonical_sequence"
            if matches
            else "none"
        ),
        "model_value_invoice_context_status": status,
        "context_label_type": selected.get("label_type") if selected else None,
        "model_value_context_label_type": selected.get("label_type") if selected else None,
        "model_value_context_candidate_type": selected.get("type") if selected else None,
        "selected_context": selected,
        "contexts": contexts,
    }


def _build_fast_text_invoice_parser_diagnostics(
    evidence: Dict[str, Any],
    *,
    selected_candidate_type: Optional[str],
    action: str,
    ambiguity_reason: Optional[str],
    model_normalized_value: Optional[str] = None,
    selected_candidate_normalized_value: Optional[str] = None,
    conflict_reason: Optional[str] = None,
    model_evidence: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    candidates = evidence.get("candidates") or []
    model_evidence = model_evidence if isinstance(model_evidence, dict) else {}
    label_first_candidates = _limited_invoice_parser_candidates(
        candidates, {"invoice_number", "generic_document_number"}
    )
    selected_candidate = next(
        (
            candidate
            for candidate in label_first_candidates
            if candidate["type"] == selected_candidate_type
            and candidate["normalized_value"] == selected_candidate_normalized_value
        ),
        None,
    )
    if selected_candidate is None and isinstance(model_evidence.get("selected_context"), dict):
        selected_context = model_evidence["selected_context"]
        selected_candidate = {
            "type": selected_context.get("type"),
            "normalized_value": selected_context.get("normalized_value"),
            "label_type": selected_context.get("label_type"),
            "strength": selected_context.get("evidence_strength"),
        }
    secondary_counts = {
        candidate_type: sum(1 for candidate in candidates if candidate["type"] == candidate_type)
        for candidate_type in sorted(_FAST_TEXT_SECONDARY_IDENTIFIER_TYPES)
    }
    secondary_counts["generic_document_number"] = sum(
        1 for candidate in candidates if candidate["type"] == "generic_document_number"
    )
    return {
        "parser_revision": _FAST_TEXT_INVOICE_PARSER_REVISION,
        "candidate_detected": bool(candidates),
        # Keep invoice_candidates for older diagnostic readers; V11 names the
        # same bounded label-first evidence explicitly.
        "invoice_candidates": label_first_candidates,
        "label_first_candidates": label_first_candidates,
        "secondary_candidate_counts": secondary_counts,
        "selected_candidate_type": selected_candidate_type,
        "selected_candidate": selected_candidate,
        "model_normalized_value": model_normalized_value,
        "selected_candidate_normalized_value": selected_candidate_normalized_value,
        "text_representation_version": model_evidence.get("text_representation_version"),
        "model_value_found": bool(model_evidence.get("model_value_found")),
        "model_value_found_in_native_text": bool(
            model_evidence.get("model_value_found_in_native_text")
        ),
        "model_value_match_count": int(model_evidence.get("model_value_match_count") or 0),
        "model_value_match_method": model_evidence.get("model_value_match_method"),
        "model_value_invoice_context_status": model_evidence.get(
            "model_value_invoice_context_status"
        ),
        "context_label_type": model_evidence.get("context_label_type"),
        "model_value_context_label_type": model_evidence.get(
            "model_value_context_label_type"
        ),
        "label_first_candidate": label_first_candidates[0] if label_first_candidates else None,
        "reconciliation_action": action,
        "ambiguity_reason": ambiguity_reason,
        "conflict_reason": conflict_reason,
    }


def _inspect_fast_text_invoice_number_evidence(text: str) -> Dict[str, Any]:
    """Classify invoice evidence without allowing secondary IDs to outrank it."""
    candidates = _extract_fast_text_typed_identifiers(text)
    invoice_candidates = [candidate for candidate in candidates if candidate["type"] == "invoice_number"]
    strong_invoice_candidates = [
        candidate
        for candidate in invoice_candidates
        if candidate.get("evidence_strength") == "strong"
    ]
    contextual_invoice_candidates = [
        candidate
        for candidate in invoice_candidates
        if candidate.get("evidence_strength") != "strong"
    ]
    generic_candidates = [
        candidate for candidate in candidates if candidate["type"] == "generic_document_number"
    ]
    reference_candidates = [
        candidate for candidate in candidates if candidate["type"] in _FAST_TEXT_SECONDARY_IDENTIFIER_TYPES
    ]
    strong_invoice_by_normalized = {
        candidate["normalized_value"]: candidate for candidate in strong_invoice_candidates
    }
    contextual_invoice_by_normalized = {
        candidate["normalized_value"]: candidate for candidate in contextual_invoice_candidates
    }
    reference_identifiers = {
        candidate["normalized_value"] for candidate in reference_candidates
    }
    # Only strong labels can establish or contradict the invoice number.  A
    # contextual "FACTURA <id>" occurrence remains diagnostic evidence, but
    # cannot make a unique strong labelled invoice candidate ambiguous.
    if len(strong_invoice_by_normalized) == 1:
        invoice_candidate = next(iter(strong_invoice_by_normalized.values()))
        return {
            "status": "unambiguous",
            "invoice_number": invoice_candidate["visible_value"],
            "invoice_candidates": list(strong_invoice_by_normalized.values()),
            "reference_identifiers": reference_identifiers,
            "candidates": candidates,
            "selected_candidate_type": "invoice_number",
        }
    if len(strong_invoice_by_normalized) > 1:
        return {
            "status": "ambiguous",
            "invoice_number": None,
            "invoice_candidates": list(strong_invoice_by_normalized.values()),
            "reference_identifiers": reference_identifiers,
            "candidates": candidates,
            "selected_candidate_type": None,
        }
    if len(contextual_invoice_by_normalized) == 1:
        invoice_candidate = next(iter(contextual_invoice_by_normalized.values()))
        return {
            "status": "contextual",
            "invoice_number": invoice_candidate["visible_value"],
            "invoice_candidates": list(contextual_invoice_by_normalized.values()),
            "reference_identifiers": reference_identifiers,
            "candidates": candidates,
            "selected_candidate_type": "invoice_number",
        }
    if len(contextual_invoice_by_normalized) > 1:
        return {
            "status": "contextual_ambiguous",
            "invoice_number": None,
            "invoice_candidates": list(contextual_invoice_by_normalized.values()),
            "reference_identifiers": reference_identifiers,
            "candidates": candidates,
            "selected_candidate_type": None,
        }
    generic_by_normalized = {
        candidate["normalized_value"]: candidate for candidate in generic_candidates
    }
    if len(generic_by_normalized) == 1 and _fast_text_has_guarded_generic_invoice_context(
        text, list(generic_by_normalized.values())
    ):
        generic_candidate = next(iter(generic_by_normalized.values()))
        return {
            "status": "generic_unambiguous",
            "invoice_number": generic_candidate["visible_value"],
            "invoice_candidates": [],
            "generic_candidates": list(generic_by_normalized.values()),
            "reference_identifiers": reference_identifiers,
            "candidates": candidates,
            "selected_candidate_type": "generic_document_number",
        }
    if len(generic_by_normalized) > 1 and _fast_text_has_guarded_generic_invoice_context(
        text, list(generic_by_normalized.values())
    ):
        return {
            "status": "generic_ambiguous",
            "invoice_number": None,
            "invoice_candidates": [],
            "generic_candidates": list(generic_by_normalized.values()),
            "reference_identifiers": reference_identifiers,
            "candidates": candidates,
            "selected_candidate_type": None,
        }
    return {
        "status": "missing",
        "invoice_number": None,
        "invoice_candidates": [],
        "generic_candidates": list(generic_by_normalized.values()),
        "reference_identifiers": reference_identifiers,
        "candidates": candidates,
        "selected_candidate_type": None,
    }


def _reconcile_fast_text_invoice_number(
    invoice_number: Optional[str], document_text: str, *, document_layout: Optional[Dict[str, Any]] = None
) -> Tuple[Optional[str], List[str], List[str], str, Dict[str, Any]]:
    """Confirm V11's full model value before using label-first evidence.

    Native PDF text can place unrelated header values next to an invoice label.
    A unique complete model identifier directly associated with a typed invoice
    label therefore takes precedence over label-first candidates. The local
    association accepts either reading-order direction to tolerate layout
    artifacts, while multiple matches and contradictory typed contexts remain
    fail-closed fallback evidence.
    """
    evidence = _inspect_fast_text_invoice_number_evidence(document_text)
    normalized_model_number = _normalize_fast_text_identifier(invoice_number)
    model_evidence = _inspect_fast_text_model_invoice_number_evidence(
        document_text, normalized_model_number
    )

    context_status = model_evidence["model_value_invoice_context_status"]

    # A unique complete model value immediately tied to a typed, strong invoice
    # label is stronger than unrelated label-first candidates elsewhere in the
    # compact reading order. This is the only path that can override a remote
    # candidate ambiguity; an identifier with local secondary context still
    # falls through to the fail-closed conflict branch below.
    if (
        context_status in {"strong_invoice_label", "guarded_generic_document_number"}
        and evidence["status"] != "ambiguous"
    ):
        selected_context = model_evidence.get("selected_context") or {}
        diagnostics = _build_fast_text_invoice_parser_diagnostics(
            evidence,
            selected_candidate_type=selected_context.get("type"),
            action=(
                "confirmed_model_value_from_strong_invoice_label"
                if context_status == "strong_invoice_label"
                else "confirmed_model_value_from_guarded_generic_document_number"
            ),
            ambiguity_reason=None,
            model_normalized_value=normalized_model_number,
            selected_candidate_normalized_value=normalized_model_number,
            model_evidence=model_evidence,
        )
        return invoice_number, [], [], "confirmed", diagnostics

    if evidence["status"] in {"ambiguous", "contextual_ambiguous", "generic_ambiguous"}:
        diagnostics = _build_fast_text_invoice_parser_diagnostics(
            evidence,
            selected_candidate_type=None,
            action="review_ambiguous_candidates",
            ambiguity_reason=evidence["status"],
            model_normalized_value=normalized_model_number,
            conflict_reason="multiple_incompatible_invoice_candidates",
            model_evidence=model_evidence,
        )
        return invoice_number, [], ["invoice_number_evidence_ambiguous"], "ambiguous", diagnostics

    if context_status in {"ambiguous_match", "conflicting_identifier_context"}:
        diagnostics = _build_fast_text_invoice_parser_diagnostics(
            evidence,
            selected_candidate_type=None,
            action=(
                "review_ambiguous_complete_model_value_matches"
                if context_status == "ambiguous_match"
                else "review_conflicting_model_value_contexts"
            ),
            ambiguity_reason=(
                "multiple_complete_model_value_matches"
                if context_status == "ambiguous_match"
                else None
            ),
            model_normalized_value=normalized_model_number,
            conflict_reason=(
                None
                if context_status == "ambiguous_match"
                else "model_value_has_invoice_and_secondary_contexts"
            ),
            model_evidence=model_evidence,
        )
        issue = (
            "invoice_number_evidence_ambiguous"
            if context_status == "ambiguous_match"
            else "invoice_number_conflicts_with_explicit_label"
        )
        status = "ambiguous" if context_status == "ambiguous_match" else "conflict"
        return invoice_number, [], [issue], status, diagnostics

    expected_number = evidence.get("invoice_number")
    normalized_expected_number = _normalize_fast_text_identifier(expected_number)
    has_unique_strong_invoice_candidate = evidence["status"] == "unambiguous"

    if context_status == "secondary_identifier":
        if has_unique_strong_invoice_candidate and normalized_expected_number:
            diagnostics = _build_fast_text_invoice_parser_diagnostics(
                evidence,
                selected_candidate_type="invoice_number",
                action="corrected_secondary_model_value_from_explicit_invoice_label",
                ambiguity_reason=None,
                model_normalized_value=normalized_model_number,
                selected_candidate_normalized_value=normalized_expected_number,
                model_evidence=model_evidence,
            )
            return (
                expected_number,
                ["invoice_number_corrected_from_explicit_label"],
                [],
                "confirmed",
                diagnostics,
            )
        diagnostics = _build_fast_text_invoice_parser_diagnostics(
            evidence,
            selected_candidate_type=None,
            action="review_model_value_in_secondary_context",
            ambiguity_reason=None,
            model_normalized_value=normalized_model_number,
            selected_candidate_normalized_value=normalized_expected_number,
            conflict_reason="model_value_has_only_secondary_context",
            model_evidence=model_evidence,
        )
        return None, [], ["invoice_number_is_non_invoice_reference"], "conflict", diagnostics

    if evidence["status"] == "contextual":
        diagnostics = _build_fast_text_invoice_parser_diagnostics(
            evidence,
            selected_candidate_type="invoice_number",
            action="review_contextual_invoice_candidate",
            ambiguity_reason=None,
            model_normalized_value=normalized_model_number,
            selected_candidate_normalized_value=normalized_expected_number,
            conflict_reason="invoice_label_not_strong_enough_for_confirmation",
            model_evidence=model_evidence,
        )
        return invoice_number, [], ["invoice_number_evidence_contextual"], "review", diagnostics

    if evidence["status"] == "generic_unambiguous":
        diagnostics = _build_fast_text_invoice_parser_diagnostics(
            evidence,
            selected_candidate_type="generic_document_number",
            action=(
                "review_generic_document_number_missing"
                if normalized_model_number is None
                else "review_generic_document_number_mismatch"
            ),
            ambiguity_reason=None,
            model_normalized_value=normalized_model_number,
            selected_candidate_normalized_value=normalized_expected_number,
            conflict_reason=(
                "missing_model_invoice_number_for_guarded_generic_candidate"
                if normalized_model_number is None
                else "model_value_differs_from_guarded_generic_document_number"
            ),
            model_evidence=model_evidence,
        )
        issue = (
            "invoice_number_evidence_missing"
            if normalized_model_number is None
            else "invoice_number_conflicts_with_explicit_label"
        )
        status = "missing" if normalized_model_number is None else "conflict"
        return invoice_number, [], [issue], status, diagnostics

    if has_unique_strong_invoice_candidate and normalized_model_number is None:
        diagnostics = _build_fast_text_invoice_parser_diagnostics(
            evidence,
            selected_candidate_type="invoice_number",
            action="corrected_missing_model_value_from_explicit_invoice_label",
            ambiguity_reason=None,
            model_normalized_value=normalized_model_number,
            selected_candidate_normalized_value=normalized_expected_number,
            model_evidence=model_evidence,
        )
        return (
            expected_number,
            ["invoice_number_corrected_from_explicit_label"],
            [],
            "confirmed",
            diagnostics,
        )

    # V14 only supplements a missing text-order relationship. The complete
    # model value must already have one native-text occurrence; layout merely
    # proves that exact value belongs to a nearby strong invoice label.
    if evidence["status"] == "missing" and context_status == "unlabeled":
        spatial_verification = _verify_fast_text_layout_invoice_number(
            invoice_number,
            document_layout,
            model_value_found_in_native_text=bool(model_evidence.get("model_value_found")),
            model_value_match_count=int(model_evidence.get("model_value_match_count") or 0),
        )
        if spatial_verification.get("status") == "confirmed":
            diagnostics = _build_fast_text_invoice_parser_diagnostics(
                evidence,
                selected_candidate_type="invoice_number",
                action="confirmed_model_value_from_spatial_invoice_label",
                ambiguity_reason=None,
                model_normalized_value=normalized_model_number,
                selected_candidate_normalized_value=normalized_model_number,
                model_evidence=model_evidence,
            )
            diagnostics["context_label_type"] = spatial_verification.get("context_type")
            diagnostics["model_value_context_label_type"] = spatial_verification.get("context_type")
            diagnostics["model_value_invoice_context_status"] = "strong_spatial_invoice_label"
            return invoice_number, [], [], "confirmed", diagnostics

    if evidence["status"] == "missing":
        diagnostics = _build_fast_text_invoice_parser_diagnostics(
            evidence,
            selected_candidate_type=None,
            action=(
                "review_unlabeled_model_value"
                if context_status == "unlabeled"
                else "review_missing_invoice_evidence"
            ),
            ambiguity_reason=None,
            model_normalized_value=normalized_model_number,
            conflict_reason=(
                "model_value_not_associated_with_invoice_label"
                if context_status == "unlabeled"
                else "no_confirmable_invoice_candidate"
            ),
            model_evidence=model_evidence,
        )
        return invoice_number, [], ["invoice_number_evidence_missing"], "missing", diagnostics

    diagnostics = _build_fast_text_invoice_parser_diagnostics(
        evidence,
        selected_candidate_type="invoice_number",
        action="review_conflicting_invoice_number",
        ambiguity_reason=None,
        model_normalized_value=normalized_model_number,
        selected_candidate_normalized_value=normalized_expected_number,
        conflict_reason=(
            "model_value_differs_from_unique_strong_invoice_candidate"
            if has_unique_strong_invoice_candidate
            else "model_value_differs_from_contextual_invoice_candidate"
        ),
        model_evidence=model_evidence,
    )
    return (
        invoice_number,
        [],
        ["invoice_number_conflicts_with_explicit_label"],
        "conflict",
        diagnostics,
    )


def _normalize_fast_text_date(value: Any) -> Optional[str]:
    """Keep V2 date normalization local so the V1 date pipeline is unchanged."""
    if value is None:
        return None
    return _normalize_date(str(value).replace(".", "/"))


def _inspect_fast_text_invoice_date_evidence(
    document_text: str,
    *,
    supplier_tax_id: Optional[str] = None,
    registered_company_tax_id: Optional[str] = None,
) -> Dict[str, Any]:
    """Find one invoice-context date from a bounded labelled relationship.

    A label and its date may be split by the PDF reading order.  Accept only a
    date on the label line or immediately following a strong invoice-date
    label; this is deliberately narrower than a document-wide date search.
    """
    candidates: Dict[str, Dict[str, Any]] = {}
    spanish_fiscal_context = _fast_text_spanish_fiscal_context(
        document_text,
        supplier_tax_id=supplier_tax_id,
        registered_company_tax_id=registered_company_tax_id,
    )
    has_invoice_context = _fast_text_document_has_invoice_context(document_text)
    lines = _fast_text_parser_lines(document_text)
    for position, (_, line) in enumerate(lines):
        if _FAST_TEXT_DATE_EXCLUSION_PATTERN.search(line):
            continue
        strong_context = bool(_FAST_TEXT_INVOICE_DATE_STRONG_PATTERN.search(line))
        contextual_date = bool(_FAST_TEXT_INVOICE_DATE_CONTEXTUAL_PATTERN.search(line))
        previous_line = lines[position - 1][1] if position else ""
        previous_is_strong_label = bool(
            _FAST_TEXT_INVOICE_DATE_STRONG_PATTERN.search(previous_line)
        ) and not _FAST_TEXT_DATE_EXCLUSION_PATTERN.search(previous_line)
        if not strong_context and not previous_is_strong_label and not (
            contextual_date and has_invoice_context and position <= 12
        ):
            continue
        for match in _FAST_TEXT_INVOICE_DATE_PATTERN.finditer(line):
            normalized = _normalize_fast_text_date(match.group(0))
            if not normalized:
                continue
            if previous_is_strong_label and not strong_context:
                match_method = "next_line_strong_invoice_date_label"
            elif strong_context:
                match_method = "same_line_strong_invoice_date_label"
            else:
                match_method = "contextual_invoice_date_label"
            # Spain's fiscal documents use the established DD/MM/YY policy,
            # but only after the document itself demonstrates Spanish fiscal
            # context. Other documents retain the fail-closed ambiguity rule.
            ambiguous_numeric = (
                len(match.group(3)) == 2
                and int(match.group(1)) <= 12
                and not spanish_fiscal_context
            )
            existing = candidates.get(normalized)
            if existing is None:
                candidates[normalized] = {
                    "match_method": match_method,
                    "ambiguous_numeric": ambiguous_numeric,
                }
            elif not ambiguous_numeric:
                # A four-digit or day-unambiguous occurrence of the same date
                # is stronger than an earlier ambiguous two-digit rendering.
                existing["ambiguous_numeric"] = False
                existing["match_method"] = match_method
    if len(candidates) == 1:
        invoice_date, candidate = next(iter(candidates.items()))
        if candidate["ambiguous_numeric"]:
            return {
                "status": "ambiguous",
                "invoice_date": None,
                "candidate_count": 1,
                "match_method": "ambiguous_two_digit_invoice_date",
                "spanish_fiscal_context": spanish_fiscal_context,
            }
        return {
            "status": "unambiguous",
            "invoice_date": invoice_date,
            "candidate_count": 1,
            "match_method": candidate["match_method"],
            "spanish_fiscal_context": spanish_fiscal_context,
        }
    if len(candidates) > 1:
        return {
            "status": "ambiguous",
            "invoice_date": None,
            "candidate_count": len(candidates),
            "match_method": "multiple_invoice_date_candidates",
            "spanish_fiscal_context": spanish_fiscal_context,
        }
    return {
        "status": "missing",
        "invoice_date": None,
        "candidate_count": 0,
        "match_method": "invoice_date_label",
        "spanish_fiscal_context": spanish_fiscal_context,
    }


def _reconcile_fast_text_invoice_date(
    invoice_date: Optional[str],
    document_text: str,
    *,
    supplier_tax_id: Optional[str] = None,
    registered_company_tax_id: Optional[str] = None,
    document_layout: Optional[Dict[str, Any]] = None,
) -> Tuple[Optional[str], List[str], List[str], Dict[str, Any]]:
    """Use only one explicit native-text invoice date as deterministic evidence."""
    evidence = _inspect_fast_text_invoice_date_evidence(
        document_text,
        supplier_tax_id=supplier_tax_id,
        registered_company_tax_id=registered_company_tax_id,
    )
    normalized_model_date = _normalize_fast_text_date(invoice_date)
    diagnostics = {
        "status": evidence["status"],
        "candidate_count": evidence["candidate_count"],
        "match_method": evidence["match_method"],
        "spanish_fiscal_context": evidence["spanish_fiscal_context"],
        "model_matches_evidence": bool(
            normalized_model_date and normalized_model_date == evidence.get("invoice_date")
        ),
        "reconciliation_action": "kept_model_date",
    }
    if evidence["status"] == "ambiguous":
        diagnostics["reconciliation_action"] = "review_ambiguous_invoice_dates"
        return normalized_model_date, [], ["invoice_date_evidence_ambiguous"], diagnostics
    if evidence["status"] != "unambiguous":
        spatial_evidence = _fast_text_layout_invoice_date_candidate(
            document_layout,
            spanish_fiscal_context=evidence.get("spanish_fiscal_context"),
        )
        if spatial_evidence["status"] == "ambiguous":
            diagnostics.update(
                {
                    "status": "ambiguous",
                    "candidate_count": spatial_evidence["candidate_count"],
                    "match_method": spatial_evidence["match_method"],
                    "reconciliation_action": "review_ambiguous_spatial_invoice_dates",
                }
            )
            return normalized_model_date, [], ["invoice_date_evidence_ambiguous"], diagnostics
        if spatial_evidence["status"] == "unambiguous":
            diagnostics.update(
                {
                    "status": "unambiguous",
                    "candidate_count": 1,
                    "match_method": spatial_evidence["match_method"],
                    "model_matches_evidence": bool(
                        normalized_model_date
                        and normalized_model_date == spatial_evidence["invoice_date"]
                    ),
                    "reconciliation_action": (
                        "confirmed_model_date_from_spatial_invoice_label"
                        if normalized_model_date == spatial_evidence["invoice_date"]
                        else "spatial_invoice_date_differs_from_model"
                    ),
                }
            )
            if normalized_model_date == spatial_evidence["invoice_date"]:
                return (
                    normalized_model_date,
                    ["invoice_date_confirmed_from_spatial_label"],
                    [],
                    diagnostics,
                )
            return normalized_model_date, [], [], diagnostics
        diagnostics["reconciliation_action"] = "no_native_invoice_date_evidence"
        return normalized_model_date, [], [], diagnostics
    if normalized_model_date == evidence["invoice_date"]:
        diagnostics["reconciliation_action"] = "confirmed_exact_invoice_date"
        return normalized_model_date, [], [], diagnostics
    diagnostics["reconciliation_action"] = "corrected_from_explicit_invoice_date"
    diagnostics["model_matches_evidence"] = True
    return (
        evidence["invoice_date"],
        ["invoice_date_corrected_from_explicit_label"],
        [],
        diagnostics,
    )


def _extract_fast_text_relative_payment_terms_days(text: str) -> Optional[int]:
    """Reuse the V1 rule first, then accept the same unambiguous phrase without RECIBO."""
    v1_terms = extract_payment_terms_days(text)
    if v1_terms is not None and 0 < v1_terms <= 365:
        return v1_terms
    matches = {
        int(match.group(1))
        for match in _FAST_TEXT_RELATIVE_PAYMENT_TERMS_PATTERN.finditer(text or "")
        if 0 < int(match.group(1)) <= 365
    }
    return next(iter(matches)) if len(matches) == 1 else None


def _reconcile_fast_text_payment_dates(
    payment_dates: List[str], invoice_date: Optional[str], document_text: str
) -> Tuple[List[str], List[str]]:
    """Prioritize explicit due dates, then derive only exact relative invoice terms."""
    explicit_due_dates = _find_due_dates_in_due_context(document_text)
    if explicit_due_dates:
        return explicit_due_dates, []
    payment_terms_days = _extract_fast_text_relative_payment_terms_days(document_text)
    if payment_terms_days is None or not invoice_date:
        return payment_dates, []
    try:
        due_date = (date.fromisoformat(invoice_date) + timedelta(days=payment_terms_days)).isoformat()
    except ValueError:
        return payment_dates, []
    return [due_date], ["due_date_derived_from_payment_terms"]


_FAST_TEXT_VERIFICATION_STATES = {"confirmed", "review", "contradiction", "not_applicable"}
_FAST_TEXT_FAST_PATH_CRITICAL_FIELDS = (
    "document_type",
    "supplier_tax_id",
    "invoice_number",
    "invoice_date",
    "base_amount",
    "vat_amount",
    "vat_breakdown",
    "withholding_amount",
    "other_taxes_amount",
    "total_amount",
    "currency",
)


def _fast_text_verification(
    status: str,
    *,
    match_method: str,
    match_count: int = 0,
    context_type: Optional[str] = None,
    reason: Optional[str] = None,
) -> Dict[str, Any]:
    """Return bounded, source-free evidence suitable for shadow persistence."""
    if status not in _FAST_TEXT_VERIFICATION_STATES:
        raise ValueError(f"Unsupported V2 document verification status: {status}")
    return {
        "status": status,
        "match_method": match_method[:64],
        "match_count": max(0, min(int(match_count or 0), 999)),
        "context_type": (context_type or "")[:64] or None,
        "reason": (reason or "")[:128] or None,
    }


def _fast_text_lines_with_positions(text: str) -> List[str]:
    return [line for _, line in _fast_text_parser_lines(text)]


def _fast_text_local_context(lines: List[str], line_index: int) -> str:
    """Use only a small line window; never infer document-wide relationships."""
    start = max(0, line_index - 1)
    end = min(len(lines), line_index + 2)
    return "\n".join(lines[start:end])


_FAST_TEXT_LAYOUT_STRONG_TOTAL_PATTERN = re.compile(
    r"\b(?:TOTAL\s+(?:FACTURA|A\s+PAGAR|GENERAL)|IMPORTE\s+TOTAL|GRAND\s+TOTAL|AMOUNT\s+DUE)\b"
)
_FAST_TEXT_LAYOUT_RATE_HEADER_PATTERN = re.compile(
    r"\b(?:TIPO|TASA|RATE|PORCENTAJE|%)\b"
)


def _fast_text_layout_rows(layout: Optional[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not isinstance(layout, dict):
        return []
    rows: List[Dict[str, Any]] = []
    for page in layout.get("pages") or []:
        if isinstance(page, dict):
            rows.extend(row for row in page.get("rows") or [] if isinstance(row, dict))
    return rows


def _fast_text_layout_page_rows(layout: Optional[Dict[str, Any]], page_number: int) -> List[Dict[str, Any]]:
    if not isinstance(layout, dict):
        return []
    for page in layout.get("pages") or []:
        if isinstance(page, dict) and page.get("page") == page_number:
            return [row for row in page.get("rows") or [] if isinstance(row, dict)]
    return []


def _fast_text_layout_rows_are_related(first: Dict[str, Any], second: Dict[str, Any]) -> bool:
    """Allow only immediately adjacent visual rows with local geometric affinity."""
    if first.get("page") != second.get("page") or abs(first.get("index", 0) - second.get("index", 0)) != 1:
        return False
    vertical_gap = max(0.0, max(first["y0"], second["y0"]) - min(first["y1"], second["y1"]))
    local_height = max(float(first["median_token_height"]), float(second["median_token_height"]), 0.1)
    if vertical_gap > local_height * 3:
        return False
    first_width = max(float(first["x1"]) - float(first["x0"]), 0.1)
    second_width = max(float(second["x1"]) - float(second["x0"]), 0.1)
    overlap = max(0.0, min(first["x1"], second["x1"]) - max(first["x0"], second["x0"]))
    if overlap / min(first_width, second_width) >= 0.25:
        return True
    first_center = (first["x0"] + first["x1"]) / 2
    second_center = (second["x0"] + second["x1"]) / 2
    return abs(first_center - second_center) <= max(first_width, second_width) * 0.8


def _fast_text_layout_local_rows(layout: Optional[Dict[str, Any]], row: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows = _fast_text_layout_page_rows(layout, row.get("page"))
    local_rows = [row]
    for candidate_index in (row.get("index", 0) - 1, row.get("index", 0) + 1):
        if not 0 <= candidate_index < len(rows):
            continue
        candidate = rows[candidate_index]
        if _fast_text_layout_rows_are_related(row, candidate):
            local_rows.append(candidate)
    return local_rows


def _fast_text_layout_row_fragments(row: Dict[str, Any]) -> List[Dict[str, Any]]:
    fragments: List[Dict[str, Any]] = []
    offset = 0
    for token in row.get("tokens") or []:
        normalized = _normalize_fast_text_parser_line(token.get("text"))
        if not normalized:
            continue
        if fragments:
            offset += 1
        fragments.append(
            {
                "start": offset,
                "end": offset + len(normalized),
                "text": normalized,
                "token": token,
            }
        )
        offset += len(normalized)
    return fragments


def _fast_text_layout_row_normalized_text(row: Dict[str, Any]) -> str:
    return " ".join(fragment["text"] for fragment in _fast_text_layout_row_fragments(row))


def _fast_text_layout_union_bbox(tokens: List[Dict[str, Any]]) -> Optional[Tuple[float, float, float, float]]:
    if not tokens:
        return None
    return (
        min(token["x0"] for token in tokens),
        min(token["y0"] for token in tokens),
        max(token["x1"] for token in tokens),
        max(token["y1"] for token in tokens),
    )


def _fast_text_layout_label_boxes(row: Dict[str, Any], pattern: re.Pattern) -> List[Tuple[float, float, float, float]]:
    fragments = _fast_text_layout_row_fragments(row)
    text = " ".join(fragment["text"] for fragment in fragments)
    boxes = []
    for match in pattern.finditer(text):
        matched_tokens = [
            fragment["token"]
            for fragment in fragments
            if fragment["start"] < match.end() and fragment["end"] > match.start()
        ]
        bbox = _fast_text_layout_union_bbox(matched_tokens)
        if bbox is not None:
            boxes.append(bbox)
    return boxes


def _fast_text_layout_money_matches(row: Dict[str, Any], proposed: Decimal) -> List[Dict[str, Any]]:
    matches = []
    for token in row.get("tokens") or []:
        value = _parse_fast_text_money(token.get("text"))
        if value is not None and abs(value - proposed) <= Decimal("0.01"):
            matches.append(token)
    return matches


def _fast_text_layout_money_values(row: Dict[str, Any]) -> List[Tuple[Dict[str, Any], Decimal]]:
    values = []
    for token in row.get("tokens") or []:
        value = _parse_fast_text_money(token.get("text"))
        if value is not None:
            values.append((token, value))
    return values


def _fast_text_layout_box_centers_align(
    label_box: Tuple[float, float, float, float], value_box: Tuple[float, float, float, float]
) -> bool:
    label_width = max(label_box[2] - label_box[0], 0.1)
    value_width = max(value_box[2] - value_box[0], 0.1)
    overlap = max(0.0, min(label_box[2], value_box[2]) - max(label_box[0], value_box[0]))
    if overlap > 0:
        return True
    label_center = (label_box[0] + label_box[2]) / 2
    value_center = (value_box[0] + value_box[2]) / 2
    return abs(label_center - value_center) <= max(label_width, value_width) * 0.75


def _fast_text_layout_column_label_boxes(field: str, row: Dict[str, Any]) -> List[Tuple[float, float, float, float]]:
    if field == "total_amount":
        return _fast_text_layout_label_boxes(row, re.compile(r"\bTOTAL\b"))
    return _fast_text_layout_label_boxes(row, _FAST_TEXT_MONEY_CONTEXT_PATTERNS[field])


def _fast_text_layout_accounting_header_fields(row: Dict[str, Any]) -> set[str]:
    fields = set()
    for field in ("base_amount", "vat_amount", "total_amount"):
        if _fast_text_layout_column_label_boxes(field, row):
            fields.add(field)
    return fields


def _fast_text_layout_monetary_associations(
    field: str, proposed: Decimal, layout: Optional[Dict[str, Any]]
) -> int:
    """Count unambiguous local label/value or header/column proof paths."""
    associations = set()
    pattern = _FAST_TEXT_MONEY_CONTEXT_PATTERNS[field]
    for row in _fast_text_layout_rows(layout):
        label_boxes = _fast_text_layout_label_boxes(row, pattern)
        if field == "total_amount" and label_boxes and not _FAST_TEXT_LAYOUT_STRONG_TOTAL_PATTERN.search(
            _fast_text_layout_row_normalized_text(row)
        ):
            label_boxes = []
        same_row_matches = _fast_text_layout_money_matches(row, proposed)
        if label_boxes and len(same_row_matches) == 1:
            associations.add((row["page"], row["index"], row["index"], same_row_matches[0]["x0"]))
        for related in _fast_text_layout_local_rows(layout, row):
            if related is row or not label_boxes:
                continue
            related_values = _fast_text_layout_money_values(related)
            if len(related_values) == 1 and abs(related_values[0][1] - proposed) <= Decimal("0.01"):
                associations.add((row["page"], row["index"], related["index"], related_values[0][0]["x0"]))

        header_fields = _fast_text_layout_accounting_header_fields(row)
        if len(header_fields) < 2 or field not in header_fields:
            continue
        for related in _fast_text_layout_local_rows(layout, row):
            if related is row:
                continue
            matching_values = _fast_text_layout_money_matches(related, proposed)
            for label_box in _fast_text_layout_column_label_boxes(field, row):
                aligned = [
                    token
                    for token in matching_values
                    if _fast_text_layout_box_centers_align(
                        label_box, (token["x0"], token["y0"], token["x1"], token["y1"])
                    )
                ]
                if len(aligned) == 1:
                    associations.add((row["page"], row["index"], related["index"], aligned[0]["x0"]))
    return len(associations)


def _verify_fast_text_layout_monetary_field(
    field: str, proposed_value: Any, layout: Optional[Dict[str, Any]]
) -> Dict[str, Any]:
    proposed = _money_decimal(proposed_value)
    if proposed is None:
        return _fast_text_verification("review", match_method="spatial_monetary", reason=f"{field}_missing")
    associations = _fast_text_layout_monetary_associations(field, proposed, layout)
    if associations == 1:
        return _fast_text_verification(
            "confirmed",
            match_method="spatial_label_value",
            match_count=1,
            context_type=field,
        )
    return _fast_text_verification(
        "review",
        match_method="spatial_label_value",
        match_count=associations,
        context_type=field,
        reason=(f"{field}_spatial_association_ambiguous" if associations else f"{field}_spatial_context_missing"),
    )


def _fast_text_layout_merge_verification(
    linear: Dict[str, Any], spatial: Optional[Dict[str, Any]], *, allow_not_applicable_refutation: bool = False
) -> Dict[str, Any]:
    """Let layout prove a review, never erase a deterministic contradiction."""
    if not isinstance(spatial, dict) or linear.get("status") == "contradiction":
        return linear
    if linear.get("status") == "review" and spatial.get("status") in {"confirmed", "contradiction"}:
        return spatial
    if (
        allow_not_applicable_refutation
        and linear.get("status") == "not_applicable"
        and spatial.get("status") == "contradiction"
    ):
        return spatial
    return linear


def _fast_text_layout_identifier_matches(
    layout: Optional[Dict[str, Any]], normalized_value: Optional[str]
) -> List[Dict[str, Any]]:
    """Find a complete canonical identifier within one visual row only."""
    if not normalized_value:
        return []
    matches = []
    for row in _fast_text_layout_rows(layout):
        tokens = row.get("tokens") or []
        normalized_tokens = [
            _normalize_fast_text_identifier(token.get("text")) or "" for token in tokens
        ]
        for start_index, first in enumerate(normalized_tokens):
            sequence = ""
            for end_index in range(start_index, min(len(normalized_tokens), start_index + 16)):
                current = normalized_tokens[end_index]
                if not current:
                    break
                sequence += current
                if not normalized_value.startswith(sequence):
                    break
                if sequence != normalized_value:
                    continue
                if start_index and any(character.isdigit() for character in normalized_tokens[start_index - 1]):
                    break
                if end_index + 1 < len(normalized_tokens) and any(
                    character.isdigit() for character in normalized_tokens[end_index + 1]
                ):
                    break
                matches.append(
                    {
                        "row": row,
                        "start_index": start_index,
                        "end_index": end_index,
                        "bbox": _fast_text_layout_union_bbox(tokens[start_index : end_index + 1]),
                    }
                )
                break
    return matches


def _fast_text_layout_strong_invoice_label_rows(
    layout: Optional[Dict[str, Any]], match: Dict[str, Any]
) -> List[Dict[str, Any]]:
    rows = []
    for row in _fast_text_layout_local_rows(layout, match["row"]):
        row_text = _fast_text_layout_row_normalized_text(row)
        if _FAST_TEXT_DOCUMENT_NON_INVOICE_PATTERN.search(row_text):
            continue
        for candidate_type, label_type, strength, expression in _FAST_TEXT_IDENTIFIER_LABEL_GRAMMAR:
            if candidate_type != "invoice_number" or strength != "strong":
                continue
            if re.search(r"(?<![A-Z0-9])(?:" + expression + r")(?![A-Z0-9])", row_text):
                rows.append(
                    {"row": row, "label_type": label_type}
                )
                break
    return rows


def _verify_fast_text_layout_invoice_number(
    invoice_number: Optional[str],
    layout: Optional[Dict[str, Any]],
    *,
    model_value_found_in_native_text: bool,
    model_value_match_count: int,
) -> Dict[str, Any]:
    """Prove an already-found complete model ID against a nearby invoice label."""
    normalized = _normalize_fast_text_identifier(invoice_number)
    if not normalized or not model_value_found_in_native_text or model_value_match_count != 1:
        return _fast_text_verification(
            "review", match_method="spatial_invoice_identifier", reason="invoice_number_not_spatially_eligible"
        )
    matches = _fast_text_layout_identifier_matches(layout, normalized)
    if len(matches) != 1:
        return _fast_text_verification(
            "review",
            match_method="spatial_invoice_identifier",
            match_count=len(matches),
            reason=("invoice_number_spatial_association_ambiguous" if matches else "invoice_number_spatial_context_missing"),
        )
    labels = _fast_text_layout_strong_invoice_label_rows(layout, matches[0])
    if len(labels) == 1:
        return _fast_text_verification(
            "confirmed",
            match_method="spatial_strong_invoice_label",
            match_count=1,
            context_type=labels[0]["label_type"],
        )
    return _fast_text_verification(
        "review",
        match_method="spatial_invoice_identifier",
        match_count=len(labels),
        reason=("invoice_number_spatial_association_ambiguous" if labels else "invoice_number_spatial_context_missing"),
    )


def _fast_text_layout_invoice_date_candidate(
    layout: Optional[Dict[str, Any]], *, spanish_fiscal_context: Optional[str]
) -> Dict[str, Any]:
    """Find one date under a strong invoice-date label in a local visual row."""
    candidates: Dict[str, bool] = {}
    for row in _fast_text_layout_rows(layout):
        row_text = _fast_text_layout_row_normalized_text(row)
        if not _FAST_TEXT_INVOICE_DATE_STRONG_PATTERN.search(row_text):
            continue
        for candidate_row in _fast_text_layout_local_rows(layout, row):
            candidate_text = _fast_text_layout_row_normalized_text(candidate_row)
            if _FAST_TEXT_DATE_EXCLUSION_PATTERN.search(candidate_text):
                continue
            for match in _FAST_TEXT_INVOICE_DATE_PATTERN.finditer(candidate_text):
                raw_date = match.group(0)
                ambiguous_numeric = (
                    len(match.group(3)) == 2
                    and int(match.group(1)) <= 12
                    and not spanish_fiscal_context
                )
                normalized = _normalize_fast_text_date(raw_date)
                if not normalized:
                    continue
                candidates[normalized] = candidates.get(normalized, False) or ambiguous_numeric
    if len(candidates) != 1:
        return {
            "status": "ambiguous" if candidates else "missing",
            "invoice_date": None,
            "candidate_count": len(candidates),
            "match_method": "spatial_invoice_date_label",
        }
    invoice_date, ambiguous_numeric = next(iter(candidates.items()))
    if ambiguous_numeric:
        return {
            "status": "ambiguous",
            "invoice_date": None,
            "candidate_count": 1,
            "match_method": "spatial_ambiguous_two_digit_invoice_date",
        }
    return {
        "status": "unambiguous",
        "invoice_date": invoice_date,
        "candidate_count": 1,
        "match_method": "spatial_strong_invoice_date_label",
    }


def _verify_fast_text_layout_invoice_date(
    invoice_date: Optional[str],
    date_diagnostics: Dict[str, Any],
    layout: Optional[Dict[str, Any]],
) -> Dict[str, Any]:
    evidence = _fast_text_layout_invoice_date_candidate(
        layout,
        spanish_fiscal_context=date_diagnostics.get("spanish_fiscal_context"),
    )
    if evidence["status"] == "unambiguous" and invoice_date == evidence["invoice_date"]:
        return _fast_text_verification(
            "confirmed",
            match_method=evidence["match_method"],
            match_count=1,
            context_type="invoice_date",
        )
    if evidence["status"] == "unambiguous" and invoice_date:
        return _fast_text_verification(
            "contradiction",
            match_method=evidence["match_method"],
            match_count=1,
            context_type="invoice_date",
            reason="invoice_date_differs_from_spatial_label",
        )
    return _fast_text_verification(
        "review",
        match_method=evidence["match_method"],
        match_count=evidence["candidate_count"],
        reason=("invoice_date_evidence_ambiguous" if evidence["status"] == "ambiguous" else "invoice_date_spatial_context_missing"),
    )


def _verify_fast_text_layout_supplier_tax_id(
    supplier_tax_id: Optional[str],
    layout: Optional[Dict[str, Any]],
    *,
    registered_company_tax_id: Optional[str],
) -> Dict[str, Any]:
    normalized = _normalize_fast_text_tax_id(supplier_tax_id)
    registered = _normalize_fast_text_tax_id(registered_company_tax_id)
    if not normalized:
        return _fast_text_verification("review", match_method="spatial_supplier_tax_id", reason="supplier_tax_id_missing")
    if registered and normalized == registered:
        return _fast_text_verification(
            "contradiction",
            match_method="registered_company_tax_id",
            context_type="recipient",
            reason="supplier_tax_id_matches_registered_company",
        )
    matches = _fast_text_layout_identifier_matches(layout, normalized)
    if len(matches) != 1:
        return _fast_text_verification(
            "review",
            match_method="spatial_supplier_tax_id",
            match_count=len(matches),
            reason=("supplier_tax_id_spatial_association_ambiguous" if matches else "supplier_tax_id_spatial_context_missing"),
        )
    local_rows = _fast_text_layout_local_rows(layout, matches[0]["row"])
    provider_rows = [
        row for row in local_rows if _FAST_TEXT_PROVIDER_CONTEXT_PATTERN.search(_fast_text_layout_row_normalized_text(row))
    ]
    recipient_rows = [
        row for row in local_rows if _FAST_TEXT_RECIPIENT_CONTEXT_PATTERN.search(_fast_text_layout_row_normalized_text(row))
    ]
    if recipient_rows and not provider_rows:
        return _fast_text_verification(
            "contradiction",
            match_method="spatial_party_region",
            match_count=1,
            context_type="recipient",
            reason="supplier_tax_id_in_recipient_context",
        )
    if len(provider_rows) == 1 and not recipient_rows:
        return _fast_text_verification(
            "confirmed",
            match_method="spatial_party_region",
            match_count=1,
            context_type="provider",
        )
    return _fast_text_verification(
        "review",
        match_method="spatial_party_region",
        match_count=len(provider_rows) + len(recipient_rows),
        reason="supplier_tax_id_spatial_region_unattributed",
    )


def _verify_fast_text_layout_document_type(layout: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    rows = _fast_text_layout_rows(layout)
    for row in rows[:20]:
        text = _fast_text_layout_row_normalized_text(row)
        if _FAST_TEXT_DOCUMENT_INVOICE_PATTERN.search(text) and not _FAST_TEXT_DOCUMENT_NON_INVOICE_PATTERN.search(text):
            return _fast_text_verification(
                "confirmed", match_method="spatial_invoice_header", match_count=1, context_type="invoice"
            )
    return _fast_text_verification(
        "review", match_method="spatial_invoice_header", reason="invoice_document_type_not_spatially_confirmed"
    )


def _verify_fast_text_document_type(
    document_text: str, invoice_number_verification: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    lines = _fast_text_lines_with_positions(document_text)
    invoice_lines = [
        line
        for line in lines[:20]
        if _FAST_TEXT_DOCUMENT_INVOICE_PATTERN.search(line)
        and not _FAST_TEXT_DOCUMENT_NON_INVOICE_PATTERN.search(line)
    ]
    if invoice_lines:
        return _fast_text_verification(
            "confirmed",
            match_method="invoice_header",
            match_count=len(invoice_lines),
            context_type="invoice",
        )
    contradictory_lines = [
        line
        for line in lines[:20]
        if _FAST_TEXT_DOCUMENT_INVOICE_PATTERN.search(line)
        and _FAST_TEXT_DOCUMENT_NON_INVOICE_PATTERN.search(line)
    ]
    if contradictory_lines:
        return _fast_text_verification(
            "contradiction",
            match_method="non_invoice_header",
            match_count=len(contradictory_lines),
            context_type="non_invoice",
            reason="document_header_is_not_an_unambiguous_invoice",
        )
    # A number already confirmed under a typed, strong invoice label is
    # sufficient proof of document type. Do not require a second, unrelated
    # standalone FACTURA heading that may have been displaced by PDF layout.
    if (invoice_number_verification or {}).get("status") == "confirmed":
        return _fast_text_verification(
            "confirmed",
            match_method="confirmed_invoice_identifier",
            match_count=max(int((invoice_number_verification or {}).get("match_count") or 1), 1),
            context_type="invoice",
        )
    return _fast_text_verification(
        "review",
        match_method="invoice_header",
        reason="invoice_document_type_not_confirmed",
    )


def _fast_text_tax_identifier_candidates(document_text: str) -> List[Dict[str, Any]]:
    lines = _fast_text_lines_with_positions(document_text)
    candidates: List[Dict[str, Any]] = []
    seen = set()
    for line_index, line in enumerate(lines):
        # Fiscal labels commonly precede the identifier. Do not include the
        # following party line, which could belong to the opposite side.
        context = "\n".join(lines[max(0, line_index - 1) : line_index + 1])
        for match in _FAST_TEXT_TAX_ID_TOKEN_PATTERN.finditer(line):
            normalized = _normalize_fast_text_tax_id(match.group(0))
            if not normalized or normalized in seen:
                continue
            seen.add(normalized)
            candidates.append(
                {
                    "value": normalized,
                    "line_index": line_index,
                    "provider_context": bool(_FAST_TEXT_PROVIDER_CONTEXT_PATTERN.search(context)),
                    "recipient_context": bool(_FAST_TEXT_RECIPIENT_CONTEXT_PATTERN.search(context)),
                    "tax_context": bool(_FAST_TEXT_TAX_ID_CONTEXT_PATTERN.search(context)),
                }
            )
    return candidates


def _is_fast_text_spanish_tax_id(value: Optional[str]) -> bool:
    """Recognize Spanish NIF/CIF/NIE and ES VAT identifiers structurally."""
    normalized = _normalize_fast_text_tax_id(value)
    return bool(normalized and _FAST_TEXT_SPANISH_TAX_ID_PATTERN.fullmatch(normalized))


def _fast_text_spanish_fiscal_context(
    document_text: str,
    *,
    supplier_tax_id: Optional[str],
    registered_company_tax_id: Optional[str],
) -> Optional[str]:
    """Return Spain's date-format proof only for a confirmed supplier ID.

    A recipient's Spanish tax identifier is common on foreign invoices and
    cannot establish the issuer's document convention.  The supplier tax ID
    must therefore be both structurally Spanish and confirmed by the existing
    party-aware verifier before an ambiguous numeric date uses DD/MM/YY.
    """
    normalized_supplier_tax_id = _normalize_fast_text_tax_id(supplier_tax_id)
    if not _is_fast_text_spanish_tax_id(normalized_supplier_tax_id):
        return None

    supplier_verification = _verify_fast_text_supplier_tax_id(
        normalized_supplier_tax_id,
        document_text,
        registered_company_tax_id=registered_company_tax_id,
    )
    if supplier_verification.get("status") != "confirmed":
        return None

    if normalized_supplier_tax_id and normalized_supplier_tax_id.startswith("ES"):
        return "confirmed_spanish_supplier_vat"
    return "confirmed_spanish_supplier_tax_id"


def _verify_fast_text_supplier_tax_id(
    supplier_tax_id: Optional[str],
    document_text: str,
    *,
    registered_company_tax_id: Optional[str],
) -> Dict[str, Any]:
    normalized = _normalize_fast_text_tax_id(supplier_tax_id)
    registered = _normalize_fast_text_tax_id(registered_company_tax_id)
    if not normalized:
        return _fast_text_verification(
            "review", match_method="missing", reason="supplier_tax_id_missing"
        )
    if registered and normalized == registered:
        return _fast_text_verification(
            "contradiction",
            match_method="registered_company_tax_id",
            context_type="recipient",
            reason="supplier_tax_id_matches_registered_company",
        )
    candidates = _fast_text_tax_identifier_candidates(document_text)
    matches = [candidate for candidate in candidates if candidate["value"] == normalized]
    if not matches:
        return _fast_text_verification(
            "review", match_method="full_identifier", reason="supplier_tax_id_not_found"
        )
    if any(candidate["recipient_context"] for candidate in matches):
        return _fast_text_verification(
            "contradiction",
            match_method="full_identifier",
            match_count=len(matches),
            context_type="recipient",
            reason="supplier_tax_id_in_recipient_context",
        )
    if any(candidate["provider_context"] for candidate in matches):
        return _fast_text_verification(
            "confirmed",
            match_method="full_identifier",
            match_count=len(matches),
            context_type="provider",
        )
    candidate_values = {candidate["value"] for candidate in candidates}
    if candidate_values == {normalized} or (
        registered and candidate_values.issubset({normalized, registered})
    ):
        return _fast_text_verification(
            "confirmed",
            match_method="unique_document_tax_id",
            match_count=len(matches),
            context_type="tax_id",
        )
    return _fast_text_verification(
        "review",
        match_method="full_identifier",
        match_count=len(matches),
        context_type="tax_id",
        reason="multiple_unattributed_tax_ids",
    )


def _normalize_fast_text_entity_for_verification(value: Any) -> str:
    normalized = unicodedata.normalize("NFKD", str(value or ""))
    normalized = "".join(char for char in normalized if not unicodedata.combining(char))
    tokens = re.findall(r"[A-Za-z0-9]+", normalized.upper())
    legal_suffixes = {"SL", "SLU", "SA", "SAL", "SCP", "LLC", "LTD", "INC"}
    return "".join(token for token in tokens if token not in legal_suffixes)


def _verify_fast_text_provider_name(
    provider_name: Optional[str], document_text: str, supplier_tax_verification: Dict[str, Any]
) -> Dict[str, Any]:
    normalized = _normalize_fast_text_entity_for_verification(provider_name)
    if not normalized:
        return _fast_text_verification(
            "review", match_method="missing", reason="provider_name_missing"
        )
    lines = _fast_text_lines_with_positions(document_text)
    matches = [
        line_index
        for line_index, line in enumerate(lines)
        if normalized in _normalize_fast_text_entity_for_verification(line)
    ]
    if not matches:
        return _fast_text_verification(
            "review", match_method="normalized_name", reason="provider_name_not_found"
        )
    def party_context(line_index: int) -> str:
        # A party label may be on the preceding line, but looking ahead can
        # associate the supplier label with the next recipient name. An
        # explicit label on the name's own line takes precedence over a
        # preceding, unrelated party label.
        current_line = lines[line_index]
        if (
            _FAST_TEXT_PROVIDER_CONTEXT_PATTERN.search(current_line)
            or _FAST_TEXT_RECIPIENT_CONTEXT_PATTERN.search(current_line)
        ):
            return current_line
        return "\n".join(lines[max(0, line_index - 1) : line_index + 1])

    recipient_matches = [
        line_index
        for line_index in matches
        if _FAST_TEXT_RECIPIENT_CONTEXT_PATTERN.search(party_context(line_index))
    ]
    provider_matches = [
        line_index
        for line_index in matches
        if _FAST_TEXT_PROVIDER_CONTEXT_PATTERN.search(party_context(line_index))
    ]
    # A name found only on the recipient side is a concrete party-role
    # contradiction. A verified supplier tax ID cannot make that contradictory
    # provider proposal safe to accept.
    if recipient_matches and not provider_matches:
        return _fast_text_verification(
            "contradiction",
            match_method="normalized_name",
            match_count=len(matches),
            context_type="recipient",
            reason="provider_name_in_recipient_context",
        )
    return _fast_text_verification(
        "confirmed",
        match_method="normalized_name",
        match_count=len(matches),
        context_type=(
            "provider"
            if provider_matches
            else "document"
        ),
    )


def _parse_fast_text_money(value: Any) -> Optional[Decimal]:
    raw = str(value or "").upper().replace("€", "").replace("£", "")
    raw = raw.replace("EUR", "").replace("USD", "").replace("GBP", "").replace("US$", "")
    raw = re.sub(r"\s+", "", raw)
    if not raw or not re.fullmatch(r"[-+]?\d[\d.,]*", raw):
        return None
    sign = -1 if raw.startswith("-") else 1
    raw = raw.lstrip("+-")
    if "," in raw and "." in raw:
        decimal_separator = "," if raw.rfind(",") > raw.rfind(".") else "."
        thousands_separator = "." if decimal_separator == "," else ","
        normalized = raw.replace(thousands_separator, "").replace(decimal_separator, ".")
    elif "," in raw or "." in raw:
        separator = "," if "," in raw else "."
        left, right = raw.rsplit(separator, 1)
        if len(right) == 2:
            normalized = left.replace(separator, "") + "." + right
        elif len(right) == 3:
            normalized = raw.replace(separator, "")
        else:
            return None
    else:
        normalized = raw
    try:
        return (Decimal(sign) * Decimal(normalized)).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
    except (InvalidOperation, ValueError):
        return None


def _fast_text_money_values_in_line(line: str) -> List[Decimal]:
    values = []
    for match in _FAST_TEXT_MONEY_TOKEN_PATTERN.finditer(line):
        if match.end() < len(line) and line[match.end()] == "%":
            continue
        value = _parse_fast_text_money(match.group(0))
        if value is not None:
            values.append(value)
    return values


_FAST_TEXT_TABLE_VALUE_HEADER_PATTERN = re.compile(
    r"\b(?:IMPORTE|AMOUNT|CUOTA|VALOR|EUROS?|CANTIDAD)\b"
)


def _fast_text_is_label_only_money_table_line(field: str, line: str) -> bool:
    """Allow a short header bridge, never arbitrary intervening prose."""
    if _fast_text_money_values_in_line(line):
        return False
    remainder = line
    remainder = _FAST_TEXT_MONEY_CONTEXT_PATTERNS[field].sub(" ", remainder)
    remainder = _FAST_TEXT_TABLE_VALUE_HEADER_PATTERN.sub(" ", remainder)
    return bool(line.strip()) and not re.search(r"[A-Z0-9]", remainder)


def _fast_text_labeled_money_context_values(
    field: str, lines: List[str], line_index: int
) -> List[Decimal]:
    """Read one bounded label-to-value relationship from native text.

    PDF extraction can place a table's amount header between a semantic label
    and its value. We may bridge at most two label-only lines; prose, another
    row or any unlabelled material ends the relationship.
    """
    pattern = _FAST_TEXT_MONEY_CONTEXT_PATTERNS[field]
    line = lines[line_index]
    values = _fast_text_money_values_in_line(line)
    if pattern.search(line):
        if values:
            return values
        for offset in (1, 2):
            next_index = line_index + offset
            if next_index >= len(lines):
                break
            candidate_line = lines[next_index]
            candidate_values = _fast_text_money_values_in_line(candidate_line)
            if candidate_values:
                return candidate_values
            if not _fast_text_is_label_only_money_table_line(field, candidate_line):
                break
        return []

    if line_index:
        previous = lines[line_index - 1]
        if pattern.search(previous) and not _fast_text_money_values_in_line(previous):
            return values
    return []


def _verify_fast_text_optional_adjustment_field(
    field: str,
    proposed_value: Any,
    document_text: str,
    *,
    accounting_equation_confirmed: bool,
) -> Dict[str, Any]:
    """Distinguish an absent adjustment from a missing required amount.

    A null/zero retention or surcharge is only non-applicable after checking
    that the native text has no labelled material amount and the invoice
    equation independently closes without that adjustment.
    """
    proposed = _money_decimal(proposed_value)
    if proposed is not None and abs(proposed) > Decimal("0.01"):
        return _verify_fast_text_monetary_field(field, proposed_value, document_text)

    lines = _fast_text_lines_with_positions(document_text)
    pattern = _FAST_TEXT_MONEY_CONTEXT_PATTERNS[field]
    label_seen = False
    labelled_values: List[Decimal] = []
    for line_index, line in enumerate(lines):
        if not pattern.search(line):
            continue
        label_seen = True
        labelled_values.extend(
            _fast_text_labeled_money_context_values(field, lines, line_index)
        )

    material_values = [
        value for value in labelled_values if abs(value) > Decimal("0.01")
    ]
    if material_values:
        return _fast_text_verification(
            "contradiction",
            match_method="labelled_adjustment",
            match_count=len(material_values),
            context_type=field,
            reason=f"{field}_zero_differs_from_document_context",
        )
    if label_seen and not labelled_values:
        return _fast_text_verification(
            "review",
            match_method="labelled_adjustment",
            context_type=field,
            reason=f"{field}_label_without_amount",
        )
    if label_seen:
        return _fast_text_verification(
            "not_applicable",
            match_method="document_zero_adjustment",
            match_count=len(labelled_values),
            context_type=field,
        )
    if accounting_equation_confirmed:
        return _fast_text_verification(
            "not_applicable",
            match_method="no_document_adjustment",
            context_type=field,
        )
    return _fast_text_verification(
        "review",
        match_method="missing_adjustment_context",
        context_type=field,
        reason=f"{field}_not_applicable_unproven",
    )


def _verify_fast_text_monetary_field(
    field: str, proposed_value: Any, document_text: str
) -> Dict[str, Any]:
    proposed = _money_decimal(proposed_value)
    if proposed is None:
        return _fast_text_verification(
            "review", match_method="missing", reason=f"{field}_missing"
        )
    pattern = _FAST_TEXT_MONEY_CONTEXT_PATTERNS[field]
    lines = _fast_text_lines_with_positions(document_text)
    document_matches = 0
    semantic_matches = 0
    contextual_values = set()
    for line_index, line in enumerate(lines):
        values = _fast_text_money_values_in_line(line)
        document_matches += sum(abs(value - proposed) <= Decimal("0.01") for value in values)
        contextual_line_values = _fast_text_labeled_money_context_values(
            field, lines, line_index
        )
        if not contextual_line_values:
            continue
        contextual_values.update(contextual_line_values)
        if any(abs(value - proposed) <= Decimal("0.01") for value in contextual_line_values):
            semantic_matches += 1
    if semantic_matches:
        return _fast_text_verification(
            "confirmed",
            match_method="contextual_monetary_value",
            match_count=semantic_matches,
            context_type=field,
        )
    if contextual_values and len(contextual_values) == 1:
        return _fast_text_verification(
            "contradiction",
            match_method="contextual_monetary_value",
            match_count=document_matches,
            context_type=field,
            reason=f"{field}_differs_from_document_context",
        )
    return _fast_text_verification(
        "review",
        match_method="document_value" if document_matches else "not_found",
        match_count=document_matches,
        context_type=field if document_matches else None,
        reason=(f"{field}_missing_context" if document_matches else f"{field}_not_found"),
    )


_FAST_TEXT_VAT_TABLE_BASE_PATTERN = re.compile(r"\b(?:BASE\s+IMPONIBLE|BASE|TAXABLE\s+BASE)\b")
_FAST_TEXT_VAT_TABLE_AMOUNT_PATTERN = re.compile(
    r"\b(?:IVA|VAT|CUOTA(?:\s+TRIBUTARIA)?|IMPUESTO)\b"
)


def _fast_text_has_vat_table_header(value: str) -> bool:
    """Recognize a fiscal table header, not a generic mention of VAT."""
    return bool(
        _FAST_TEXT_VAT_TABLE_BASE_PATTERN.search(value)
        and _FAST_TEXT_VAT_TABLE_AMOUNT_PATTERN.search(value)
    )


def _fast_text_single_tax_line_in_bounded_table(
    lines: List[str], rate: Any, base: Decimal, vat: Decimal
) -> bool:
    """Confirm one VAT row when column extraction split its cells by lines.

    This intentionally handles only a single material tax line with one rate
    in a small header-plus-row block. Multiple-rate tables remain fail-closed
    unless the existing local-row verifier can prove each line independently.
    """
    rate_pattern = re.compile(
        rf"(?<![0-9]){re.escape(format(float(rate), 'g'))}(?:[.,]0+)?\s*%(?![0-9])"
    )
    for start in range(len(lines)):
        header = "\n".join(lines[start : min(len(lines), start + 2)])
        if not _fast_text_has_vat_table_header(header):
            continue
        block = "\n".join(lines[start : min(len(lines), start + 5)])
        if len(rate_pattern.findall(block)) != 1:
            continue
        values = _fast_text_money_values_in_line(block)
        if (
            any(abs(value - base) <= Decimal("0.01") for value in values)
            and any(abs(value - vat) <= Decimal("0.01") for value in values)
        ):
            return True
    return False


def _verify_fast_text_vat_breakdown(
    taxes: List[Dict[str, Any]], document_text: str
) -> Dict[str, Any]:
    material_lines = [
        line
        for line in taxes or []
        if abs(_money_decimal(line.get("base")) or Decimal("0")) > Decimal("0.01")
        or abs(_money_decimal(line.get("vat_amount")) or Decimal("0")) > Decimal("0.01")
    ]
    if not material_lines:
        return _fast_text_verification("not_applicable", match_method="zero_tax_lines")
    lines = _fast_text_lines_with_positions(document_text)
    confirmed = 0
    for tax_line in material_lines:
        rate = tax_line.get("rate")
        base = _money_decimal(tax_line.get("base"))
        vat = _money_decimal(tax_line.get("vat_amount"))
        if rate is None or base is None or vat is None:
            return _fast_text_verification(
                "review", match_method="incomplete_tax_line", reason="vat_breakdown_incomplete"
            )
        rate_pattern = re.compile(
            rf"(?<![0-9]){re.escape(format(float(rate), 'g'))}(?:[.,]0+)?\s*%(?![0-9])"
        )
        line_confirmed = False
        for line_index, line in enumerate(lines):
            context = _fast_text_local_context(lines, line_index)
            values = _fast_text_money_values_in_line(context)
            if (
                rate_pattern.search(context)
                and any(abs(value - base) <= Decimal("0.01") for value in values)
                and any(abs(value - vat) <= Decimal("0.01") for value in values)
            ):
                line_confirmed = True
                break
        if (
            not line_confirmed
            and len(material_lines) == 1
            and _fast_text_single_tax_line_in_bounded_table(lines, rate, base, vat)
        ):
            line_confirmed = True
        if not line_confirmed:
            return _fast_text_verification(
                "review",
                match_method="tax_row",
                match_count=confirmed,
                context_type="vat_breakdown",
                reason="vat_breakdown_line_not_documented",
            )
        confirmed += 1
    return _fast_text_verification(
        "confirmed", match_method="tax_row", match_count=confirmed, context_type="vat_breakdown"
    )


def _fast_text_layout_rate_matches(row: Dict[str, Any], rate: Any) -> List[Dict[str, Any]]:
    try:
        normalized_rate = format(float(rate), "g")
    except (TypeError, ValueError):
        return []
    pattern = re.compile(rf"(?<![0-9]){re.escape(normalized_rate)}(?:[.,]0+)?\s*%(?![0-9])")
    return [
        token
        for token in row.get("tokens") or []
        if pattern.search(_normalize_fast_text_parser_line(token.get("text")))
    ]


def _fast_text_layout_table_rows_after(
    layout: Optional[Dict[str, Any]], header: Dict[str, Any], count: int
) -> List[Dict[str, Any]]:
    rows = _fast_text_layout_page_rows(layout, header.get("page"))
    result = []
    previous = header
    for candidate in rows[header.get("index", 0) + 1 :]:
        if not _fast_text_layout_rows_are_related(previous, candidate):
            break
        result.append(candidate)
        previous = candidate
        if len(result) >= max(1, count):
            break
    return result


def _fast_text_layout_vat_row_matches(
    layout: Optional[Dict[str, Any]], rate: Any, base: Decimal, vat: Decimal, tax_line_count: int
) -> List[Tuple[int, int]]:
    matches = []
    for header in _fast_text_layout_rows(layout):
        base_boxes = _fast_text_layout_label_boxes(header, _FAST_TEXT_VAT_TABLE_BASE_PATTERN)
        vat_boxes = _fast_text_layout_label_boxes(header, _FAST_TEXT_VAT_TABLE_AMOUNT_PATTERN)
        rate_boxes = _fast_text_layout_label_boxes(header, _FAST_TEXT_LAYOUT_RATE_HEADER_PATTERN)
        if len(base_boxes) != 1 or len(vat_boxes) != 1 or len(rate_boxes) != 1:
            continue
        for row in _fast_text_layout_table_rows_after(layout, header, tax_line_count):
            rate_matches = _fast_text_layout_rate_matches(row, rate)
            base_matches = _fast_text_layout_money_matches(row, base)
            vat_matches = _fast_text_layout_money_matches(row, vat)
            aligned_base = [
                token
                for token in base_matches
                if _fast_text_layout_box_centers_align(
                    base_boxes[0], (token["x0"], token["y0"], token["x1"], token["y1"])
                )
            ]
            aligned_vat = [
                token
                for token in vat_matches
                if _fast_text_layout_box_centers_align(
                    vat_boxes[0], (token["x0"], token["y0"], token["x1"], token["y1"])
                )
            ]
            aligned_rate = [
                token
                for token in rate_matches
                if _fast_text_layout_box_centers_align(
                    rate_boxes[0], (token["x0"], token["y0"], token["x1"], token["y1"])
                )
            ]
            if len(aligned_rate) == len(aligned_base) == len(aligned_vat) == 1:
                matches.append((header["page"], row["index"]))
    return matches


def _verify_fast_text_layout_vat_breakdown(
    taxes: List[Dict[str, Any]], layout: Optional[Dict[str, Any]]
) -> Dict[str, Any]:
    material_lines = [
        line
        for line in taxes or []
        if abs(_money_decimal(line.get("base")) or Decimal("0")) > Decimal("0.01")
        or abs(_money_decimal(line.get("vat_amount")) or Decimal("0")) > Decimal("0.01")
    ]
    if not material_lines:
        return _fast_text_verification("not_applicable", match_method="zero_tax_lines")
    confirmed_rows = set()
    for tax_line in material_lines:
        rate = tax_line.get("rate")
        base = _money_decimal(tax_line.get("base"))
        vat = _money_decimal(tax_line.get("vat_amount"))
        if rate is None or base is None or vat is None:
            return _fast_text_verification(
                "review", match_method="spatial_tax_row", reason="vat_breakdown_incomplete"
            )
        expected_vat = (base * Decimal(str(rate)) / Decimal("100")).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        if abs(expected_vat - vat) > Decimal("0.01"):
            return _fast_text_verification(
                "contradiction", match_method="spatial_tax_row", reason="vat_breakdown_math_mismatch"
            )
        matches = _fast_text_layout_vat_row_matches(
            layout, rate, base, vat, len(material_lines)
        )
        if len(matches) != 1:
            return _fast_text_verification(
                "review",
                match_method="spatial_tax_row",
                match_count=len(confirmed_rows),
                context_type="vat_breakdown",
                reason=("vat_breakdown_spatial_association_ambiguous" if matches else "vat_breakdown_spatial_context_missing"),
            )
        confirmed_rows.add(matches[0])
    if len(confirmed_rows) != len(material_lines):
        return _fast_text_verification(
            "review", match_method="spatial_tax_row", reason="vat_breakdown_spatial_rows_not_unique"
        )
    return _fast_text_verification(
        "confirmed",
        match_method="spatial_tax_row",
        match_count=len(confirmed_rows),
        context_type="vat_breakdown",
    )


def _verify_fast_text_layout_optional_adjustment_field(
    field: str, proposed_value: Any, layout: Optional[Dict[str, Any]]
) -> Dict[str, Any]:
    proposed = _money_decimal(proposed_value)
    if proposed is not None and abs(proposed) > Decimal("0.01"):
        return _verify_fast_text_layout_monetary_field(field, proposed_value, layout)
    material_associations = 0
    pattern = _FAST_TEXT_MONEY_CONTEXT_PATTERNS[field]
    for row in _fast_text_layout_rows(layout):
        if not _fast_text_layout_label_boxes(row, pattern):
            continue
        candidates = _fast_text_layout_money_values(row)
        for related in _fast_text_layout_local_rows(layout, row):
            if related is not row:
                candidates.extend(_fast_text_layout_money_values(related))
        material_values = [value for _token, value in candidates if abs(value) > Decimal("0.01")]
        if len(material_values) == 1:
            material_associations += 1
    if material_associations == 1:
        return _fast_text_verification(
            "contradiction",
            match_method="spatial_labelled_adjustment",
            match_count=1,
            context_type=field,
            reason=f"{field}_zero_differs_from_spatial_context",
        )
    return _fast_text_verification(
        "review",
        match_method="spatial_labelled_adjustment",
        match_count=material_associations,
        context_type=field,
        reason=(f"{field}_spatial_association_ambiguous" if material_associations > 1 else f"{field}_spatial_context_missing"),
    )


def _verify_fast_text_currency(currency: Optional[str], document_text: str) -> Dict[str, Any]:
    normalized = str(currency or "").upper().strip()
    if normalized not in _FAST_TEXT_CURRENCY_MARKERS:
        return _fast_text_verification(
            "review", match_method="currency_marker", reason="currency_missing_or_unsupported"
        )
    detected = {
        code
        for code, pattern in _FAST_TEXT_CURRENCY_MARKERS.items()
        if pattern.search(document_text or "")
    }
    if normalized in detected and detected == {normalized}:
        return _fast_text_verification(
            "confirmed", match_method="currency_marker", match_count=1, context_type=normalized
        )
    if normalized not in detected:
        return _fast_text_verification(
            "contradiction",
            match_method="currency_marker",
            match_count=len(detected),
            reason="currency_differs_from_document",
        )
    return _fast_text_verification(
        "review",
        match_method="currency_marker",
        match_count=len(detected),
        reason="multiple_document_currencies",
    )


def _verify_fast_text_invoice_date(
    invoice_date: Optional[str], date_diagnostics: Dict[str, Any]
) -> Dict[str, Any]:
    status = date_diagnostics.get("status")
    if (
        status == "unambiguous"
        and invoice_date
        and date_diagnostics.get("model_matches_evidence")
    ):
        return _fast_text_verification(
            "confirmed",
            match_method=str(date_diagnostics.get("match_method") or "invoice_date_label"),
            match_count=1,
            context_type="invoice_date",
        )
    if status == "unambiguous" and invoice_date:
        return _fast_text_verification(
            "contradiction",
            match_method=str(date_diagnostics.get("match_method") or "invoice_date_label"),
            match_count=1,
            context_type="invoice_date",
            reason="invoice_date_differs_from_document_context",
        )
    if status == "ambiguous":
        return _fast_text_verification(
            "review", match_method="invoice_date_label", reason="invoice_date_evidence_ambiguous"
        )
    return _fast_text_verification(
        "review", match_method="invoice_date_label", reason="invoice_date_not_documented"
    )


def _verify_fast_text_invoice_number(
    invoice_number_evidence_status: str, diagnostics: Dict[str, Any]
) -> Dict[str, Any]:
    if invoice_number_evidence_status == "confirmed":
        return _fast_text_verification(
            "confirmed",
            match_method=str(diagnostics.get("model_value_match_method") or "label_first"),
            match_count=diagnostics.get("model_value_match_count") or 1,
            context_type=diagnostics.get("model_value_context_label_type"),
        )
    if invoice_number_evidence_status in {"ambiguous", "conflict"}:
        return _fast_text_verification(
            "contradiction",
            match_method="invoice_identifier",
            match_count=diagnostics.get("model_value_match_count") or 0,
            reason=diagnostics.get("conflict_reason") or diagnostics.get("ambiguity_reason"),
        )
    return _fast_text_verification(
        "review",
        match_method="invoice_identifier",
        match_count=diagnostics.get("model_value_match_count") or 0,
        reason=diagnostics.get("conflict_reason") or "invoice_number_not_confirmed",
    )


def _verify_fast_text_payment_dates(
    payment_dates: List[str],
    invoice_date: Optional[str],
    document_text: str,
    deterministic_corrections: List[str],
) -> Dict[str, Any]:
    if not payment_dates:
        return _fast_text_verification("not_applicable", match_method="no_payment_dates")
    if "due_date_derived_from_payment_terms" in deterministic_corrections:
        return _fast_text_verification(
            "confirmed", match_method="deterministic_payment_terms", match_count=len(payment_dates)
        )
    explicit_dates = _find_due_dates_in_due_context(document_text)
    normalized_payment_dates = []
    for value in payment_dates:
        normalized = _normalize_fast_text_date(value)
        if normalized and normalized not in normalized_payment_dates:
            normalized_payment_dates.append(normalized)
    if explicit_dates == normalized_payment_dates:
        return _fast_text_verification(
            "confirmed", match_method="explicit_due_date", match_count=len(payment_dates)
        )
    if explicit_dates:
        return _fast_text_verification(
            "contradiction", match_method="explicit_due_date", reason="payment_dates_differ_from_document"
        )
    return _fast_text_verification(
        "review", match_method="due_date_context", reason="payment_dates_not_documented"
    )


def _verify_fast_text_accounting_equation(normalized: Dict[str, Any]) -> Dict[str, Any]:
    base = _money_decimal(normalized.get("base_amount"))
    vat = _money_decimal(normalized.get("vat_amount"))
    withholding = _money_decimal(normalized.get("withholding_amount")) or Decimal("0.00")
    other_taxes = _money_decimal(normalized.get("other_taxes")) or Decimal("0.00")
    total = _money_decimal(normalized.get("total_amount"))
    if base is None or vat is None or total is None:
        return _fast_text_verification(
            "review", match_method="accounting_equation", reason="accounting_values_missing"
        )
    expected = (base + vat + other_taxes - abs(withholding)).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )
    if abs(expected - total) > Decimal("0.01"):
        return _fast_text_verification(
            "contradiction", match_method="accounting_equation", reason="accounting_equation_mismatch"
        )
    return _fast_text_verification("confirmed", match_method="accounting_equation")


def _verify_fast_text_document(
    normalized: Dict[str, Any],
    document_text: str,
    *,
    document_text_complete: bool,
    registered_company_tax_id: Optional[str],
    invoice_number_evidence_status: str,
    invoice_parser_diagnostics: Dict[str, Any],
    invoice_date_diagnostics: Dict[str, Any],
    deterministic_corrections: List[str],
    document_layout: Optional[Dict[str, Any]] = None,
) -> Dict[str, Dict[str, Any]]:
    """Verify V2 proposals against compact native text without extracting anew."""
    supplier_tax_id = _fast_text_layout_merge_verification(
        _verify_fast_text_supplier_tax_id(
            normalized.get("supplier_tax_id"),
            document_text,
            registered_company_tax_id=registered_company_tax_id,
        ),
        _verify_fast_text_layout_supplier_tax_id(
            normalized.get("supplier_tax_id"),
            document_layout,
            registered_company_tax_id=registered_company_tax_id,
        ),
    )
    invoice_number = _verify_fast_text_invoice_number(
        invoice_number_evidence_status, invoice_parser_diagnostics
    )
    accounting_equation = _verify_fast_text_accounting_equation(normalized)
    verification = {
        "document_type": _fast_text_layout_merge_verification(
            _verify_fast_text_document_type(document_text, invoice_number),
            _verify_fast_text_layout_document_type(document_layout),
        ),
        "supplier_tax_id": supplier_tax_id,
        "provider_name": _verify_fast_text_provider_name(
            normalized.get("provider_name"), document_text, supplier_tax_id
        ),
        "invoice_number": invoice_number,
        "invoice_date": _fast_text_layout_merge_verification(
            _verify_fast_text_invoice_date(
                normalized.get("invoice_date"), invoice_date_diagnostics
            ),
            _verify_fast_text_layout_invoice_date(
                normalized.get("invoice_date"), invoice_date_diagnostics, document_layout
            ),
        ),
        "base_amount": _fast_text_layout_merge_verification(
            _verify_fast_text_monetary_field(
                "base_amount", normalized.get("base_amount"), document_text
            ),
            _verify_fast_text_layout_monetary_field(
                "base_amount", normalized.get("base_amount"), document_layout
            ),
        ),
        "vat_amount": _fast_text_layout_merge_verification(
            _verify_fast_text_monetary_field(
                "vat_amount", normalized.get("vat_amount"), document_text
            ),
            _verify_fast_text_layout_monetary_field(
                "vat_amount", normalized.get("vat_amount"), document_layout
            ),
        ),
        "vat_breakdown": _fast_text_layout_merge_verification(
            _verify_fast_text_vat_breakdown(normalized.get("vat_breakdown") or [], document_text),
            _verify_fast_text_layout_vat_breakdown(
                normalized.get("vat_breakdown") or [], document_layout
            ),
        ),
        "withholding_amount": _fast_text_layout_merge_verification(
            _verify_fast_text_optional_adjustment_field(
                "withholding_amount",
                normalized.get("withholding_amount"),
                document_text,
                accounting_equation_confirmed=accounting_equation["status"] == "confirmed",
            ),
            _verify_fast_text_layout_optional_adjustment_field(
                "withholding_amount", normalized.get("withholding_amount"), document_layout
            ),
            allow_not_applicable_refutation=True,
        ),
        "other_taxes_amount": _fast_text_layout_merge_verification(
            _verify_fast_text_optional_adjustment_field(
                "other_taxes_amount",
                normalized.get("other_taxes"),
                document_text,
                accounting_equation_confirmed=accounting_equation["status"] == "confirmed",
            ),
            _verify_fast_text_layout_optional_adjustment_field(
                "other_taxes_amount", normalized.get("other_taxes"), document_layout
            ),
            allow_not_applicable_refutation=True,
        ),
        "total_amount": _fast_text_layout_merge_verification(
            _verify_fast_text_monetary_field(
                "total_amount", normalized.get("total_amount"), document_text
            ),
            _verify_fast_text_layout_monetary_field(
                "total_amount", normalized.get("total_amount"), document_layout
            ),
        ),
        "currency": _verify_fast_text_currency(normalized.get("currency"), document_text),
        "payment_dates": _verify_fast_text_payment_dates(
            normalized.get("payment_dates") or [],
            normalized.get("invoice_date"),
            document_text,
            deterministic_corrections,
        ),
        "accounting_equation": accounting_equation,
    }
    verification["document_text"] = _fast_text_verification(
        "confirmed" if document_text_complete else "review",
        match_method="complete_native_text" if document_text_complete else "truncated_native_text",
        reason=None if document_text_complete else "document_text_truncated",
    )
    return verification


def _decide_fast_text_fast_path(
    verification: Dict[str, Dict[str, Any]],
) -> Tuple[str, List[str]]:
    """Return a shadow-only, fail-closed V2 decision independent of V1."""
    reasons: List[str] = []
    for field in ("document_text",) + _FAST_TEXT_FAST_PATH_CRITICAL_FIELDS + ("accounting_equation",):
        result = verification.get(field) or {}
        status = result.get("status")
        if status in {"confirmed", "not_applicable"}:
            continue
        reasons.append(f"{field}:{result.get('reason') or status or 'not_confirmed'}")
    provider_status = (verification.get("provider_name") or {}).get("status")
    if provider_status == "contradiction":
        reasons.append("provider_name:provider_identity_contradiction")
    return ("accept_v2", []) if not reasons else ("fallback_v1", list(dict.fromkeys(reasons)))


def _normalize_fast_text_invoice(structured_data: Dict[str, Any]) -> Dict[str, Any]:
    supplier = structured_data.get("supplier") or {}
    customer = structured_data.get("customer") or {}
    invoice = structured_data.get("invoice") or {}
    totals = structured_data.get("totals") or {}
    due_dates = []
    for value in structured_data.get("due_dates") or []:
        normalized = _normalize_date(value)
        if normalized and normalized not in due_dates:
            due_dates.append(normalized)
    taxes = _normalize_fast_text_tax_lines(structured_data.get("taxes"))
    provider_name = _strip_inline_tax_id(supplier.get("legal_name")) if supplier.get("legal_name") else None
    customer_name = _strip_inline_tax_id(customer.get("legal_name")) if customer.get("legal_name") else None
    base_amount = _money_decimal(totals.get("taxable_base"))
    vat_amount = _money_decimal(totals.get("vat_amount"))
    withholding = _money_decimal(totals.get("withholding"))
    other_taxes = _money_decimal(totals.get("other_taxes"))
    total_amount = _money_decimal(totals.get("total"))
    return {
        "provider_name": provider_name.strip() if isinstance(provider_name, str) else None,
        "supplier_tax_id": _normalize_fast_text_tax_id(supplier.get("tax_id")),
        "client_name": customer_name.strip() if isinstance(customer_name, str) else None,
        "customer_tax_id": _normalize_fast_text_tax_id(customer.get("tax_id")),
        "invoice_number": str(invoice.get("invoice_number") or "").strip() or None,
        "invoice_date": _normalize_fast_text_date(invoice.get("issue_date")),
        "payment_dates": due_dates,
        "currency": str(invoice.get("currency") or "").strip().upper() or None,
        "base_amount": _round_amount(float(base_amount)) if base_amount is not None else None,
        "vat_amount": _round_amount(float(vat_amount)) if vat_amount is not None else None,
        "withholding_amount": _round_amount(float(abs(withholding))) if withholding is not None else None,
        "other_taxes": _round_amount(float(other_taxes)) if other_taxes is not None else None,
        "total_amount": _round_amount(float(total_amount)) if total_amount is not None else None,
        "vat_breakdown": taxes,
    }


def _validate_fast_text_invoice(
    structured_data: Dict[str, Any],
    normalized: Dict[str, Any],
    company_names: Optional[List[str]],
    *,
    registered_company_tax_id: Optional[str] = None,
) -> List[str]:
    """Reject questionable V2 output without correcting or enriching it."""
    issues: List[str] = []
    provider_name = normalized.get("provider_name")
    client_name = normalized.get("client_name")
    invoice_number = normalized.get("invoice_number")
    invoice_date = normalized.get("invoice_date")
    supplier_tax_id = normalized.get("supplier_tax_id")
    base_amount = _money_decimal(normalized.get("base_amount"))
    vat_amount = _money_decimal(normalized.get("vat_amount"))
    withholding = _money_decimal(normalized.get("withholding_amount")) or Decimal("0.00")
    other_taxes = _money_decimal(normalized.get("other_taxes")) or Decimal("0.00")
    total_amount = _money_decimal(normalized.get("total_amount"))
    taxes = normalized.get("vat_breakdown") or []

    if not provider_name:
        issues.append("missing_supplier")
    elif _is_same_entity(provider_name, company_names):
        issues.append("supplier_matches_registered_customer")
    registered_tax_id = _normalize_fast_text_tax_id(registered_company_tax_id)
    if supplier_tax_id and registered_tax_id and supplier_tax_id == registered_tax_id:
        issues.append("supplier_tax_id_matches_registered_customer")
    if provider_name and client_name and _normalize_entity_name(provider_name) == _normalize_entity_name(client_name):
        issues.append("supplier_matches_customer")
    if not invoice_number:
        issues.append("missing_invoice_number")
    if not invoice_date:
        issues.append("invalid_invoice_date")
    if not _is_plausible_fast_text_tax_id(supplier_tax_id):
        issues.append("invalid_supplier_tax_id")
    if base_amount is None or vat_amount is None or total_amount is None:
        issues.append("missing_accounting_totals")
    else:
        expected_total = (base_amount + vat_amount + other_taxes - abs(withholding)).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        if abs(expected_total - total_amount) > Decimal("0.01"):
            issues.append("inconsistent_accounting_equation")
    if not taxes:
        issues.append("missing_tax_lines")
    else:
        line_bases = Decimal("0.00")
        line_vat = Decimal("0.00")
        for line in taxes:
            rate = line.get("rate")
            line_base = _money_decimal(line.get("base"))
            line_amount = _money_decimal(line.get("vat_amount"))
            if rate is None or rate < 0 or rate > 30 or line_base is None or line_amount is None:
                issues.append("invalid_tax_line")
                break
            line_bases += line_base
            line_vat += line_amount
        if base_amount is not None and abs(line_bases - base_amount) > Decimal("0.01"):
            issues.append("tax_base_mismatch")
        if vat_amount is not None and abs(line_vat - vat_amount) > Decimal("0.01"):
            issues.append("tax_amount_mismatch")
    for critical_field in ("supplier", "invoice_number", "issue_date", "totals", "taxes"):
        if not _fast_text_evidence_present(structured_data, critical_field):
            issues.append(f"missing_evidence_{critical_field}")
    return list(dict.fromkeys(issues))


_FAST_TEXT_ACCOUNTING_SAFETY_ISSUES = {
    "invalid_invoice_date",
    "missing_accounting_totals",
    "inconsistent_accounting_equation",
    "missing_tax_lines",
    "invalid_tax_line",
    "tax_base_mismatch",
    "tax_amount_mismatch",
    "missing_evidence_issue_date",
    "missing_evidence_totals",
    "missing_evidence_taxes",
}
_FAST_TEXT_METADATA_FAILURE_ISSUES = {
    "missing_supplier",
    "supplier_matches_registered_customer",
    "supplier_tax_id_matches_registered_customer",
    "supplier_matches_customer",
}


def _assess_fast_text_invoice_safety(
    structured_data: Dict[str, Any],
    normalized: Dict[str, Any],
    validation_issues: List[str],
    *,
    invoice_number_evidence_status: str,
    deterministic_corrections: List[str],
) -> Dict[str, Any]:
    """Separate fiscal integrity from document metadata without accepting V2.

    The existing validation result remains strict: only accounting-safe and
    fully confirmed metadata gets ``validation_status=passed``.  V5 persists
    both dimensions so the benchmark can measure safe coverage without
    changing the official V1 route.
    """
    issues = list(dict.fromkeys(validation_issues))
    accounting_issues = [
        issue for issue in issues if issue in _FAST_TEXT_ACCOUNTING_SAFETY_ISSUES
    ]
    metadata_issues = [
        issue for issue in issues if issue not in _FAST_TEXT_ACCOUNTING_SAFETY_ISSUES
    ]

    if not normalized.get("supplier_tax_id"):
        metadata_issues.append("missing_supplier_tax_id")
    if invoice_number_evidence_status != "confirmed":
        metadata_issues.append(f"invoice_number_evidence_{invoice_number_evidence_status}")
    payment_dates = normalized.get("payment_dates") or []
    due_dates_have_evidence = _fast_text_evidence_present(structured_data, "due_dates")
    if payment_dates and not due_dates_have_evidence and (
        "due_date_derived_from_payment_terms" not in deterministic_corrections
    ):
        metadata_issues.append("payment_dates_missing_evidence")

    metadata_issues = list(dict.fromkeys(metadata_issues))
    accounting_safety_status = "failed" if accounting_issues else "passed"
    if any(issue in _FAST_TEXT_METADATA_FAILURE_ISSUES for issue in metadata_issues):
        metadata_quality_status = "failed"
    elif metadata_issues:
        metadata_quality_status = "review_required"
    else:
        metadata_quality_status = "confirmed"

    return {
        "accounting_safety_status": accounting_safety_status,
        "accounting_safety_issues": accounting_issues,
        "metadata_quality_status": metadata_quality_status,
        "metadata_issues": metadata_issues,
    }


def analyze_invoice_v2_fast_text(
    *,
    file_bytes: bytes,
    filename: str,
    mime_type: Optional[str] = None,
    company_names: Optional[List[str]] = None,
    company_context: Optional[Dict[str, Any]] = None,
    prepared_text: Optional[Dict[str, Any]] = None,
    return_telemetry: bool = False,
) -> Union[Dict[str, Any], Tuple[Dict[str, Any], Dict[str, Any]]]:
    """Run the native-text-only V2 path. It is intended only for shadow use."""
    started = time.monotonic()
    telemetry: Dict[str, Any] = {
        "preprocessing_ms": None,
        "openai_ms": 0,
        "parsing_ms": None,
        "openai_model": None,
        "input_tokens": None,
        "output_tokens": None,
        "reasoning_tokens": None,
        "total_tokens": None,
        "processing_type": "v2_fast_text_native",
        "route": "v2_fast_text_native",
    }
    prepared = prepared_text or prepare_invoice_v2_fast_text(
        file_bytes, filename=filename, mime_type=mime_type
    )
    telemetry["preprocessing_ms"] = round((time.monotonic() - started) * 1000)
    document_layout = prepared.get("_document_layout")
    layout_build_ms = prepared.get("layout_build_ms")
    if isinstance(layout_build_ms, (int, float)) and not isinstance(layout_build_ms, bool):
        # Internal timing only. The shadow persistence layer has no layout
        # column and never receives rows, tokens or bounding boxes.
        telemetry["layout_build_ms"] = max(0, round(layout_build_ms))
    if not prepared.get("eligible"):
        result = {
            "analysis_status": "skipped",
            "validation_status": "not_applicable",
            "eligibility_reason": prepared.get("reason") or "not_eligible",
            "document_text_complete": False,
            "document_text_chars_original": max(int(prepared.get("native_text_chars") or 0), 0),
            "document_text_chars_used": max(int(prepared.get("sent_text_chars") or 0), 0),
            "document_verification": {},
            "fast_path_decision": "fallback_v1",
            "fast_path_reasons": [
                f"eligibility:{prepared.get('reason') or 'not_eligible'}"
            ],
        }
        return _invoice_analysis_telemetry_result(result, telemetry, return_telemetry)

    normalized_company_context = _normalize_fast_text_company_context(company_context)
    validation_company_names = list(company_names or [])
    if normalized_company_context.get("company_name"):
        validation_company_names.append(normalized_company_context["company_name"])
    prompt = _fast_text_invoice_prompt(normalized_company_context)
    try:
        _get_invoice_model()
        response_data = _call_invoice_responses(
            _get_client(),
            file_bytes=b"",
            filename=filename,
            mime_type=mime_type or "application/pdf",
            extracted_text="",
            prompt=prompt,
            telemetry=telemetry,
            queue_managed_rate_limits=True,
            response_input=_response_input_for_invoice_fast_text(prompt, prepared["text"]),
            response_schema=INVOICE_FAST_TEXT_SCHEMA,
            schema_name="invoice_fast_text_extraction",
            route="v2_fast_text_native",
        )
    except InvoiceAnalysisResponseError as exc:
        result = {
            "analysis_status": "failed",
            "validation_status": "not_run",
            "eligibility_reason": prepared.get("reason"),
            "analysis_error": {
                "status": exc.status,
                "detail": exc.detail,
                "metadata": exc.metadata,
            },
            "fast_path_decision": "fallback_v1",
            "fast_path_reasons": ["analysis:v2_response_unavailable"],
        }
        return _invoice_analysis_telemetry_result(result, telemetry, return_telemetry)
    except RuntimeError as exc:
        result = {
            "analysis_status": "failed",
            "validation_status": "not_run",
            "eligibility_reason": prepared.get("reason"),
            "analysis_error": {"status": "configuration_error", "detail": str(exc)},
            "fast_path_decision": "fallback_v1",
            "fast_path_reasons": ["analysis:v2_configuration_error"],
        }
        return _invoice_analysis_telemetry_result(result, telemetry, return_telemetry)

    validation_started = time.monotonic()
    normalized = _normalize_fast_text_invoice(response_data)
    correction_codes: List[str] = []
    (
        normalized["invoice_number"],
        invoice_number_corrections,
        invoice_number_issues,
        invoice_number_evidence_status,
        invoice_parser_diagnostics,
    ) = _reconcile_fast_text_invoice_number(
        normalized.get("invoice_number"), prepared["text"], document_layout=document_layout
    )
    (
        normalized["invoice_date"],
        invoice_date_corrections,
        invoice_date_issues,
        invoice_date_diagnostics,
    ) = _reconcile_fast_text_invoice_date(
        normalized.get("invoice_date"),
        prepared["text"],
        supplier_tax_id=normalized.get("supplier_tax_id"),
        registered_company_tax_id=normalized_company_context.get("company_tax_id"),
        document_layout=document_layout,
    )
    normalized["payment_dates"], due_date_corrections = _reconcile_fast_text_payment_dates(
        normalized.get("payment_dates") or [], normalized.get("invoice_date"), prepared["text"]
    )
    correction_codes.extend(invoice_number_corrections)
    correction_codes.extend(invoice_date_corrections)
    correction_codes.extend(due_date_corrections)
    validation_issues = _validate_fast_text_invoice(
        response_data,
        normalized,
        validation_company_names,
        registered_company_tax_id=normalized_company_context.get("company_tax_id"),
    )
    if {
        "invoice_date_corrected_from_explicit_label",
        "invoice_date_confirmed_from_spatial_label",
    }.intersection(correction_codes):
        # A unique date under a strict native or visual invoice label is
        # independent evidence even when the model omitted field_evidence.
        validation_issues = [
            issue for issue in validation_issues if issue != "missing_evidence_issue_date"
        ]
    validation_issues = list(
        dict.fromkeys(invoice_number_issues + invoice_date_issues + validation_issues)
    )
    safety_assessment = _assess_fast_text_invoice_safety(
        response_data,
        normalized,
        validation_issues,
        invoice_number_evidence_status=invoice_number_evidence_status,
        deterministic_corrections=correction_codes,
    )
    validation_issues = list(
        dict.fromkeys(
            validation_issues
            + safety_assessment["accounting_safety_issues"]
            + safety_assessment["metadata_issues"]
        )
    )
    validation_status = (
        "passed"
        if safety_assessment["accounting_safety_status"] == "passed"
        and safety_assessment["metadata_quality_status"] == "confirmed"
        else "failed"
    )
    document_text_complete = bool(
        prepared.get(
            "document_text_complete",
            prepared.get("sent_text_chars") == prepared.get("native_text_chars"),
        )
    )
    document_verification = _verify_fast_text_document(
        normalized,
        prepared["text"],
        document_text_complete=document_text_complete,
        registered_company_tax_id=normalized_company_context.get("company_tax_id"),
        invoice_number_evidence_status=invoice_number_evidence_status,
        invoice_parser_diagnostics=invoice_parser_diagnostics,
        invoice_date_diagnostics=invoice_date_diagnostics,
        deterministic_corrections=correction_codes,
        document_layout=document_layout,
    )
    fast_path_decision, fast_path_reasons = _decide_fast_text_fast_path(
        document_verification
    )
    telemetry["validation_ms"] = round((time.monotonic() - validation_started) * 1000)
    result = {
        "analysis_status": "ok" if validation_status == "passed" else "failed",
        "validation_status": validation_status,
        "eligibility_reason": prepared.get("reason"),
        "validation_issues": validation_issues,
        "deterministic_corrections": correction_codes,
        "invoice_number_evidence_status": invoice_number_evidence_status,
        "invoice_parser_diagnostics": {
            **invoice_parser_diagnostics,
            "invoice_date": invoice_date_diagnostics,
        },
        "document_text_complete": document_text_complete,
        "document_text_chars_original": max(int(prepared.get("native_text_chars") or 0), 0),
        "document_text_chars_used": max(int(prepared.get("sent_text_chars") or 0), 0),
        "document_verification": document_verification,
        "fast_path_decision": fast_path_decision,
        "fast_path_reasons": fast_path_reasons,
        **safety_assessment,
        **normalized,
    }
    logger.info(
        "Invoice V2 fast text: status=%s validation_status=%s accounting_safety=%s metadata_quality=%s invoice_number_evidence=%s corrections=%s pages=%s native_text_chars=%s sent_text_chars=%s layout_build_ms=%s total_elapsed_ms=%s",
        result["analysis_status"],
        result["validation_status"],
        result["accounting_safety_status"],
        result["metadata_quality_status"],
        result["invoice_number_evidence_status"],
        correction_codes,
        prepared.get("page_count"),
        prepared.get("native_text_chars"),
        prepared.get("sent_text_chars"),
        telemetry.get("layout_build_ms"),
        round((time.monotonic() - started) * 1000),
    )
    return _invoice_analysis_telemetry_result(result, telemetry, return_telemetry)


def analyze_invoice(*args, **kwargs):
    """Compatibility entry point: V1 remains the only production result path."""
    return analyze_invoice_v1(*args, **kwargs)


def extract_loan_schedule(text: str) -> List[Dict[str, Any]]:
    if not text or len(text.strip()) < 50:
        return []

    client = _get_client()
    prompt = (
        "Analiza el siguiente texto de un plan de amortización de préstamo. "
        "Devuelve SOLO JSON válido con la clave installments, que es una lista de cuotas. "
        "Cada cuota debe incluir: payment_date (YYYY-MM-DD), total_amount, interest_amount, principal_amount. "
        "Usa null si un campo no se puede inferir con seguridad. "
        "No incluyas texto adicional fuera del JSON.\n\n"
        f"TEXTO_PLAN:\n{text}"
    )

    logger.info("Prompt enviado (loan_schedule): %s", prompt)
    response = client.chat.completions.create(
        model=DEFAULT_MODEL,
        max_tokens=MAX_OUTPUT_TOKENS,
        temperature=0,
        messages=[{"role": "user", "content": prompt}],
    )

    raw_text = ""
    if response.choices:
        raw_text = response.choices[0].message.content or ""
    logger.info("Respuesta cruda modelo (loan_schedule): %s", raw_text)

    data = _extract_json(raw_text)
    items: List[Dict[str, Any]] = []
    if isinstance(data, list):
        items = data
    elif isinstance(data, dict):
        items = data.get("installments") or data.get("cuotas") or []

    normalized: List[Dict[str, Any]] = []
    for item in items or []:
        if not isinstance(item, dict):
            continue
        payment_date = _normalize_date(item.get("payment_date") or item.get("fecha_pago"))
        bank_name = (
            item.get("bank_name")
            or item.get("banco")
            or item.get("entidad")
            or item.get("bank")
        )
        total_amount = _normalize_amount(item.get("total_amount") or item.get("importe_total"))
        interest_amount = _normalize_amount(item.get("interest_amount") or item.get("interes"))
        principal_amount = _normalize_amount(item.get("principal_amount") or item.get("amortizacion"))

        if total_amount is None and principal_amount is not None and interest_amount is not None:
            total_amount = principal_amount + interest_amount
        if principal_amount is None and total_amount is not None and interest_amount is not None:
            principal_amount = total_amount - interest_amount
        if interest_amount is None and total_amount is not None and principal_amount is not None:
            interest_amount = total_amount - principal_amount

        if not payment_date or total_amount is None:
            continue

        normalized.append(
            {
                "payment_date": payment_date,
                "bank_name": str(bank_name).strip() if bank_name else None,
                "total_amount": round(total_amount, 2),
                "interest_amount": round(interest_amount or 0, 2),
                "principal_amount": round(principal_amount or 0, 2),
            }
        )

    return normalized
