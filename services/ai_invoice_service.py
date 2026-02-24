import gc
import json
import logging
import os
import re
import time
from datetime import date, timedelta
from typing import Any, Dict, Optional, List, Tuple

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

DEFAULT_MODEL = os.getenv("OPENAI_CHAT_MODEL", os.getenv("OPENAI_VISION_MODEL", "gpt-4o-mini"))
MAX_OUTPUT_TOKENS = int(os.getenv("OPENAI_MAX_OUTPUT_TOKENS", "500"))
PDF_TEXT_THRESHOLD = int(os.getenv("PDF_TEXT_THRESHOLD", "100"))
PDF_OCR_ZOOM = float(os.getenv("PDF_OCR_ZOOM", "2.0"))
OCR_MAX_PAGES = int(os.getenv("OCR_MAX_PAGES", "5"))
OCR_MAX_SECONDS = int(os.getenv("OCR_MAX_SECONDS", "7"))
OCR_MAX_DIM = int(os.getenv("OCR_MAX_DIM", "1600"))
_client: Optional[OpenAI] = None
_ocr_reader = None
_EU_AMOUNT_RE = re.compile(r"\d{1,3}(?:[.\s]\d{3})*,\d{2}|\d+,\d{2}")
_EU_THOUSANDS_RE = re.compile(r"^\d{1,3}\.\d{3},\d{2}$")


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
        if any(keyword in lowered for keyword in keywords):
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
    unique_dates = sorted({d for d in dates if d})
    return unique_dates


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
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    def pick_best_amount(numbers: List[str]) -> Optional[float]:
        values = [_normalize_amount(value) for value in numbers]
        values = [value for value in values if value is not None]
        if not values:
            return None
        large_values = [value for value in values if value > 30]
        if large_values:
            return max(large_values)
        return values[-1]

    def find_amount_for_keywords(
        keywords: List[str],
        *,
        forbid_if_contains: Optional[List[str]] = None,
        require_currency_on_keyword_line: bool = False,
        require_single_amount: bool = False,
    ) -> Optional[float]:
        for idx, line in enumerate(lines):
            upper = line.upper()
            if any(keyword in upper for keyword in keywords):
                if forbid_if_contains and any(token in upper for token in forbid_if_contains):
                    continue
                amount = None
                numbers = re.findall(r"\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})|\d+[.,]\d{2}", line)
                if require_currency_on_keyword_line and not numbers:
                    if "€" not in line and "EUR" not in upper:
                        continue
                if require_single_amount and len(numbers) > 1:
                    continue
                if numbers:
                    amount = pick_best_amount(numbers)
                if amount is None and idx + 1 < len(lines):
                    candidates = []
                    for offset in range(1, 8):
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
                if amount is not None:
                    return amount
        return None

    base_amount = find_amount_for_keywords(
        ["BASE IMPONIBLE", "BASE IVA", "BASE", "TOTAL BRUTO"]
    )
    total_amount = find_amount_for_keywords(
        [
            "TOTAL FACTURA",
            "TOTAL IVA INCLUIDO",
            "TOTAL CON IVA",
            "TOTAL A PAGAR",
            "TOTAL EUR",
            "TOTAL",
        ],
        forbid_if_contains=["BRUTO", "BASE", "IMPONIBLE", "I.V.A", "IVA", "REC.EQUIV"],
        require_single_amount=True,
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
    return {"base": base_amount, "vat": vat_amount, "total": total_amount}


def _extract_invoice_date_from_text(text: str) -> Optional[str]:
    if not text:
        return None
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for idx, line in enumerate(lines):
        lowered = line.lower()
        if "vencimiento" in lowered or "fecha de pago" in lowered:
            continue
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
            for offset in range(1, 11):
                if idx + offset < len(lines):
                    candidate = _extract_first_date(lines[idx + offset])
                    if candidate:
                        return candidate
    return None


def _extract_tax_summary_from_text(text: str) -> Dict[str, Any]:
    if not text:
        return {"found": False}
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
        return {"found": False}
    block = lines[start_idx : start_idx + 20]

    def find_amount_after_keywords(
        keywords: List[str],
        *,
        forbid_if_contains: Optional[List[str]] = None,
        prefer_last: bool = False,
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
    )
    vat_value, vat_raw = find_amount_after_keywords(
        ["I.V.A", "IVA"],
        forbid_if_contains=["REC", "RECARGO"],
        prefer_last=False,
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
) -> Tuple[Optional[float], Optional[float], Optional[float], Optional[float], str]:
    source = "llm"
    if not summary or not summary.get("found"):
        return base_amount, vat_amount, total_amount, vat_rate, source

    summary_base = summary.get("base_amount")
    summary_vat = summary.get("vat_amount")
    summary_total = summary.get("total_amount")
    summary_rate = summary.get("vat_rate") or vat_rate
    summary_base_raw = summary.get("base_raw")

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
    if summary_rate is None and summary_base and summary_vat:
        summary_rate = round(100 * summary_vat / summary_base, 2)

    summary_math_ok = (
        summary_base is not None
        and summary_vat is not None
        and summary_total is not None
        and within_tolerance(summary_base + summary_vat, summary_total)
    )
    llm_math_ok = _validate_math(base_amount, vat_amount, total_amount).get("is_consistent")

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
        base_amount = summary_base or base_amount
        vat_amount = summary_vat or vat_amount
        total_amount = summary_total or total_amount
        vat_rate = summary_rate or vat_rate
        source = "regex_tax_summary"
        return base_amount, vat_amount, total_amount, vat_rate, source

    return base_amount, vat_amount, total_amount, vat_rate, source


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

    math_ok = _validate_math(base_amount, vat_amount, total_amount)
    text_math_ok = _validate_math(text_base, text_vat, text_total)
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
    compact = re.sub(r"[\\s\\.]", "", name).upper()
    legal_tokens = [
        "SLU",
        "SL",
        "SLL",
        "SAU",
        "SA",
        "SLP",
        "SCP",
        "SC",
        "SCOOP",
        "COOP",
        "COOPERATIVA",
        "AIE",
        "UTE",
        "CB",
        "LTD",
        "LIMITED",
        "INC",
        "GMBH",
        "SARL",
        "BV",
        "NV",
        "SAS",
        "SRL",
    ]
    return any(token in compact for token in legal_tokens)


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
    if looks_like_person(value):
        return False
    if contains_forbidden_keyword(value):
        return False
    has_form = has_legal_form(value)
    has_tax = bool(text and _supplier_has_near_tax_id_or_iban(text, value))
    if not has_form:
        return False
    if _is_same_entity(value, company_names):
        return False
    if require_tax_id and text and not has_tax:
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

    for line in lines:
        lowered = line.lower()
        if "en nombre de" in lowered:
            match = re.split(r"en nombre de", line, flags=re.IGNORECASE)
            if len(match) > 1:
                candidate = match[1].strip(" :-")
                if _is_valid_supplier(candidate, company_names, text, require_tax_id=False):
                    return candidate

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


def _confidence_score_for_source(source: Optional[str]) -> Optional[float]:
    if not source:
        return None
    mapping = {
        "regex_tax_summary": 0.98,
        "llm": 0.85,
        "fallback": 0.60,
    }
    return mapping.get(source)


def _group_standard_vat_rate(rate: Optional[float], tolerance: float = 0.25) -> Optional[float]:
    if rate is None:
        return None
    for standard in (0.0, 4.0, 10.0, 21.0):
        if abs(rate - standard) <= tolerance:
            return standard
    return rate


def normalize_and_validate_amounts(extracted: Dict[str, Any]) -> Dict[str, Any]:
    result = dict(extracted or {})
    totals_payload = result.get("totals") if isinstance(result.get("totals"), dict) else {}

    base_amount = _normalize_amount(
        totals_payload.get("base")
        or totals_payload.get("base_amount")
        or result.get("base_amount")
    )
    vat_amount = _normalize_amount(
        totals_payload.get("vat")
        or totals_payload.get("vat_amount")
        or result.get("vat_amount")
    )
    total_amount = _normalize_amount(
        totals_payload.get("total")
        or totals_payload.get("total_amount")
        or result.get("total_amount")
    )
    vat_rate = _normalize_rate(result.get("vat_rate"))

    raw_breakdown = result.get("vat_breakdown") or []
    if isinstance(raw_breakdown, str):
        try:
            raw_breakdown = json.loads(raw_breakdown)
        except json.JSONDecodeError:
            raw_breakdown = []
    if isinstance(raw_breakdown, dict):
        raw_breakdown = [raw_breakdown]

    normalized_breakdown: List[Dict[str, Any]] = []
    for line in raw_breakdown:
        if not isinstance(line, dict):
            continue
        base_line = _normalize_amount(line.get("base") or line.get("base_amount"))
        vat_line = _normalize_amount(line.get("vat_amount") or line.get("vat"))
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

    amount_source = "llm"
    if normalized_breakdown:
        base_sum = round(sum(line["base"] or 0 for line in normalized_breakdown), 2)
        vat_sum = round(sum(line["vat_amount"] or 0 for line in normalized_breakdown), 2)
        total_sum = round(base_sum + vat_sum, 2)

        def totals_match(a: Optional[float], b: Optional[float]) -> bool:
            if a is None or b is None:
                return False
            return abs(a - b) <= 0.02

        if (
            base_amount is None
            or vat_amount is None
            or total_amount is None
            or not totals_match(base_amount, base_sum)
            or not totals_match(vat_amount, vat_sum)
            or not totals_match(total_amount, total_sum)
        ):
            base_amount = base_sum
            vat_amount = vat_sum
            total_amount = total_sum

    analysis_status = result.get("analysis_status") or "ok"
    if (
        base_amount is None
        or vat_amount is None
        or total_amount is None
        or abs((base_amount + vat_amount) - total_amount) > 0.02
    ):
        analysis_status = "partial"
        base_amount = None
        vat_amount = None
        total_amount = None
        vat_rate = None
        normalized_breakdown = []
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
            "analysis_status": analysis_status,
            "amount_source": amount_source,
        }
    )
    return result


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
    runtime_env = os.getenv("ENV", "").strip().lower()
    max_pages = OCR_MAX_PAGES
    if runtime_env == "production" and OCR_MAX_PAGES > 2 and "OCR_MAX_PAGES" not in os.environ:
        max_pages = 2
    with fitz.open(file_path) as doc:
        for idx, page in enumerate(doc):
            if idx >= max_pages:
                break
            base_scale = PDF_OCR_ZOOM
            if runtime_env == "production" and base_scale > 1.4 and "PDF_OCR_ZOOM" not in os.environ:
                base_scale = 1.4
            max_dim = max(page.rect.width, page.rect.height, 1)
            if max_dim * base_scale > OCR_MAX_DIM:
                base_scale = OCR_MAX_DIM / max_dim
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
    runtime_env = os.getenv("ENV", "").strip().lower()
    max_pages = OCR_MAX_PAGES
    if runtime_env == "production" and OCR_MAX_PAGES > 2 and "OCR_MAX_PAGES" not in os.environ:
        max_pages = 2
    with fitz.open(stream=data, filetype="pdf") as doc:
        for idx, page in enumerate(doc):
            if idx >= max_pages:
                break
            if time.time() - start_time > OCR_MAX_SECONDS:
                break
            base_scale = PDF_OCR_ZOOM
            if runtime_env == "production" and base_scale > 1.4 and "PDF_OCR_ZOOM" not in os.environ:
                base_scale = 1.4
            max_dim = max(page.rect.width, page.rect.height, 1)
            if max_dim * base_scale > OCR_MAX_DIM:
                base_scale = OCR_MAX_DIM / max_dim
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
    height, width = image.shape[:2]
    max_dim = max(width, height, 1)
    if max_dim > OCR_MAX_DIM:
        scale = OCR_MAX_DIM / max_dim
        image = cv2.resize(
            image, (int(width * scale), int(height * scale)), interpolation=cv2.INTER_AREA
        )
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
    known_suppliers: Optional[List[str]] = None,
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

    analysis_status = "ok"
    if used_ocr and _is_low_quality_ocr(extracted_text) and not _has_amount_hints(extracted_text):
        analysis_status = "low_quality_scan"
        logger.warning(
            "OCR de baja calidad (%s). Se omite analisis IA.", filename
        )
        return {
            "analysis_status": analysis_status,
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

    is_income = document_type == "income"
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
            "No incluyas texto adicional fuera del JSON.\n\n"
            f"TEXTO_FACTURA:\n{extracted_text}"
        )
    else:
        prompt = (
            "Analiza el siguiente texto extraido de una factura recibida (gasto). "
            "Devuelve SOLO JSON valido con estas claves: "
            "supplier, invoice_date, payment_terms_days, payment_dates, totals, vat_breakdown. "
            "totals es un objeto con {base, vat, total} (pueden ser null). "
            "vat_breakdown es una lista de lineas IVA con {base, vat_amount} y opcional {rate}. "
            "Si hay varias lineas IVA, NO rellenes un vat_rate unico (deja rate en cada linea o null). "
            "El supplier debe ser la razon social del emisor (forma juridica si aparece) "
            "y no debe ser el cliente/receptor. "
            "payment_terms_days es el numero de dias si aparece una condicion tipo "
            "\"RECIBO X DIAS FECHA FACTURA\". "
            "Si hay payment_terms_days y invoice_date, devuelve payment_dates con invoice_date + X dias. "
            "payment_dates debe ser una lista de fechas (YYYY-MM-DD) y puede estar vacia. "
            "Usa null si no puedes inferir un dato con seguridad. "
            "No incluyas texto adicional fuera del JSON.\n\n"
            f"TEXTO_FACTURA:\n{extracted_text}"
        )

    logger.info("Prompt enviado (%s): %s", filename, prompt)

    response = client.chat.completions.create(
        model=DEFAULT_MODEL,
        max_tokens=MAX_OUTPUT_TOKENS,
        temperature=0,
        top_p=1,
        seed=42,
        messages=[
            {"role": "user", "content": prompt},
        ],
    )

    raw_text = ""
    if response.choices:
        raw_text = response.choices[0].message.content or ""
    logger.info("Respuesta cruda modelo (%s): %s", filename, raw_text)

    data = _extract_json(raw_text)
    if not data:
        logger.warning("No se pudo extraer JSON valido (%s). Se usara regex/fallback.", filename)

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
    if invoice_date is None:
        invoice_date = _extract_invoice_date_from_text(extracted_text)
    payment_terms_days = data.get("payment_terms_days") or data.get("payment_terms")
    try:
        payment_terms_days = int(payment_terms_days) if payment_terms_days is not None else None
    except (TypeError, ValueError):
        payment_terms_days = None

    payment_dates: List[str] = []
    raw_payment_dates = (
        data.get("payment_dates")
        or data.get("fechas_pago")
        or data.get("fechas_vencimiento")
        or data.get("vencimientos")
    )
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

    single_payment_date = _normalize_date(
        data.get("payment_date")
        or data.get("fecha_pago")
        or data.get("fecha_vencimiento")
        or data.get("vencimiento")
    )
    if single_payment_date:
        payment_dates.append(single_payment_date)
    totals_payload = data.get("totals") if isinstance(data.get("totals"), dict) else {}
    base_amount = _normalize_amount(
        totals_payload.get("base")
        or totals_payload.get("base_amount")
        or data.get("base_amount")
        or data.get("base_imponible")
        or data.get("base")
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
        totals_payload.get("vat")
        or totals_payload.get("vat_amount")
        or data.get("vat_amount")
        or data.get("importe_iva")
        or data.get("iva_importe")
    )
    total_amount = _normalize_amount(
        totals_payload.get("total")
        or totals_payload.get("total_amount")
        or data.get("total_amount")
        or data.get("total_factura")
        or data.get("total")
    )
    vat_breakdown = (
        data.get("vat_breakdown")
        or data.get("iva_breakdown")
        or data.get("vat_lines")
        or data.get("iva_lines")
        or []
    )

    if company_names is None:
        company_names = []

    if document_type != "income":
        supplier_source_text = embedded_text if pdf_kind == "original" else extracted_text
        provider_name = provider_name.strip() if isinstance(provider_name, str) else provider_name
        if provider_name is not None and not _is_valid_supplier(
            provider_name, company_names, supplier_source_text, require_tax_id=False
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

    if not payment_dates and payment_terms_days is None:
        payment_terms_days = extract_payment_terms_days(extracted_text)
    if not payment_dates and payment_terms_days is not None and invoice_date:
        try:
            base_date = date.fromisoformat(invoice_date)
            payment_dates = [
                (base_date + timedelta(days=payment_terms_days)).isoformat()
            ]
        except ValueError:
            payment_dates = []
    if not payment_dates:
        payment_dates = _find_payment_dates_by_keywords(extracted_text, invoice_date)
    payment_dates = sorted({d for d in payment_dates if d})
    payment_date = payment_dates[0] if payment_dates else None

    normalized = normalize_and_validate_amounts(
        {
            "analysis_status": analysis_status,
            "base_amount": base_amount,
            "vat_amount": vat_amount,
            "total_amount": total_amount,
            "vat_rate": vat_rate,
            "vat_breakdown": vat_breakdown,
            "totals": totals_payload,
        }
    )

    analysis_status = normalized.get("analysis_status") or analysis_status
    base_amount = normalized.get("base_amount")
    vat_amount = normalized.get("vat_amount")
    total_amount = normalized.get("total_amount")
    vat_rate = normalized.get("vat_rate")
    vat_breakdown = normalized.get("vat_breakdown") or []
    amount_source = normalized.get("amount_source") or ("llm" if data else "fallback")

    validation = _validate_math(base_amount, vat_amount, total_amount)

    confidence_score = _confidence_score_for_source(amount_source)
    logger.info("Fuente importes (%s): %s", filename, amount_source)
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
        "analysis_status": analysis_status,
        "supplier": provider_name,
        "provider_name": provider_name,
        "client_name": client_name,
        "invoice_date": invoice_date,
        "payment_terms_days": payment_terms_days,
        "payment_dates": payment_dates,
        "payment_date": payment_date,
        "base_amount": base_amount,
        "vat_rate": vat_rate,
        "vat_amount": vat_amount,
        "total_amount": total_amount,
        "vat_breakdown": vat_breakdown,
        "extraction_source": amount_source,
        "confidence_score": confidence_score,
        "analysis_text": raw_text[:500],
        "validation": validation,
    }


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
