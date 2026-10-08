from __future__ import annotations

import json
import re
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Mapping


CENT = Decimal("0.01")

TEXT_FIELDS = {
    "invoice_number": (("invoice_number", "invoiceNumber", "document_number", "documentNumber"), 255),
    "invoice_series": (("invoice_series", "invoiceSeries", "series"), 100),
    "rectified_invoice_reference": (("rectified_invoice_reference", "rectifiedInvoiceReference"), 255),
    "counterparty_tax_id": (("counterparty_tax_id", "counterpartyTaxId"), 100),
    "counterparty_country": (("counterparty_country", "counterpartyCountry"), 100),
    "counterparty_address": (("counterparty_address", "counterpartyAddress"), 1000),
    "concept": (("concept",), 1000),
    "document_reference": (("document_reference", "documentReference"), 255),
    "order_reference": (("order_reference", "orderReference"), 255),
    "delivery_note_reference": (("delivery_note_reference", "deliveryNoteReference"), 255),
    "source_system": (("source_system", "sourceSystem"), 100),
    "external_document_id": (("external_document_id", "externalDocumentId"), 255),
}


def _first_present(payload: Mapping[str, Any], names):
    for name in names:
        if name in payload:
            return True, payload.get(name)
    return False, None


def _optional_text(value: Any, *, field: str, max_length: int, errors: list[str]):
    if value in (None, ""):
        return None
    if isinstance(value, (dict, list, tuple)):
        errors.append(f"Formato inválido para {field}.")
        return None
    normalized = str(value).strip()
    if not normalized:
        return None
    if len(normalized) > max_length:
        errors.append(f"{field} supera la longitud permitida.")
        return None
    return normalized


def _decimal(value: Any):
    if value in (None, ""):
        return None
    try:
        return Decimal(str(value).replace(",", ".")).quantize(CENT, rounding=ROUND_HALF_UP)
    except (InvalidOperation, TypeError, ValueError):
        return None


def _json_array(value: Any):
    if value in (None, ""):
        return []
    if isinstance(value, str):
        try:
            value = json.loads(value)
        except (TypeError, ValueError):
            return None
    return value if isinstance(value, list) else None


def _iso_date(value: Any):
    if not value:
        return None
    normalized = str(value).strip().split("T", 1)[0]
    try:
        return date.fromisoformat(normalized).isoformat()
    except ValueError:
        return None


def _optional_bool(value: Any, *, field: str, errors: list[str]):
    if value in (None, ""):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and value in {0, 1}:
        return bool(value)
    normalized = str(value).strip().lower()
    if normalized in {"1", "true", "yes", "si", "sí"}:
        return True
    if normalized in {"0", "false", "no"}:
        return False
    errors.append(f"Formato inválido para {field}.")
    return False


def _normalize_tax_components(value: Any, *, field: str, errors: list[str]):
    items = _json_array(value)
    if items is None:
        errors.append(f"Formato inválido para {field}.")
        return None
    normalized = []
    for item in items:
        if not isinstance(item, dict):
            errors.append(f"Formato inválido para {field}.")
            return None
        amount = _decimal(item.get("tax_amount", item.get("amount")))
        rate = _decimal(item.get("rate"))
        base = _decimal(item.get("base"))
        if amount is None:
            errors.append(f"Importe inválido en {field}.")
            return None
        normalized.append(
            {
                "tax_type": _optional_text(
                    item.get("tax_type", item.get("type")),
                    field=f"tipo de {field}",
                    max_length=100,
                    errors=errors,
                ),
                "rate": float(rate) if rate is not None else None,
                "base": float(base) if base is not None else None,
                "tax_amount": float(abs(amount)),
            }
        )
    return normalized


def _normalize_payment_schedule(value: Any, errors: list[str]):
    items = _json_array(value)
    if items is None:
        errors.append("Formato inválido para los vencimientos.")
        return None
    normalized = []
    for item in items:
        if not isinstance(item, dict):
            errors.append("Formato inválido para los vencimientos.")
            return None
        due_date = _iso_date(item.get("due_date") or item.get("dueDate"))
        if not due_date:
            errors.append("Fecha inválida en los vencimientos.")
            return None
        amount_value = item.get("amount")
        amount = _decimal(amount_value)
        if amount_value not in (None, "") and amount is None:
            errors.append("Importe inválido en los vencimientos.")
            return None
        actual_value = item.get("actual_payment_date") or item.get("actualPaymentDate")
        actual_date = _iso_date(actual_value)
        if actual_value and not actual_date:
            errors.append("Fecha de pago real inválida en los vencimientos.")
            return None
        normalized.append(
            {
                "due_date": due_date,
                "amount": float(amount) if amount is not None else None,
                "actual_payment_date": actual_date,
            }
        )
    return normalized


def normalize_accounting_document_metadata(payload: Mapping[str, Any], *, partial=False):
    """Validate optional accounting metadata without deriving absent values."""
    errors: list[str] = []
    values = {}
    for column, (names, max_length) in TEXT_FIELDS.items():
        present, raw = _first_present(payload, names)
        if partial and not present:
            continue
        values[column] = _optional_text(
            raw, field=column, max_length=max_length, errors=errors
        )

    present, currency = _first_present(payload, ("currency",))
    if not partial or present:
        normalized_currency = _optional_text(
            currency, field="currency", max_length=3, errors=errors
        )
        if normalized_currency:
            normalized_currency = normalized_currency.upper()
            if not re.fullmatch(r"[A-Z]{3}", normalized_currency):
                errors.append("La moneda debe usar un código de tres letras.")
                normalized_currency = None
        values["currency"] = normalized_currency

    present, rectificative = _first_present(
        payload,
        (
            "is_rectificative",
            "is_rectificativa",
            "isRectificative",
            "isRectificativa",
        ),
    )
    if not partial or present:
        values["is_rectificative"] = _optional_bool(
            rectificative, field="is_rectificative", errors=errors
        )

    for column, aliases, label in (
        ("withholding_details", ("withholding_details", "withholdingDetails"), "retenciones"),
        ("other_taxes", ("other_taxes", "otherTaxes"), "otros impuestos"),
    ):
        present, raw = _first_present(payload, aliases)
        if partial and not present:
            continue
        normalized = _normalize_tax_components(raw, field=label, errors=errors)
        values[column] = json.dumps(normalized, ensure_ascii=False) if normalized else None

    present, raw_schedule = _first_present(
        payload, ("payment_schedule", "paymentSchedule")
    )
    if not partial or present:
        schedule = _normalize_payment_schedule(raw_schedule, errors)
        values["payment_schedule"] = (
            json.dumps(schedule, ensure_ascii=False) if schedule else None
        )
    return values, errors


def accounting_document_metadata_response(row: Mapping[str, Any]):
    def parsed_list(field):
        raw = row.get(field)
        if not raw:
            return []
        if isinstance(raw, list):
            return raw
        try:
            parsed = json.loads(raw)
        except (TypeError, ValueError):
            return []
        return parsed if isinstance(parsed, list) else []

    return {
        "invoice_number": row.get("invoice_number"),
        "invoice_series": row.get("invoice_series"),
        "is_rectificative": bool(row.get("is_rectificative")),
        "rectified_invoice_reference": row.get("rectified_invoice_reference"),
        "counterparty_tax_id": row.get("counterparty_tax_id"),
        "counterparty_country": row.get("counterparty_country"),
        "counterparty_address": row.get("counterparty_address"),
        "currency": row.get("currency"),
        "concept": row.get("concept"),
        "document_reference": row.get("document_reference"),
        "order_reference": row.get("order_reference"),
        "delivery_note_reference": row.get("delivery_note_reference"),
        "source_system": row.get("source_system"),
        "external_document_id": row.get("external_document_id"),
        "withholding_details": parsed_list("withholding_details"),
        "other_taxes": parsed_list("other_taxes"),
        "payment_schedule": parsed_list("payment_schedule"),
    }


def accounting_tax_components_total(value: Any) -> float:
    items = _json_array(value)
    if not items:
        return 0.0
    total = Decimal("0.00")
    for item in items:
        if not isinstance(item, dict):
            continue
        amount = _decimal(item.get("tax_amount", item.get("amount")))
        if amount is not None:
            total += amount
    return float(total.quantize(CENT, rounding=ROUND_HALF_UP))
