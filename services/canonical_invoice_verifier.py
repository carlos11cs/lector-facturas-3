"""Diagnostic invoice-number and invoice-date verification from canonical layout.

This module is deliberately independent from extraction and rollout decisions.
It consumes only ephemeral Canonical DocumentLayout geometry and returns bounded
status codes; it never returns source text, token values, or coordinates.
"""

from __future__ import annotations

from datetime import date
import re
import unicodedata
from typing import Any, Iterable

from services.document_layout import DocumentLayout, LayoutRow, LayoutToken


_NUMBER_MARKERS = {"N", "NO", "NUM", "NUMERO", "NUMBER"}
_INVOICE_WORDS = {"FACTURA", "INVOICE"}
_SECONDARY_WORDS = {
    "PEDIDO",
    "ORDER",
    "DELIVERY",
    "ALBARAN",
    "PRESUPUESTO",
    "QUOTE",
    "REFERENCIA",
    "REFERENCE",
}
_DATE_EXCLUSION_WORDS = {
    "DUE",
    "VENCIMIENTO",
    "PAYMENT",
    "PAGO",
    "DELIVERY",
    "ENTREGA",
    "ORDER",
    "PEDIDO",
    "SERVICE",
    "SERVICIO",
    "ALBARAN",
}
_PROVIDER_WORDS = {
    "PROVEEDOR",
    "EMISOR",
    "EMISORA",
    "VENDEDOR",
    "SUPPLIER",
    "ISSUER",
    "SELLER",
    "FROM",
}
_RECIPIENT_WORDS = {
    "CLIENTE",
    "RECEPTOR",
    "DESTINATARIO",
    "COMPRADOR",
    "CUSTOMER",
    "BUYER",
    "BILL",
    "TO",
}
_SPANISH_TAX_ID_PATTERN = re.compile(
    r"^(?:ES)?(?:[A-HJ-NP-SUVW]\d{7}[0-9A-J]|\d{8}[A-Z]|[XYZ]\d{7}[A-Z])$"
)
_DATE_PATTERN = re.compile(
    r"^(?:(\d{4})[./-](\d{1,2})[./-](\d{1,2})|"
    r"(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4}))$"
)


def _ascii_word(value: Any) -> str:
    normalized = unicodedata.normalize("NFKD", str(value or ""))
    normalized = "".join(
        character for character in normalized if not unicodedata.combining(character)
    )
    normalized = re.sub(r"[^A-Za-z0-9]", "", normalized).upper()
    if normalized in _NUMBER_MARKERS:
        return "NUMBER"
    return normalized


def _identifier(value: Any) -> str:
    return _ascii_word(value)


def _diagnostic(
    status: str,
    reason: str,
    *,
    match_count: int = 0,
    relation_type: str = "none",
) -> dict[str, Any]:
    return {
        "status": status,
        "reason": reason,
        "match_count": max(int(match_count), 0),
        "relation_type": relation_type,
    }


def _page_rows(layout: DocumentLayout) -> Iterable[tuple[Any, list[LayoutRow]]]:
    for page in layout.pages:
        rows = sorted(page.rows, key=lambda row: (row.bbox[1], row.bbox[0], row.token_ids))
        yield page, rows


def _row_tokens(page: Any, row: LayoutRow) -> list[LayoutToken]:
    tokens = {token.token_id: token for token in page.tokens}
    return sorted(
        (tokens[token_id] for token_id in row.token_ids if token_id in tokens),
        key=lambda token: (token.bbox[0], token.bbox[1], token.token_id),
    )


def _row_words(page: Any, row: LayoutRow) -> list[tuple[LayoutToken, str]]:
    return [
        (token, word)
        for token in _row_tokens(page, row)
        if (word := _ascii_word(token.original_text))
    ]


def _bbox(tokens: list[LayoutToken]) -> tuple[float, float, float, float]:
    return (
        min(token.bbox[0] for token in tokens),
        min(token.bbox[1] for token in tokens),
        max(token.bbox[2] for token in tokens),
        max(token.bbox[3] for token in tokens),
    )


def _invoice_number_labels(words: list[tuple[LayoutToken, str]]) -> list[dict[str, Any]]:
    labels = []
    normalized = [word for _, word in words]
    for index, word in enumerate(normalized):
        if word not in _INVOICE_WORDS:
            continue
        previous = normalized[index - 1] if index else ""
        following = normalized[index + 1] if index + 1 < len(normalized) else ""
        if previous in {"FECHA", "DATE"} or following in {"FECHA", "DATE"}:
            continue
        start = index - 1 if previous == "NUMBER" else index
        end = index + 1 if following == "NUMBER" else index
        labels.append(
            {
                "start": start,
                "end": end,
                "bbox": _bbox([token for token, _ in words[start : end + 1]]),
                "explicit_number_marker": previous == "NUMBER" or following == "NUMBER",
            }
        )
    return labels


def _invoice_date_labels(words: list[tuple[LayoutToken, str]]) -> list[dict[str, Any]]:
    normalized = [word for _, word in words]
    patterns = (
        ("FECHA", "FACTURA"),
        ("FECHA", "DE", "FACTURA"),
        ("FACTURA", "FECHA"),
        ("INVOICE", "DATE"),
        ("DOCUMENT", "DATE"),
        ("FECHA", "DOCUMENTO"),
        ("FECHA", "DEL", "DOCUMENTO"),
    )
    labels = []
    for pattern in patterns:
        width = len(pattern)
        for start in range(len(normalized) - width + 1):
            if tuple(normalized[start : start + width]) != pattern:
                continue
            labels.append(
                {
                    "start": start,
                    "end": start + width - 1,
                    "bbox": _bbox(
                        [token for token, _ in words[start : start + width]]
                    ),
                }
            )
    return labels


def _secondary_labels(words: list[tuple[LayoutToken, str]]) -> list[dict[str, Any]]:
    return [
        {"index": index, "bbox": token.bbox, "word": word}
        for index, (token, word) in enumerate(words)
        if word in _SECONDARY_WORDS
    ]


def _date_exclusion_labels(words: list[tuple[LayoutToken, str]]) -> list[dict[str, Any]]:
    return [
        {"index": index, "bbox": token.bbox, "word": word}
        for index, (token, word) in enumerate(words)
        if word in _DATE_EXCLUSION_WORDS
    ]


def _identifier_occurrences(layout: DocumentLayout, target: str) -> list[dict[str, Any]]:
    occurrences = []
    for page, rows in _page_rows(layout):
        token_map = {token.token_id: token for token in page.tokens}
        for row_index, row in enumerate(rows):
            for segment in row.segments:
                words = [
                    token_map[token_id]
                    for token_id in segment.token_ids
                    if token_id in token_map and _ascii_word(token_map[token_id].original_text)
                ]
                normalized = [_ascii_word(token.original_text) for token in words]
                for start in range(len(words)):
                    combined = ""
                    for end in range(start, min(len(words), start + 6)):
                        combined += normalized[end]
                        if combined == target:
                            occurrences.append(
                                {
                                    "page": page,
                                    "rows": rows,
                                    "row": row,
                                    "row_index": row_index,
                                    "tokens": words[start : end + 1],
                                    "bbox": _bbox(words[start : end + 1]),
                                }
                            )
                            break
                        if not target.startswith(combined):
                            break
    return occurrences


def _vertical_relation(label_bbox: tuple[float, ...], value_bbox: tuple[float, ...]) -> bool:
    label_height = max(label_bbox[3] - label_bbox[1], 1.0)
    value_height = max(value_bbox[3] - value_bbox[1], 1.0)
    gap = value_bbox[1] - label_bbox[3]
    if gap < -1 or gap > max(label_height, value_height) * 3:
        return False
    margin = max(label_height, value_height) * 8
    return value_bbox[0] <= label_bbox[2] + margin and value_bbox[2] >= label_bbox[0] - margin


def _same_row_relation(
    label_bbox: tuple[float, ...], value_bbox: tuple[float, ...]
) -> bool:
    height = max(
        label_bbox[3] - label_bbox[1],
        value_bbox[3] - value_bbox[1],
        1.0,
    )
    gap = value_bbox[0] - label_bbox[2]
    return -1 <= gap <= height * 16


def _number_relation(occurrence: dict[str, Any]) -> tuple[str, str | None]:
    page = occurrence["page"]
    row = occurrence["row"]
    words = _row_words(page, row)
    value_x0 = occurrence["bbox"][0]
    invoice_labels = [
        label
        for label in _invoice_number_labels(words)
        if label["bbox"][2] <= value_x0 + 1
        and _same_row_relation(label["bbox"], occurrence["bbox"])
    ]
    secondary = [
        label
        for label in _secondary_labels(words)
        if label["bbox"][2] <= value_x0 + 1
        and _same_row_relation(label["bbox"], occurrence["bbox"])
    ]
    nearest_invoice = max(invoice_labels, key=lambda label: label["bbox"][2], default=None)
    nearest_secondary = max(secondary, key=lambda label: label["bbox"][2], default=None)
    if nearest_secondary and (
        nearest_invoice is None
        or nearest_secondary["bbox"][2] > nearest_invoice["bbox"][2]
    ):
        return "secondary", nearest_secondary["word"].lower()
    if nearest_invoice is not None:
        return "same_row", "invoice_number_label"

    row_index = occurrence["row_index"]
    if row_index:
        previous = occurrence["rows"][row_index - 1]
        previous_words = _row_words(page, previous)
        previous_secondary = _secondary_labels(previous_words)
        previous_invoice = _invoice_number_labels(previous_words)
        if previous_secondary and _vertical_relation(previous.bbox, occurrence["bbox"]):
            return "secondary", previous_secondary[-1]["word"].lower()
        if (
            len(previous_invoice) == 1
            and _vertical_relation(previous_invoice[0]["bbox"], occurrence["bbox"])
        ):
            # A bare title can precede an address or tax ID. Across rows it is
            # strong only when the label explicitly says number.
            if previous_invoice[0]["explicit_number_marker"]:
                return "vertical", "invoice_number_label"
    return "unlabeled", None


def _explicit_invoice_number_candidates(layout: DocumentLayout) -> list[str]:
    candidates = []
    for page, rows in _page_rows(layout):
        for row_index, row in enumerate(rows):
            words = _row_words(page, row)
            for label in _invoice_number_labels(words):
                after = [
                    word
                    for token, word in words[label["end"] + 1 :]
                    if token.bbox[0] >= label["bbox"][2] - 1
                    and _same_row_relation(label["bbox"], token.bbox)
                ]
                candidate = next(
                    (word for word in after if any(character.isdigit() for character in word)),
                    None,
                )
                if candidate:
                    candidates.append(candidate)
                    continue
                if label["explicit_number_marker"] and row_index + 1 < len(rows):
                    following = rows[row_index + 1]
                    if not _vertical_relation(label["bbox"], following.bbox):
                        continue
                    following_words = _row_words(page, following)
                    candidate = next(
                        (
                            word
                            for _, word in following_words
                            if any(character.isdigit() for character in word)
                        ),
                        None,
                    )
                    if candidate:
                        candidates.append(candidate)
    return list(dict.fromkeys(candidates))


def verify_invoice_number(layout: DocumentLayout, proposed_value: Any) -> dict[str, Any]:
    target = _identifier(proposed_value)
    if not target or not any(character.isdigit() for character in target):
        return _diagnostic("review", "invoice_number_missing")

    occurrences = _identifier_occurrences(layout, target)
    relations = [_number_relation(occurrence) for occurrence in occurrences]
    strong = [relation for relation, _ in relations if relation in {"same_row", "vertical"}]
    secondary = [relation for relation, _ in relations if relation == "secondary"]
    explicit_candidates = _explicit_invoice_number_candidates(layout)
    alternatives = [
        candidate
        for candidate in explicit_candidates
        if candidate != target
    ]

    if len(explicit_candidates) > 1:
        return _diagnostic(
            "review",
            "multiple_explicit_invoice_numbers",
            match_count=len(occurrences),
            relation_type="multiple",
        )
    if len(strong) == 1 and not secondary:
        return _diagnostic(
            "confirmed",
            "unique_spatial_invoice_label",
            match_count=len(occurrences),
            relation_type=strong[0],
        )
    if len(strong) > 1:
        return _diagnostic(
            "review",
            "multiple_invoice_number_associations",
            match_count=len(occurrences),
            relation_type="multiple",
        )
    if alternatives and target not in explicit_candidates:
        return _diagnostic(
            "contradiction",
            "invoice_number_conflicts_with_explicit_label",
            match_count=len(occurrences),
            relation_type="incompatible_invoice_label",
        )
    if secondary:
        return _diagnostic(
            "review",
            "invoice_number_has_secondary_reference_context",
            match_count=len(occurrences),
            relation_type="secondary",
        )
    return _diagnostic(
        "review",
        "invoice_number_not_found" if not occurrences else "invoice_number_label_missing",
        match_count=len(occurrences),
        relation_type="none",
    )


def _spanish_supplier_context(
    layout: DocumentLayout,
    supplier_tax_id: Any,
    registered_company_tax_id: Any,
) -> bool:
    supplier = _identifier(supplier_tax_id)
    registered = _identifier(registered_company_tax_id)
    if (
        not supplier
        or not _SPANISH_TAX_ID_PATTERN.fullmatch(supplier)
        or (registered and supplier == registered)
    ):
        return False
    occurrences = _identifier_occurrences(layout, supplier)
    if len(occurrences) != 1:
        return False
    occurrence = occurrences[0]
    words = {word for _, word in _row_words(occurrence["page"], occurrence["row"])}
    if words.intersection(_RECIPIENT_WORDS):
        return False
    if words.intersection(_PROVIDER_WORDS):
        return True
    if occurrence["row_index"]:
        previous_words = {
            word
            for _, word in _row_words(
                occurrence["page"], occurrence["rows"][occurrence["row_index"] - 1]
            )
        }
        return bool(previous_words.intersection(_PROVIDER_WORDS)) and not bool(
            previous_words.intersection(_RECIPIENT_WORDS)
        )
    return False


def _date_possibilities(raw_value: str, *, spanish_supplier: bool) -> tuple[set[str], bool]:
    compact = re.sub(r"\s+", "", raw_value)
    match = _DATE_PATTERN.fullmatch(compact)
    if not match:
        return set(), False
    values = match.groups()
    possibilities = set()
    ambiguous = False
    if values[0] is not None:
        year, month, day = (int(values[0]), int(values[1]), int(values[2]))
        orders = ((year, month, day),)
    else:
        first, second, raw_year = int(values[3]), int(values[4]), values[5]
        year = int(raw_year) + (2000 if len(raw_year) == 2 else 0)
        if first <= 12 and second <= 12 and first != second:
            ambiguous = not spanish_supplier
            orders = (
                (year, second, first),
            ) if spanish_supplier else (
                (year, second, first),
                (year, first, second),
            )
        elif first > 12:
            orders = ((year, second, first),)
        elif second > 12:
            orders = ((year, first, second),)
        else:
            orders = ((year, second, first),)
    for year, month, day in orders:
        try:
            parsed = date(year, month, day)
        except ValueError:
            continue
        if 2000 <= parsed.year <= date.today().year + 2:
            possibilities.add(parsed.isoformat())
    return possibilities, ambiguous


def _row_date_occurrences(
    page: Any,
    row: LayoutRow,
    *,
    spanish_supplier: bool,
) -> list[dict[str, Any]]:
    occurrences = []
    seen = set()
    for segment in row.segments:
        token_map = {token.token_id: token for token in page.tokens}
        tokens = [token_map[token_id] for token_id in segment.token_ids if token_id in token_map]
        for start in range(len(tokens)):
            for end in range(start, min(len(tokens), start + 5)):
                raw = "".join(token.original_text for token in tokens[start : end + 1])
                possibilities, ambiguous = _date_possibilities(
                    raw, spanish_supplier=spanish_supplier
                )
                if not possibilities:
                    continue
                token_ids = tuple(token.token_id for token in tokens[start : end + 1])
                if token_ids in seen:
                    continue
                seen.add(token_ids)
                occurrences.append(
                    {
                        "possibilities": possibilities,
                        "ambiguous": ambiguous,
                        "bbox": _bbox(tokens[start : end + 1]),
                    }
                )
                break
    return occurrences


def _date_relation(
    page: Any,
    rows: list[LayoutRow],
    row_index: int,
    occurrence: dict[str, Any],
) -> str:
    row = rows[row_index]
    words = _row_words(page, row)
    value_x0 = occurrence["bbox"][0]
    invoice_labels = [
        label
        for label in _invoice_date_labels(words)
        if label["bbox"][2] <= value_x0 + 1
        and _same_row_relation(label["bbox"], occurrence["bbox"])
    ]
    exclusions = [
        label
        for label in _date_exclusion_labels(words)
        if label["bbox"][2] <= value_x0 + 1
        and _same_row_relation(label["bbox"], occurrence["bbox"])
    ]
    nearest_invoice = max(invoice_labels, key=lambda label: label["bbox"][2], default=None)
    nearest_exclusion = max(exclusions, key=lambda label: label["bbox"][2], default=None)
    if nearest_exclusion and (
        nearest_invoice is None
        or nearest_exclusion["bbox"][2] > nearest_invoice["bbox"][2]
    ):
        return "excluded"
    if nearest_invoice is not None:
        return "same_row"
    if row_index:
        previous = rows[row_index - 1]
        previous_words = _row_words(page, previous)
        if _date_exclusion_labels(previous_words) and _vertical_relation(
            previous.bbox, occurrence["bbox"]
        ):
            return "excluded"
        labels = _invoice_date_labels(previous_words)
        if len(labels) == 1 and _vertical_relation(labels[0]["bbox"], occurrence["bbox"]):
            return "vertical"
    return "unlabeled"


def verify_invoice_date(
    layout: DocumentLayout,
    proposed_value: Any,
    *,
    supplier_tax_id: Any = None,
    registered_company_tax_id: Any = None,
) -> dict[str, Any]:
    proposed = str(proposed_value or "").strip()
    try:
        target = date.fromisoformat(proposed).isoformat()
    except ValueError:
        return _diagnostic("review", "invoice_date_missing_or_invalid")

    spanish_supplier = _spanish_supplier_context(
        layout, supplier_tax_id, registered_company_tax_id
    )
    associated = []
    target_matches = 0
    excluded_matches = 0
    for page, rows in _page_rows(layout):
        for row_index, row in enumerate(rows):
            for occurrence in _row_date_occurrences(
                page, row, spanish_supplier=spanish_supplier
            ):
                relation = _date_relation(page, rows, row_index, occurrence)
                if target in occurrence["possibilities"]:
                    target_matches += 1
                    if relation == "excluded":
                        excluded_matches += 1
                if relation in {"same_row", "vertical"}:
                    associated.append({**occurrence, "relation": relation})

    if len(associated) > 1:
        return _diagnostic(
            "review",
            "multiple_invoice_date_associations",
            match_count=target_matches,
            relation_type="multiple",
        )
    if len(associated) == 1:
        evidence = associated[0]
        if evidence["ambiguous"]:
            return _diagnostic(
                "review",
                "invoice_date_format_ambiguous",
                match_count=target_matches,
                relation_type=evidence["relation"],
            )
        if target in evidence["possibilities"]:
            return _diagnostic(
                "confirmed",
                "unique_spatial_invoice_date_label",
                match_count=target_matches,
                relation_type=evidence["relation"],
            )
        return _diagnostic(
            "contradiction",
            "invoice_date_conflicts_with_explicit_label",
            match_count=target_matches,
            relation_type=evidence["relation"],
        )
    if excluded_matches:
        return _diagnostic(
            "review",
            "invoice_date_only_in_excluded_context",
            match_count=target_matches,
            relation_type="excluded",
        )
    return _diagnostic(
        "review",
        "invoice_date_not_found" if not target_matches else "invoice_date_label_missing",
        match_count=target_matches,
        relation_type="none",
    )


def verify_canonical_invoice_fields(
    layout: DocumentLayout,
    *,
    invoice_number: Any,
    invoice_date: Any,
    supplier_tax_id: Any = None,
    registered_company_tax_id: Any = None,
) -> dict[str, dict[str, Any]]:
    return {
        "invoice_number": verify_invoice_number(layout, invoice_number),
        "invoice_date": verify_invoice_date(
            layout,
            invoice_date,
            supplier_tax_id=supplier_tax_id,
            registered_company_tax_id=registered_company_tax_id,
        ),
    }


def compare_with_legacy_verification(
    canonical: dict[str, dict[str, Any]],
    legacy: dict[str, dict[str, Any]],
) -> dict[str, dict[str, Any]]:
    comparison = {}
    for field in ("invoice_number", "invoice_date"):
        legacy_field = legacy.get(field) if isinstance(legacy.get(field), dict) else {}
        canonical_field = (
            canonical.get(field) if isinstance(canonical.get(field), dict) else {}
        )
        legacy_status = str(legacy_field.get("status") or "unknown")
        canonical_status = str(canonical_field.get("status") or "unknown")
        transition = (
            "same_status"
            if legacy_status == canonical_status
            else f"legacy_{legacy_status}_to_canonical_{canonical_status}"
        )
        comparison[field] = {
            "legacy_status": legacy_status,
            "legacy_reason": str(legacy_field.get("reason") or "") or None,
            "canonical_status": canonical_status,
            "canonical_reason": str(canonical_field.get("reason") or "") or None,
            "transition": transition,
        }
    return comparison
