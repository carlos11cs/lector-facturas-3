"""Derived structural views over the immutable Canonical DocumentLayout.

The objects in this module are ephemeral. They preserve token provenance while
adding semantic anchors, typed candidates, conservative regions and geometric
relations. None of these objects is a second document authority or is suitable
for persistence.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import re
import unicodedata
from typing import Any, Iterable, Mapping

from services.document_layout import BBox, DocumentLayout, LayoutToken


_NUMBER_MARKERS = {"N", "NO", "NUM", "NUMERO", "NUMBER"}
_CURRENCY_ALIASES = {
    "EUR": "EUR",
    "EURO": "EUR",
    "EUROS": "EUR",
    "€": "EUR",
    "USD": "USD",
    "US$": "USD",
    "$": "USD",
    "GBP": "GBP",
    "£": "GBP",
}
_SPANISH_TAX_ID_PATTERN = re.compile(
    r"^(?:ES)?(?:[A-HJ-NP-SUVW]\d{7}[0-9A-J]|\d{8}[A-Z]|[XYZ]\d{7}[A-Z])$"
)
_GENERIC_TAX_ID_PATTERN = re.compile(r"^(?:[A-Z]{2})?[A-Z0-9]{8,16}$")
_DATE_PATTERN = re.compile(
    r"^(?:(\d{4})[./-](\d{1,2})[./-](\d{1,2})|"
    r"(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4}))$"
)
_MONTH_ALIASES = {
    "ENE": 1,
    "ENERO": 1,
    "JAN": 1,
    "JANUARY": 1,
    "FEB": 2,
    "FEBRERO": 2,
    "FEBRUARY": 2,
    "MAR": 3,
    "MARZO": 3,
    "MARCH": 3,
    "ABR": 4,
    "ABRIL": 4,
    "APR": 4,
    "APRIL": 4,
    "MAY": 5,
    "MAYO": 5,
    "JUN": 6,
    "JUNIO": 6,
    "JUNE": 6,
    "JUL": 7,
    "JULIO": 7,
    "JULY": 7,
    "AGO": 8,
    "AGOSTO": 8,
    "AUG": 8,
    "AUGUST": 8,
    "SEP": 9,
    "SEPT": 9,
    "SEPTIEMBRE": 9,
    "SETIEMBRE": 9,
    "SEPTEMBER": 9,
    "OCT": 10,
    "OCTUBRE": 10,
    "OCTOBER": 10,
    "NOV": 11,
    "NOVIEMBRE": 11,
    "NOVEMBER": 11,
    "DIC": 12,
    "DICIEMBRE": 12,
    "DEC": 12,
    "DECEMBER": 12,
}
_PERCENTAGE_PATTERN = re.compile(r"^[-+]?\d{1,2}(?:[.,]\d+)?%$")
_IDENTIFIER_COMPONENT_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._/\-]*$")


def normalize_word(value: Any) -> str:
    normalized = unicodedata.normalize("NFKD", str(value or ""))
    normalized = "".join(
        character for character in normalized
        if not unicodedata.combining(character)
    )
    normalized = re.sub(r"[^A-Za-z0-9]", "", normalized).upper()
    return "NUMBER" if normalized in _NUMBER_MARKERS else normalized


def normalize_identifier(value: Any) -> str:
    return normalize_word(value)


def normalize_tax_id(value: Any) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", str(value or "")).upper()


def normalize_entity(value: Any) -> str:
    normalized = unicodedata.normalize("NFKD", str(value or ""))
    normalized = "".join(
        character for character in normalized
        if not unicodedata.combining(character)
    )
    return "".join(re.findall(r"[A-Za-z0-9]+", normalized.upper()))


def is_spanish_tax_id(value: Any) -> bool:
    return bool(_SPANISH_TAX_ID_PATTERN.fullmatch(normalize_tax_id(value)))


def parse_money(value: Any) -> Decimal | None:
    raw = str(value or "").upper()
    for marker in ("EUR", "USD", "GBP", "US$", "€", "£", "$", "\u00a0", " "):
        raw = raw.replace(marker, "")
    if not raw or not re.fullmatch(r"[-+]?\d[\d.,]*", raw):
        return None
    sign = -1 if raw.startswith("-") else 1
    raw = raw.lstrip("+-")
    if "," in raw and "." in raw:
        decimal_separator = "," if raw.rfind(",") > raw.rfind(".") else "."
        thousands_separator = "." if decimal_separator == "," else ","
        normalized = raw.replace(thousands_separator, "").replace(
            decimal_separator, "."
        )
    elif "," in raw or "." in raw:
        separator = "," if "," in raw else "."
        left, right = raw.rsplit(separator, 1)
        if len(right) == 2:
            normalized = left.replace(separator, "") + "." + right
        elif len(right) == 3:
            normalized = raw.replace(separator, "")
        else:
            normalized = raw
    else:
        normalized = raw
    try:
        return (Decimal(sign) * Decimal(normalized)).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
    except (InvalidOperation, ValueError):
        return None


def date_interpretations(value: Any) -> tuple[str, ...]:
    compact = re.sub(r"\s+", "", str(value or ""))
    match = _DATE_PATTERN.fullmatch(compact)
    if not match:
        if len(re.findall(r"\d+", compact)) < 2:
            return ()
        normalized = unicodedata.normalize("NFKD", compact)
        normalized = "".join(
            character
            for character in normalized
            if not unicodedata.combining(character)
        )
        normalized = re.sub(r"[^A-Za-z0-9]", "", normalized).upper()
        day_first = re.fullmatch(
            r"(\d{1,2})(?:DE)?([A-Z]+?)(?:DE)?(\d{2}|\d{4})", normalized
        )
        month_first = re.fullmatch(
            r"([A-Z]+)(\d{1,2})(\d{2}|\d{4})", normalized
        )
        if day_first:
            raw_day, raw_month, raw_year = day_first.groups()
        elif month_first:
            raw_month, raw_day, raw_year = month_first.groups()
        else:
            return ()
        month = _MONTH_ALIASES.get(raw_month)
        if month is None:
            return ()
        year = int(raw_year) + (2000 if len(raw_year) == 2 else 0)
        try:
            parsed = date(year, month, int(raw_day))
        except ValueError:
            return ()
        return (
            (parsed.isoformat(),)
            if 2000 <= parsed.year <= date.today().year + 2
            else ()
        )
    year_first, month_first, day_first, first, second, raw_year = match.groups()
    if year_first is not None:
        possibilities = ((int(year_first), int(month_first), int(day_first)),)
    else:
        year = int(raw_year) + (2000 if len(raw_year) == 2 else 0)
        first_number, second_number = int(first), int(second)
        if first_number <= 12 and second_number <= 12 and first_number != second_number:
            possibilities = (
                (year, second_number, first_number),
                (year, first_number, second_number),
            )
        elif first_number > 12:
            possibilities = ((year, second_number, first_number),)
        elif second_number > 12:
            possibilities = ((year, first_number, second_number),)
        else:
            possibilities = ((year, second_number, first_number),)
    normalized = []
    for year, month, day in possibilities:
        try:
            parsed = date(year, month, day)
        except ValueError:
            continue
        if 2000 <= parsed.year <= date.today().year + 2:
            normalized.append(parsed.isoformat())
    return tuple(dict.fromkeys(normalized))


@dataclass(frozen=True)
class SemanticSegment:
    segment_id: str
    row_id: str
    page: int
    token_ids: tuple[str, ...]
    bbox: BBox
    normalized_terms: tuple[str, ...]
    isolation_reasons: tuple[str, ...] = ()


@dataclass(frozen=True)
class AnchorCluster:
    anchor_id: str
    anchor_type: str
    token_ids: tuple[str, ...]
    segment_ids: tuple[str, ...]
    page: int
    bbox: BBox
    strength: str
    region_hint: str


@dataclass(frozen=True)
class StructuralRegion:
    region_id: str
    region_type: str
    page: int
    segment_ids: tuple[str, ...]
    anchor_ids: tuple[str, ...]
    bbox: BBox


@dataclass(frozen=True)
class FieldCandidate:
    candidate_id: str
    candidate_type: str
    token_ids: tuple[str, ...]
    segment_id: str
    row_id: str
    page: int
    bbox: BBox
    normalized_value: str
    source: str
    possible_values: tuple[str, ...] = ()


@dataclass(frozen=True)
class EvidenceRelation:
    anchor_id: str
    candidate_id: str
    region_id: str | None
    same_page: bool
    same_region: bool
    same_segment: bool
    same_row: bool
    right_of: bool
    below: bool
    horizontal_distance: float
    vertical_distance: float
    vertical_overlap: float
    column_alignment: bool
    intervening_token_count: int
    negative_context: bool
    evidence_class: str
    score: float


@dataclass(frozen=True)
class StructuralDocument:
    segments: tuple[SemanticSegment, ...]
    anchors: tuple[AnchorCluster, ...]
    regions: tuple[StructuralRegion, ...]
    candidates: tuple[FieldCandidate, ...]
    relations: tuple[EvidenceRelation, ...]
    token_by_id: Mapping[str, LayoutToken]
    segment_region: Mapping[str, str]
    anchors_by_type: Mapping[str, tuple[AnchorCluster, ...]]
    candidates_by_type: Mapping[str, tuple[FieldCandidate, ...]]
    relations_by_anchor: Mapping[str, tuple[EvidenceRelation, ...]]
    relations_by_candidate: Mapping[str, tuple[EvidenceRelation, ...]]


@dataclass(frozen=True)
class _AnchorSpec:
    anchor_type: str
    terms: tuple[str, ...]
    strength: str
    region_hint: str


_ANCHOR_SPECS = (
    _AnchorSpec("invoice_number", ("ID", "DE", "FACTURA"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_number", ("ID", "FACTURA"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_number", ("NUMBER", "FACTURA"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_number", ("NUMBER", "DE", "FACTURA"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_number", ("NUMBER", "FACT"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_number", ("FACTURA", "NUMBER"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_number", ("FACT", "NUMBER"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_number", ("INVOICE", "NUMBER"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_number", ("INVOICE", "NO"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_number", ("FACTURA",), "contextual", "invoice_metadata"),
    _AnchorSpec("invoice_date", ("FECHA", "FACTURA"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_date", ("FECHA", "DE", "FACTURA"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_date", ("FECHA", "DE", "LA", "FACTURA"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_date", ("FACTURA", "FECHA"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_date", ("INVOICE", "DATE"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_date", ("DATE", "OF", "INVOICE"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_date", ("DOCUMENT", "DATE"), "strong", "invoice_metadata"),
    _AnchorSpec("invoice_date", ("FECHA",), "contextual", "invoice_metadata"),
    _AnchorSpec("invoice_date", ("DATE",), "contextual", "invoice_metadata"),
    _AnchorSpec("due_date", ("FECHA", "VENCIMIENTO"), "strong", "payment"),
    _AnchorSpec("due_date", ("VENCIMIENTO",), "strong", "payment"),
    _AnchorSpec("due_date", ("DUE", "DATE"), "strong", "payment"),
    _AnchorSpec("payment_date", ("PAYMENT", "DATE"), "strong", "payment"),
    _AnchorSpec("order_date", ("ORDER", "DATE"), "strong", "invoice_metadata"),
    _AnchorSpec("order_date", ("FECHA", "PEDIDO"), "strong", "invoice_metadata"),
    _AnchorSpec("order_date", ("FECHA", "DE", "PEDIDO"), "strong", "invoice_metadata"),
    _AnchorSpec("delivery_date", ("DELIVERY", "DATE"), "strong", "invoice_metadata"),
    _AnchorSpec("delivery_date", ("FECHA", "ENTREGA"), "strong", "invoice_metadata"),
    _AnchorSpec("delivery_date", ("FECHA", "DE", "ENTREGA"), "strong", "invoice_metadata"),
    _AnchorSpec("service_date", ("SERVICE", "DATE"), "strong", "invoice_metadata"),
    _AnchorSpec("service_date", ("FECHA", "SERVICIO"), "strong", "invoice_metadata"),
    _AnchorSpec("service_date", ("FECHA", "DE", "SERVICIO"), "strong", "invoice_metadata"),
    _AnchorSpec("order_reference", ("NUMBER", "PEDIDO"), "strong", "invoice_metadata"),
    _AnchorSpec("order_reference", ("PEDIDO", "NUMBER"), "strong", "invoice_metadata"),
    _AnchorSpec("order_reference", ("ORDER", "NUMBER"), "strong", "invoice_metadata"),
    _AnchorSpec("order_reference", ("PEDIDO",), "contextual", "invoice_metadata"),
    _AnchorSpec("delivery_note", ("NUMBER", "ALBARAN"), "strong", "invoice_metadata"),
    _AnchorSpec("delivery_note", ("ALBARAN", "NUMBER"), "strong", "invoice_metadata"),
    _AnchorSpec("delivery_note", ("DELIVERY", "NOTE"), "strong", "invoice_metadata"),
    _AnchorSpec("delivery_note", ("ALBARAN",), "contextual", "invoice_metadata"),
    _AnchorSpec("quotation_reference", ("PRESUPUESTO",), "strong", "invoice_metadata"),
    _AnchorSpec("quotation_reference", ("QUOTATION",), "strong", "invoice_metadata"),
    _AnchorSpec("reference", ("REFERENCIA",), "contextual", "invoice_metadata"),
    _AnchorSpec("reference", ("REFERENCE",), "contextual", "invoice_metadata"),
    _AnchorSpec("document_type_invoice", ("FACTURA",), "strong", "document_header"),
    _AnchorSpec("document_type_invoice", ("INVOICE",), "strong", "document_header"),
    _AnchorSpec("document_type_non_invoice", ("PROFORMA",), "strong", "document_header"),
    _AnchorSpec("document_type_non_invoice", ("PRESUPUESTO",), "strong", "document_header"),
    _AnchorSpec("document_type_non_invoice", ("QUOTATION",), "strong", "document_header"),
    _AnchorSpec("supplier", ("PROVEEDOR",), "strong", "supplier_identity"),
    _AnchorSpec("supplier", ("EMISOR",), "strong", "supplier_identity"),
    _AnchorSpec("supplier", ("SUPPLIER",), "strong", "supplier_identity"),
    _AnchorSpec("supplier", ("ISSUER",), "strong", "supplier_identity"),
    _AnchorSpec("recipient", ("CLIENTE",), "strong", "recipient_identity"),
    _AnchorSpec("recipient", ("RECEPTOR",), "strong", "recipient_identity"),
    _AnchorSpec("recipient", ("DESTINATARIO",), "strong", "recipient_identity"),
    _AnchorSpec("recipient", ("CUSTOMER",), "strong", "recipient_identity"),
    _AnchorSpec("recipient", ("BILL", "TO"), "strong", "recipient_identity"),
    _AnchorSpec("tax_id", ("NIF",), "contextual", "unknown"),
    _AnchorSpec("tax_id", ("CIF",), "contextual", "unknown"),
    _AnchorSpec("tax_id", ("VAT",), "contextual", "unknown"),
    _AnchorSpec("tax_id", ("TAX", "ID"), "contextual", "unknown"),
    _AnchorSpec("base_amount", ("BASE", "IMPONIBLE"), "strong", "fiscal"),
    _AnchorSpec("base_amount", ("TAXABLE", "BASE"), "strong", "fiscal"),
    _AnchorSpec("base_amount", ("BASE",), "contextual", "fiscal"),
    _AnchorSpec("vat_amount", ("TOTAL", "IVA"), "strong", "fiscal"),
    _AnchorSpec("vat_amount", ("TOTAL", "CUOTAS"), "strong", "fiscal"),
    _AnchorSpec("vat_amount", ("CUOTA", "IVA"), "strong", "fiscal"),
    _AnchorSpec("vat_amount", ("IMPORTE", "IVA"), "strong", "fiscal"),
    _AnchorSpec("vat_amount", ("IVA",), "contextual", "fiscal"),
    _AnchorSpec("vat_amount", ("VAT",), "contextual", "fiscal"),
    _AnchorSpec("vat_breakdown", ("TIPO", "IVA", "BASE", "IMPONIBLE", "CUOTA"), "strong", "fiscal"),
    _AnchorSpec("vat_breakdown", ("TIPO", "BASE", "CUOTA"), "strong", "fiscal"),
    _AnchorSpec("vat_breakdown", ("RATE", "BASE", "TAX"), "strong", "fiscal"),
    _AnchorSpec("withholding_amount", ("RETENCION", "IRPF"), "strong", "fiscal"),
    _AnchorSpec("withholding_amount", ("RETENCION",), "strong", "fiscal"),
    _AnchorSpec("withholding_amount", ("IRPF",), "strong", "fiscal"),
    _AnchorSpec("withholding_amount", ("WITHHOLDING",), "strong", "fiscal"),
    _AnchorSpec("other_taxes", ("OTROS", "IMPUESTOS"), "strong", "fiscal"),
    _AnchorSpec("other_taxes", ("OTHER", "TAXES"), "strong", "fiscal"),
    _AnchorSpec("other_taxes", ("RECARGO",), "contextual", "fiscal"),
    _AnchorSpec("subtotal", ("SUBTOTAL",), "strong", "totals"),
    _AnchorSpec("total_amount", ("TOTAL", "FACTURA"), "strong", "totals"),
    _AnchorSpec("total_amount", ("IMPORTE", "TOTAL"), "strong", "totals"),
    _AnchorSpec("total_amount", ("GRAND", "TOTAL"), "strong", "totals"),
    _AnchorSpec("total_amount", ("TOTAL",), "contextual", "totals"),
    _AnchorSpec("amount_due", ("TOTAL", "A", "PAGAR"), "strong", "totals"),
    _AnchorSpec("amount_due", ("NETO", "A", "PAGAR"), "strong", "totals"),
    _AnchorSpec("amount_due", ("AMOUNT", "DUE"), "strong", "totals"),
    _AnchorSpec("footer_marker", ("REGISTRO", "MERCANTIL"), "contextual", "footer"),
    _AnchorSpec("footer_marker", ("GENERADO", "POR", "ORDENADOR"), "contextual", "footer"),
)

_NEGATIVE_ANCHORS = {
    "order_reference",
    "delivery_note",
    "quotation_reference",
    "reference",
    "due_date",
    "payment_date",
    "order_date",
    "delivery_date",
    "service_date",
    "subtotal",
}


def _bbox_for_tokens(tokens: Iterable[LayoutToken]) -> BBox:
    boxes = [token.bbox for token in tokens]
    return (
        min(box[0] for box in boxes),
        min(box[1] for box in boxes),
        max(box[2] for box in boxes),
        max(box[3] for box in boxes),
    )


def _bbox_union(boxes: Iterable[BBox]) -> BBox:
    values = list(boxes)
    return (
        min(box[0] for box in values),
        min(box[1] for box in values),
        max(box[2] for box in values),
        max(box[3] for box in values),
    )


def _segment_height(segment: SemanticSegment) -> float:
    return max(segment.bbox[3] - segment.bbox[1], 1.0)


def _column_aligned(left: BBox, right: BBox) -> bool:
    overlap = min(left[2], right[2]) - max(left[0], right[0])
    height = max(left[3] - left[1], right[3] - right[1], 1.0)
    center_distance = abs((left[0] + left[2] - right[0] - right[2]) / 2)
    return overlap >= 0 or center_distance <= height * 6


def build_semantic_segments(
    layout: DocumentLayout,
) -> tuple[tuple[SemanticSegment, ...], dict[str, LayoutToken]]:
    token_by_id = {token.token_id: token for token in layout.tokens}
    segments = []
    for page in layout.pages:
        rows = sorted(page.rows, key=lambda row: (row.bbox[1], row.bbox[0], row.token_ids))
        for row_index, row in enumerate(rows):
            row_id = f"p{page.page}:r{row_index}"
            for segment_index, source_segment in enumerate(row.segments):
                tokens = [
                    token_by_id[token_id]
                    for token_id in source_segment.token_ids
                    if token_id in token_by_id
                ]
                if not tokens:
                    continue
                segments.append(
                    SemanticSegment(
                        segment_id=f"{row_id}:s{segment_index}",
                        row_id=row_id,
                        page=page.page,
                        token_ids=tuple(token.token_id for token in tokens),
                        bbox=source_segment.bbox,
                        normalized_terms=tuple(normalize_word(token.original_text) for token in tokens),
                        isolation_reasons=row.isolation_reasons,
                    )
                )
    return tuple(segments), token_by_id


def _matches_in_terms(terms: tuple[str, ...], pattern: tuple[str, ...]):
    width = len(pattern)
    for start in range(len(terms) - width + 1):
        if terms[start : start + width] == pattern:
            yield start, start + width


def build_anchor_clusters(
    segments: tuple[SemanticSegment, ...], token_by_id: Mapping[str, LayoutToken]
) -> tuple[AnchorCluster, ...]:
    provisional = []
    for segment in segments:
        if segment.isolation_reasons:
            continue
        for spec in _ANCHOR_SPECS:
            for start, end in _matches_in_terms(segment.normalized_terms, spec.terms):
                token_ids = segment.token_ids[start:end]
                provisional.append((spec, (segment,), token_ids))

    by_page = {}
    for segment in segments:
        by_page.setdefault(segment.page, []).append(segment)
    for page_segments in by_page.values():
        ordered = sorted(page_segments, key=lambda item: (item.bbox[1], item.bbox[0]))
        for first, second in zip(ordered, ordered[1:]):
            if first.isolation_reasons or second.isolation_reasons:
                continue
            scale = max(_segment_height(first), _segment_height(second))
            if second.row_id == first.row_id:
                horizontal_gap = second.bbox[0] - first.bbox[2]
                if horizontal_gap < -1 or horizontal_gap > scale * 3:
                    continue
            else:
                vertical_gap = second.bbox[1] - first.bbox[3]
                if (
                    vertical_gap < -1
                    or vertical_gap > scale * 2.5
                    or not _column_aligned(first.bbox, second.bbox)
                ):
                    continue
            terms = first.normalized_terms + second.normalized_terms
            for spec in _ANCHOR_SPECS:
                if terms != spec.terms:
                    continue
                provisional.append(
                    (spec, (first, second), first.token_ids + second.token_ids)
                )

    # One semantic label is one cluster. Longer/multiline matches subsume
    # overlapping fragments of the same anchor type.
    retained = []
    for spec, matched_segments, token_ids in sorted(
        provisional,
        key=lambda item: (
            item[0].anchor_type,
            item[1][0].page,
            -len(item[2]),
            item[1][0].bbox[1],
            item[1][0].bbox[0],
        ),
    ):
        token_set = set(token_ids)
        if any(
            existing[0].anchor_type == spec.anchor_type
            and token_set.intersection(existing[2])
            for existing in retained
        ):
            continue
        retained.append((spec, matched_segments, token_set))

    # A bare "TOTAL" inside a stronger fiscal label such as "TOTAL IVA" is
    # not an independent invoice-total anchor. Keeping both would make the
    # same amount compete for two incompatible fields in the totals block.
    retained = [
        item
        for item in retained
        if not (
            item[0].anchor_type == "total_amount"
            and item[0].strength == "contextual"
            and any(
                other[0].strength == "strong"
                and other[0].anchor_type
                in {"vat_amount", "withholding_amount", "other_taxes", "amount_due"}
                and item[2].intersection(other[2])
                for other in retained
            )
        )
    ]
    retained = [
        item
        for item in retained
        if not (
            item[0].anchor_type == "invoice_number"
            and item[0].strength == "contextual"
            and any(
                other[0].anchor_type == "invoice_date"
                and other[0].strength == "strong"
                and item[2].intersection(other[2])
                for other in retained
            )
        )
    ]

    anchors = []
    for index, (spec, matched_segments, token_set) in enumerate(retained, 1):
        ordered_token_ids = tuple(
            token_id
            for segment in matched_segments
            for token_id in segment.token_ids
            if token_id in token_set
        )
        anchors.append(
            AnchorCluster(
                anchor_id=f"a{index}",
                anchor_type=spec.anchor_type,
                token_ids=ordered_token_ids,
                segment_ids=tuple(segment.segment_id for segment in matched_segments),
                page=matched_segments[0].page,
                bbox=_bbox_for_tokens(token_by_id[token_id] for token_id in ordered_token_ids),
                strength=spec.strength,
                region_hint=spec.region_hint,
            )
        )
    return tuple(anchors)


def _token_windows(tokens: list[LayoutToken], maximum: int = 5):
    for start in range(len(tokens)):
        for end in range(start + 1, min(len(tokens), start + maximum) + 1):
            yield tokens[start:end]


def _tax_id_has_anchor_context(
    tokens: list[LayoutToken],
    segment: SemanticSegment,
    anchors: tuple[AnchorCluster, ...],
) -> bool:
    candidate_box = _bbox_for_tokens(tokens)
    candidate_height = max(candidate_box[3] - candidate_box[1], 1.0)
    for anchor in anchors:
        if anchor.anchor_type != "tax_id" or anchor.page != segment.page:
            continue
        anchor_height = max(anchor.bbox[3] - anchor.bbox[1], 1.0)
        scale = max(candidate_height, anchor_height)
        same_row = not (
            candidate_box[3] < anchor.bbox[1] or candidate_box[1] > anchor.bbox[3]
        )
        if (
            same_row
            and candidate_box[0] >= anchor.bbox[2] - 1
            and candidate_box[0] - anchor.bbox[2] <= scale * 10
        ):
            return True
        vertical_gap = candidate_box[1] - anchor.bbox[3]
        if (
            0 <= vertical_gap <= scale * 3
            and _column_aligned(anchor.bbox, candidate_box)
        ):
            return True
    return False


def _is_tax_id_token_window(tokens: list[LayoutToken]) -> bool:
    if len(tokens) == 1:
        return True
    if len(tokens) != 2:
        return False
    prefix = normalize_tax_id(tokens[0].original_text)
    value = normalize_tax_id(tokens[1].original_text)
    return bool(
        1 <= len(prefix) <= 2
        and prefix.isalpha()
        and len(value) >= 7
        and any(character.isdigit() for character in value)
    )


def extract_field_candidates(
    segments: tuple[SemanticSegment, ...],
    anchors: tuple[AnchorCluster, ...],
    token_by_id: Mapping[str, LayoutToken],
) -> tuple[FieldCandidate, ...]:
    anchor_token_ids = {token_id for anchor in anchors for token_id in anchor.token_ids}
    candidates = []
    seen = set()

    def add(
        candidate_type: str,
        tokens: list[LayoutToken],
        segment: SemanticSegment,
        normalized_value: str,
        *,
        possible_values: tuple[str, ...] = (),
    ):
        key = (
            candidate_type,
            tuple(token.token_id for token in tokens),
            normalized_value,
        )
        if not normalized_value or key in seen:
            return
        seen.add(key)
        candidates.append(
            FieldCandidate(
                candidate_id=f"c{len(candidates) + 1}",
                candidate_type=candidate_type,
                token_ids=key[1],
                segment_id=segment.segment_id,
                row_id=segment.row_id,
                page=segment.page,
                bbox=_bbox_for_tokens(tokens),
                normalized_value=normalized_value,
                source="native_tokens",
                possible_values=possible_values,
            )
        )

    for segment in segments:
        tokens = [token_by_id[token_id] for token_id in segment.token_ids]
        if segment.isolation_reasons:
            continue
        for window in _token_windows(tokens):
            compact = "".join(token.original_text for token in window)
            date_values = date_interpretations(compact)
            if date_values:
                add("date", window, segment, date_values[0], possible_values=date_values)
            tax_value = normalize_tax_id(compact)
            if (
                not any(token.token_id in anchor_token_ids for token in window)
                and _GENERIC_TAX_ID_PATTERN.fullmatch(tax_value)
                and any(character.isalpha() for character in tax_value)
                and any(character.isdigit() for character in tax_value)
                and _is_tax_id_token_window(window)
                and (
                    is_spanish_tax_id(tax_value)
                    or _tax_id_has_anchor_context(window, segment, anchors)
                )
            ):
                add("tax_id", window, segment, tax_value)
            separate_decimal_amounts = bool(
                len(window) == 2
                and all(
                    re.search(r"[.,]\d{2}(?:\D{0,3})$", token.original_text.strip())
                    for token in window
                )
            )
            if (
                len(window) <= 2
                and not date_values
                and "%" not in compact
                and not separate_decimal_amounts
            ):
                money = parse_money(compact)
                if money is not None:
                    add("money", window, segment, format(money, ".2f"))
            if len(window) <= 2 and _PERCENTAGE_PATTERN.fullmatch(compact.strip()):
                percentage = compact.strip().rstrip("%").replace(",", ".")
                try:
                    normalized_percentage = format(Decimal(percentage).normalize(), "f")
                except InvalidOperation:
                    normalized_percentage = ""
                add("percentage", window, segment, normalized_percentage)
            if len(window) == 1:
                currency = _CURRENCY_ALIASES.get(compact.strip().upper())
                if currency:
                    add("currency", window, segment, currency)

        # Identifier candidates are maximal runs outside semantic labels. This
        # preserves compound identifiers (e.g. "2025IR 0625") and prevents a
        # suffix such as "0625" from becoming an equivalent full candidate.
        run = []
        for token in tokens + [None]:
            raw = token.original_text.strip() if token is not None else ""
            is_component = bool(
                token is not None
                and token.token_id not in anchor_token_ids
                and _IDENTIFIER_COMPONENT_PATTERN.fullmatch(raw)
            )
            if is_component:
                run.append(token)
                continue
            if run:
                normalized = "".join(normalize_identifier(item.original_text) for item in run)
                if (
                    2 <= len(normalized) <= 64
                    and any(character.isdigit() for character in normalized)
                ):
                    add("identifier", run, segment, normalized)
                run = []

        entity_tokens = [
            token
            for token in tokens
            if token.token_id not in anchor_token_ids
            and normalize_word(token.original_text)
            and not any(character.isdigit() for character in normalize_word(token.original_text))
        ]
        if entity_tokens:
            entity_value = normalize_entity(
                " ".join(token.original_text for token in entity_tokens)
            )
            if len(entity_value) >= 3:
                add("entity", entity_tokens, segment, entity_value)

    retained = []
    for candidate in sorted(
        candidates,
        key=lambda item: (
            item.candidate_type,
            item.normalized_value,
            -len(item.token_ids),
            item.page,
            item.bbox[1],
            item.bbox[0],
        ),
    ):
        token_ids = set(candidate.token_ids)
        if any(
            existing.candidate_type == candidate.candidate_type
            and existing.normalized_value == candidate.normalized_value
            and token_ids.intersection(existing.token_ids)
            for existing in retained
        ):
            continue
        retained.append(candidate)

    # Labels such as "VAT Reg No." must not become part of the identifier.
    # Prefer the contained fiscal value when an overlapping wider candidate
    # differs only by a common number/tax marker prefix.
    tax_prefixes = {
        "N",
        "NO",
        "NUM",
        "NUMERO",
        "NUMBER",
        "NIF",
        "CIF",
        "VAT",
        "VATID",
        "REGNO",
        "VATREGNO",
    }
    filtered = []
    for candidate in retained:
        if candidate.candidate_type == "tax_id" and any(
            other is not candidate
            and other.candidate_type == "tax_id"
            and set(other.token_ids) < set(candidate.token_ids)
            and candidate.normalized_value.endswith(other.normalized_value)
            and candidate.normalized_value[: -len(other.normalized_value)]
            in tax_prefixes
            for other in retained
        ):
            continue
        filtered.append(candidate)
    return tuple(
        replace(candidate, candidate_id=f"c{index}")
        for index, candidate in enumerate(filtered, 1)
    )


def build_regions(
    segments: tuple[SemanticSegment, ...],
    anchors: tuple[AnchorCluster, ...],
    candidates: tuple[FieldCandidate, ...],
    *,
    registered_company_tax_id: Any = None,
    registered_company_name: Any = None,
) -> tuple[tuple[StructuralRegion, ...], dict[str, str]]:
    anchors_by_segment = {}
    for anchor in anchors:
        for segment_id in anchor.segment_ids:
            anchors_by_segment.setdefault(segment_id, []).append(anchor)
    candidates_by_segment = {}
    for candidate in candidates:
        candidates_by_segment.setdefault(candidate.segment_id, []).append(candidate)
    registered = normalize_tax_id(registered_company_tax_id)
    registered_name = normalize_entity(registered_company_name)

    classifications = {}
    for segment in segments:
        hints = {
            anchor.region_hint
            for anchor in anchors_by_segment.get(segment.segment_id, ())
            if anchor.region_hint != "unknown"
        }
        if registered and any(
            candidate.candidate_type == "tax_id"
            and candidate.normalized_value == registered
            for candidate in candidates_by_segment.get(segment.segment_id, ())
        ):
            hints.add("recipient_identity")
        if registered_name and any(
            candidate.candidate_type == "entity"
            and candidate.normalized_value == registered_name
            for candidate in candidates_by_segment.get(segment.segment_id, ())
        ):
            hints.add("recipient_identity")
        classifications[segment.segment_id] = (
            next(iter(hints)) if len(hints) == 1 else "unknown"
        )

    ordered = sorted(segments, key=lambda item: (item.page, item.bbox[1], item.bbox[0]))
    # A value segment can inherit a region only from one immediately adjacent,
    # aligned semantic anchor. Competing predecessors leave it unknown.
    for index, segment in enumerate(ordered):
        if classifications[segment.segment_id] != "unknown" or index == 0:
            continue
        predecessors = []
        for previous in reversed(ordered[:index]):
            if previous.page != segment.page:
                break
            gap = segment.bbox[1] - previous.bbox[3]
            if gap > max(_segment_height(previous), _segment_height(segment)) * 3:
                break
            previous_type = classifications[previous.segment_id]
            if previous_type != "unknown" and _column_aligned(previous.bbox, segment.bbox):
                predecessors.append(previous_type)
        if len(set(predecessors)) == 1:
            classifications[segment.segment_id] = predecessors[0]

    groups = []
    for segment in ordered:
        region_type = classifications[segment.segment_id]
        compatible = None
        for group in reversed(groups):
            if group["page"] != segment.page:
                break
            if group["region_type"] != region_type:
                continue
            gap = segment.bbox[1] - group["bbox"][3]
            if (
                gap <= _segment_height(segment) * 3
                and _column_aligned(group["bbox"], segment.bbox)
            ):
                compatible = group
                break
        if compatible is None:
            compatible = {
                "page": segment.page,
                "region_type": region_type,
                "segments": [],
                "bbox": segment.bbox,
            }
            groups.append(compatible)
        compatible["segments"].append(segment)
        compatible["bbox"] = _bbox_union(
            [compatible["bbox"], segment.bbox]
        )

    regions = []
    segment_region = {}
    for index, group in enumerate(groups, 1):
        segment_ids = tuple(segment.segment_id for segment in group["segments"])
        anchor_ids = tuple(
            anchor.anchor_id
            for anchor in anchors
            if any(segment_id in anchor.segment_ids for segment_id in segment_ids)
        )
        region = StructuralRegion(
            region_id=f"r{index}",
            region_type=group["region_type"],
            page=group["page"],
            segment_ids=segment_ids,
            anchor_ids=anchor_ids,
            bbox=group["bbox"],
        )
        regions.append(region)
        for segment_id in segment_ids:
            segment_region[segment_id] = region.region_id
    return tuple(regions), segment_region


def _overlap_ratio(first: BBox, second: BBox) -> float:
    overlap = max(0.0, min(first[3], second[3]) - max(first[1], second[1]))
    height = max(min(first[3] - first[1], second[3] - second[1]), 1.0)
    return min(overlap / height, 1.0)


def build_evidence_relations(
    segments: tuple[SemanticSegment, ...],
    anchors: tuple[AnchorCluster, ...],
    candidates: tuple[FieldCandidate, ...],
    segment_region: Mapping[str, str],
) -> tuple[EvidenceRelation, ...]:
    segment_by_id = {segment.segment_id: segment for segment in segments}
    candidates_by_page = {}
    for candidate in candidates:
        candidates_by_page.setdefault(candidate.page, []).append(candidate)
    relations = []
    for anchor in anchors:
        anchor_height = max(anchor.bbox[3] - anchor.bbox[1], 1.0)
        anchor_segment_ids = set(anchor.segment_ids)
        anchor_rows = {
            segment_by_id[segment_id].row_id
            for segment_id in anchor.segment_ids
            if segment_id in segment_by_id
        }
        anchor_region_ids = {
            segment_region.get(segment_id) for segment_id in anchor.segment_ids
        }
        for candidate in candidates_by_page.get(anchor.page, ()):
            candidate_height = max(candidate.bbox[3] - candidate.bbox[1], 1.0)
            scale = max(anchor_height, candidate_height)
            same_segment = candidate.segment_id in anchor_segment_ids
            vertical_overlap = _overlap_ratio(anchor.bbox, candidate.bbox)
            same_row = (
                candidate.row_id in anchor_rows
                or vertical_overlap >= 0.6
            )
            region_id = segment_region.get(candidate.segment_id)
            same_region = bool(region_id and region_id in anchor_region_ids)
            right_of = candidate.bbox[0] >= anchor.bbox[2] - 1
            below = candidate.bbox[1] >= anchor.bbox[3] - 1
            horizontal_distance = max(0.0, candidate.bbox[0] - anchor.bbox[2])
            vertical_distance = max(0.0, candidate.bbox[1] - anchor.bbox[3])
            aligned = _column_aligned(anchor.bbox, candidate.bbox)
            if not (
                same_segment
                or same_row
                or same_region
                or (below and aligned and vertical_distance <= scale * 3)
            ):
                continue
            intervening = 0
            if same_segment:
                segment = segment_by_id[candidate.segment_id]
                positions = {token_id: index for index, token_id in enumerate(segment.token_ids)}
                anchor_positions = [positions[token_id] for token_id in anchor.token_ids if token_id in positions]
                candidate_positions = [positions[token_id] for token_id in candidate.token_ids if token_id in positions]
                if anchor_positions and candidate_positions:
                    intervening = max(0, min(candidate_positions) - max(anchor_positions) - 1)
            if (same_segment or same_row) and right_of and horizontal_distance <= scale * 10:
                evidence_class = "strong_same_row"
                score = 100 - min(horizontal_distance / scale, 40) - intervening * 3
            elif below and aligned and vertical_distance <= scale * 3:
                evidence_class = "strong_below"
                score = 82 - min(vertical_distance / scale, 30)
            elif same_region:
                evidence_class = "contextual_region"
                score = 45
            else:
                evidence_class = "weak_proximity"
                score = 20
            relations.append(
                EvidenceRelation(
                    anchor_id=anchor.anchor_id,
                    candidate_id=candidate.candidate_id,
                    region_id=region_id,
                    same_page=True,
                    same_region=same_region,
                    same_segment=same_segment,
                    same_row=same_row,
                    right_of=right_of,
                    below=below,
                    horizontal_distance=round(horizontal_distance, 3),
                    vertical_distance=round(vertical_distance, 3),
                    vertical_overlap=round(vertical_overlap, 4),
                    column_alignment=aligned,
                    intervening_token_count=intervening,
                    negative_context=anchor.anchor_type in _NEGATIVE_ANCHORS,
                    evidence_class=evidence_class,
                    score=round(score, 3),
                )
            )
    return tuple(relations)


def assemble_structural_document(
    segments: tuple[SemanticSegment, ...],
    anchors: tuple[AnchorCluster, ...],
    regions: tuple[StructuralRegion, ...],
    candidates: tuple[FieldCandidate, ...],
    relations: tuple[EvidenceRelation, ...],
    token_by_id: Mapping[str, LayoutToken],
    segment_region: Mapping[str, str],
) -> StructuralDocument:
    anchors_by_type = {}
    candidates_by_type = {}
    relations_by_anchor = {}
    relations_by_candidate = {}
    for anchor in anchors:
        anchors_by_type.setdefault(anchor.anchor_type, []).append(anchor)
    for candidate in candidates:
        candidates_by_type.setdefault(candidate.candidate_type, []).append(candidate)
    for relation in relations:
        relations_by_anchor.setdefault(relation.anchor_id, []).append(relation)
        relations_by_candidate.setdefault(relation.candidate_id, []).append(relation)
    return StructuralDocument(
        segments=segments,
        anchors=anchors,
        regions=regions,
        candidates=candidates,
        relations=relations,
        token_by_id=token_by_id,
        segment_region=dict(segment_region),
        anchors_by_type={key: tuple(value) for key, value in anchors_by_type.items()},
        candidates_by_type={key: tuple(value) for key, value in candidates_by_type.items()},
        relations_by_anchor={key: tuple(value) for key, value in relations_by_anchor.items()},
        relations_by_candidate={key: tuple(value) for key, value in relations_by_candidate.items()},
    )
