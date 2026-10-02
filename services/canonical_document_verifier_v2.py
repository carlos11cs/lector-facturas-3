"""Second-generation canonical document verifier for shadow diagnostics only."""

from __future__ import annotations

from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import time
from typing import Any, Iterable

from services.canonical_document_structure import (
    AnchorCluster,
    EvidenceRelation,
    FieldCandidate,
    StructuralDocument,
    assemble_structural_document,
    build_anchor_clusters,
    build_evidence_relations,
    build_regions,
    build_semantic_segments,
    extract_field_candidates,
    is_spanish_tax_id,
    normalize_entity,
    normalize_identifier,
    normalize_tax_id,
)
from services.document_layout import DocumentLayout


VERIFIER_VERSION = "canonical_document_verifier_v2"
_STRONG_EVIDENCE = {"strong_same_row", "strong_below"}
_INVOICE_NUMBER_NEGATIVE = {
    "order_reference",
    "delivery_note",
    "quotation_reference",
    "reference",
}
_INVOICE_DATE_NEGATIVE = {
    "due_date",
    "payment_date",
    "order_date",
    "delivery_date",
    "service_date",
}
_MONEY_FIELD_ANCHORS = {
    "base_amount": {"base_amount"},
    "vat_amount": {"vat_amount"},
    "withholding_amount": {"withholding_amount"},
    "other_taxes": {"other_taxes"},
    "total_amount": {"total_amount"},
}


def _diagnostic(
    status: str,
    reason_code: str,
    *,
    value_occurrences: int = 0,
    anchor_clusters: int = 0,
    valid_associations: int = 0,
    competing_associations: int = 0,
    relation_type: str = "none",
    evidence_class: str = "none",
) -> dict[str, Any]:
    return {
        "status": status,
        "reason_code": reason_code,
        "value_occurrences": max(int(value_occurrences), 0),
        "anchor_clusters": max(int(anchor_clusters), 0),
        "valid_associations": max(int(valid_associations), 0),
        "competing_associations": max(int(competing_associations), 0),
        "relation_type": relation_type,
        "evidence_class": evidence_class,
    }


def _candidate_map(structure: StructuralDocument) -> dict[str, FieldCandidate]:
    return {candidate.candidate_id: candidate for candidate in structure.candidates}


def _anchor_map(structure: StructuralDocument) -> dict[str, AnchorCluster]:
    return {anchor.anchor_id: anchor for anchor in structure.anchors}


def _relation_shadowed_by_competing_context(
    structure: StructuralDocument,
    relation: EvidenceRelation,
    competing_anchor_types: set[str],
) -> bool:
    if relation.evidence_class != "strong_below":
        return False
    anchors = _anchor_map(structure)
    return any(
        other.anchor_id != relation.anchor_id
        and other.evidence_class == "strong_same_row"
        and other.score > relation.score
        and (anchor := anchors.get(other.anchor_id)) is not None
        and anchor.anchor_type in competing_anchor_types
        for other in structure.relations_by_candidate.get(relation.candidate_id, ())
    )


def _relations(
    structure: StructuralDocument,
    candidates: Iterable[FieldCandidate],
    anchor_types: set[str],
    *,
    strong_only: bool = True,
    competing_anchor_types: set[str] | None = None,
) -> list[EvidenceRelation]:
    anchors = _anchor_map(structure)
    found = []
    for candidate in candidates:
        for relation in structure.relations_by_candidate.get(candidate.candidate_id, ()):
            anchor = anchors.get(relation.anchor_id)
            if anchor is None or anchor.anchor_type not in anchor_types:
                continue
            if strong_only and relation.evidence_class not in _STRONG_EVIDENCE:
                continue
            found.append(relation)
    grouped = {}
    for relation in found:
        grouped.setdefault(relation.anchor_id, []).append(relation)
    material = []
    for anchor_relations in grouped.values():
        same_row = [
            relation
            for relation in anchor_relations
            if relation.evidence_class == "strong_same_row"
        ]
        material.extend(same_row or anchor_relations)
    if not competing_anchor_types:
        return material
    return [
        relation
        for relation in material
        if not _relation_shadowed_by_competing_context(
            structure, relation, competing_anchor_types
        )
    ]


def _relation_shadowed_by_local_candidate(
    structure: StructuralDocument,
    relation: EvidenceRelation,
    candidate_type: str,
) -> bool:
    if relation.evidence_class != "strong_below":
        return False
    candidates = _candidate_map(structure)
    return any(
        local.evidence_class == "strong_same_row"
        and (candidate := candidates.get(local.candidate_id)) is not None
        and candidate.candidate_type == candidate_type
        for local in structure.relations_by_anchor.get(relation.anchor_id, ())
    )


def _relation_summary(
    relation: EvidenceRelation | None,
) -> tuple[str, str]:
    if relation is None:
        return "none", "none"
    if relation.same_segment:
        relation_type = "same_segment"
    elif relation.below:
        relation_type = "below"
    elif relation.same_region:
        relation_type = "same_region"
    else:
        relation_type = "proximity"
    return relation_type, relation.evidence_class


def _dominant_relation(
    relations: list[EvidenceRelation],
) -> tuple[EvidenceRelation | None, bool]:
    if not relations:
        return None, False
    ordered = sorted(relations, key=lambda item: item.score, reverse=True)
    if len(ordered) == 1:
        return ordered[0], True
    return ordered[0], ordered[0].score >= ordered[1].score + 15


def _distinct_relation_values(
    structure: StructuralDocument,
    relations: Iterable[EvidenceRelation],
) -> set[str]:
    candidates = _candidate_map(structure)
    return {
        candidate.normalized_value
        for relation in relations
        if (candidate := candidates.get(relation.candidate_id)) is not None
    }


def _matching_identifier_candidates(
    structure: StructuralDocument, value: Any
) -> list[FieldCandidate]:
    target = normalize_identifier(value)
    if not target:
        return []
    return [
        candidate
        for candidate in structure.candidates_by_type.get("identifier", ())
        if candidate.normalized_value == target
    ]


def _explicit_alternative_relations(
    structure: StructuralDocument,
    anchor_type: str,
    matching_candidates: list[FieldCandidate],
    *,
    competing_anchor_types: set[str],
) -> list[EvidenceRelation]:
    matching_token_sets = [set(candidate.token_ids) for candidate in matching_candidates]
    candidate_map = _candidate_map(structure)
    alternatives = []
    for anchor in structure.anchors_by_type.get(anchor_type, ()):
        for relation in structure.relations_by_anchor.get(anchor.anchor_id, ()):
            if relation.evidence_class not in _STRONG_EVIDENCE:
                continue
            candidate = candidate_map.get(relation.candidate_id)
            if candidate is None or candidate.candidate_type != "identifier":
                continue
            if _relation_shadowed_by_local_candidate(
                structure, relation, "identifier"
            ):
                continue
            if _relation_shadowed_by_competing_context(
                structure, relation, competing_anchor_types
            ):
                continue
            candidate_tokens = set(candidate.token_ids)
            if any(candidate_tokens.intersection(tokens) for tokens in matching_token_sets):
                continue
            alternatives.append(relation)
    return alternatives


def resolve_invoice_number(
    structure: StructuralDocument, proposed_value: Any
) -> dict[str, Any]:
    anchors = list(structure.anchors_by_type.get("invoice_number", ()))
    matches = _matching_identifier_candidates(structure, proposed_value)
    positive = _relations(
        structure,
        matches,
        {"invoice_number"},
        competing_anchor_types=_INVOICE_NUMBER_NEGATIVE,
    )
    negative = _relations(
        structure,
        matches,
        _INVOICE_NUMBER_NEGATIVE,
        competing_anchor_types={"invoice_number"},
    )
    alternatives = _explicit_alternative_relations(
        structure,
        "invoice_number",
        matches,
        competing_anchor_types=_INVOICE_NUMBER_NEGATIVE,
    )
    best, dominant = _dominant_relation(positive)
    relation_type, evidence_class = _relation_summary(best)
    competing = len(alternatives) + len(negative)

    if best is not None and dominant and not alternatives and not negative:
        return _diagnostic(
            "confirmed",
            "unique_or_dominant_invoice_number_association",
            value_occurrences=len(matches),
            anchor_clusters=len(anchors),
            valid_associations=len(positive),
            competing_associations=0,
            relation_type=relation_type,
            evidence_class=evidence_class,
        )
    if positive:
        return _diagnostic(
            "review",
            "invoice_number_associations_compete",
            value_occurrences=len(matches),
            anchor_clusters=len(anchors),
            valid_associations=len(positive),
            competing_associations=competing + max(0, len(positive) - 1),
            relation_type=relation_type,
            evidence_class=evidence_class,
        )
    alternative_values = _distinct_relation_values(structure, alternatives)
    if len(alternative_values) > 1:
        alternative = max(alternatives, key=lambda item: item.score)
        relation_type, evidence_class = _relation_summary(alternative)
        return _diagnostic(
            "review",
            "explicit_invoice_number_alternatives_compete",
            value_occurrences=len(matches),
            anchor_clusters=len(anchors),
            competing_associations=len(alternative_values),
            relation_type=relation_type,
            evidence_class=evidence_class,
        )
    if alternatives:
        alternative = max(alternatives, key=lambda item: item.score)
        relation_type, evidence_class = _relation_summary(alternative)
        return _diagnostic(
            "contradiction",
            "explicit_invoice_number_differs",
            value_occurrences=len(matches),
            anchor_clusters=len(anchors),
            valid_associations=0,
            competing_associations=len(alternatives),
            relation_type=relation_type,
            evidence_class=evidence_class,
        )
    return _diagnostic(
        "review",
        "invoice_number_evidence_missing" if not matches else "invoice_number_anchor_missing",
        value_occurrences=len(matches),
        anchor_clusters=len(anchors),
        competing_associations=len(negative),
        relation_type="negative_context" if negative else "none",
        evidence_class=(max(negative, key=lambda item: item.score).evidence_class if negative else "none"),
    )


def _candidate_date_values(
    candidate: FieldCandidate, *, spanish_supplier: bool
) -> tuple[str, ...]:
    values = candidate.possible_values or (candidate.normalized_value,)
    if spanish_supplier and len(values) > 1:
        # date_interpretations always emits DD/MM before MM/DD.
        return values[:1]
    return values


def resolve_invoice_date(
    structure: StructuralDocument,
    proposed_value: Any,
    *,
    spanish_supplier: bool,
) -> dict[str, Any]:
    try:
        target = date.fromisoformat(str(proposed_value or "")).isoformat()
    except ValueError:
        return _diagnostic("review", "invoice_date_missing_or_invalid")
    anchors = list(structure.anchors_by_type.get("invoice_date", ()))
    date_candidates = list(structure.candidates_by_type.get("date", ()))
    matches = [
        candidate
        for candidate in date_candidates
        if target in _candidate_date_values(candidate, spanish_supplier=spanish_supplier)
    ]
    ambiguous_matches = [
        candidate
        for candidate in matches
        if len(candidate.possible_values) > 1 and not spanish_supplier
    ]
    positive = _relations(
        structure,
        matches,
        {"invoice_date"},
        competing_anchor_types=_INVOICE_DATE_NEGATIVE,
    )
    negative = _relations(
        structure,
        matches,
        _INVOICE_DATE_NEGATIVE,
        competing_anchor_types={"invoice_date"},
    )
    matching_ids = {candidate.candidate_id for candidate in matches}
    candidate_map = _candidate_map(structure)
    alternatives = []
    for anchor in anchors:
        for relation in structure.relations_by_anchor.get(anchor.anchor_id, ()):
            candidate = candidate_map.get(relation.candidate_id)
            if (
                relation.evidence_class in _STRONG_EVIDENCE
                and candidate is not None
                and candidate.candidate_type == "date"
                and candidate.candidate_id not in matching_ids
                and not _relation_shadowed_by_local_candidate(
                    structure, relation, "date"
                )
                and not _relation_shadowed_by_competing_context(
                    structure, relation, _INVOICE_DATE_NEGATIVE
                )
            ):
                alternatives.append(relation)
    best, dominant = _dominant_relation(positive)
    relation_type, evidence_class = _relation_summary(best)

    if ambiguous_matches and positive:
        return _diagnostic(
            "review",
            "invoice_date_format_ambiguous",
            value_occurrences=len(matches),
            anchor_clusters=len(anchors),
            valid_associations=len(positive),
            competing_associations=len(negative),
            relation_type=relation_type,
            evidence_class=evidence_class,
        )
    if best is not None and dominant and not alternatives and not negative:
        return _diagnostic(
            "confirmed",
            "unique_or_dominant_invoice_date_association",
            value_occurrences=len(matches),
            anchor_clusters=len(anchors),
            valid_associations=len(positive),
            relation_type=relation_type,
            evidence_class=evidence_class,
        )
    if positive:
        return _diagnostic(
            "review",
            "invoice_date_associations_compete",
            value_occurrences=len(matches),
            anchor_clusters=len(anchors),
            valid_associations=len(positive),
            competing_associations=len(alternatives) + len(negative) + max(0, len(positive) - 1),
            relation_type=relation_type,
            evidence_class=evidence_class,
        )
    alternative_values = _distinct_relation_values(structure, alternatives)
    if len(alternative_values) > 1:
        alternative = max(alternatives, key=lambda item: item.score)
        relation_type, evidence_class = _relation_summary(alternative)
        return _diagnostic(
            "review",
            "explicit_invoice_date_alternatives_compete",
            value_occurrences=len(matches),
            anchor_clusters=len(anchors),
            competing_associations=len(alternative_values),
            relation_type=relation_type,
            evidence_class=evidence_class,
        )
    if alternatives:
        alternative = max(alternatives, key=lambda item: item.score)
        relation_type, evidence_class = _relation_summary(alternative)
        return _diagnostic(
            "contradiction",
            "explicit_invoice_date_differs",
            value_occurrences=len(matches),
            anchor_clusters=len(anchors),
            competing_associations=len(alternatives),
            relation_type=relation_type,
            evidence_class=evidence_class,
        )
    return _diagnostic(
        "review",
        "invoice_date_evidence_missing" if not matches else "invoice_date_anchor_missing",
        value_occurrences=len(matches),
        anchor_clusters=len(anchors),
        competing_associations=len(negative),
        relation_type="negative_context" if negative else "none",
        evidence_class=(max(negative, key=lambda item: item.score).evidence_class if negative else "none"),
    )


def _resolve_document_type(structure: StructuralDocument) -> dict[str, Any]:
    invoice_anchors = list(structure.anchors_by_type.get("document_type_invoice", ()))
    negative = list(structure.anchors_by_type.get("document_type_non_invoice", ()))
    if invoice_anchors and not negative:
        return _diagnostic(
            "confirmed",
            "invoice_document_anchor_present",
            anchor_clusters=len(invoice_anchors),
            valid_associations=len(invoice_anchors),
            relation_type="anchor_cluster",
            evidence_class="strong_label",
        )
    if negative and not invoice_anchors:
        return _diagnostic(
            "contradiction",
            "non_invoice_document_anchor_present",
            anchor_clusters=len(negative),
            competing_associations=len(negative),
            relation_type="anchor_cluster",
            evidence_class="strong_label",
        )
    return _diagnostic(
        "review",
        "document_type_ambiguous" if invoice_anchors or negative else "document_type_missing",
        anchor_clusters=len(invoice_anchors) + len(negative),
        competing_associations=min(len(invoice_anchors), len(negative)),
    )


def _role_relations(
    structure: StructuralDocument,
    candidates: list[FieldCandidate],
    role: str,
) -> list[EvidenceRelation]:
    return _relations(structure, candidates, {role})


def _resolve_tax_identity(
    structure: StructuralDocument,
    proposed_supplier_tax_id: Any,
    proposed_recipient_tax_id: Any,
    registered_company_tax_id: Any,
) -> tuple[dict[str, Any], dict[str, Any], str | None]:
    tax_candidates = list(structure.candidates_by_type.get("tax_id", ()))
    supplier_target = normalize_tax_id(proposed_supplier_tax_id)
    registered_target = normalize_tax_id(registered_company_tax_id)
    recipient_target = registered_target or normalize_tax_id(proposed_recipient_tax_id)
    supplier_matches = [
        candidate for candidate in tax_candidates
        if candidate.normalized_value == supplier_target
    ] if supplier_target else []
    recipient_matches = [
        candidate for candidate in tax_candidates
        if candidate.normalized_value == recipient_target
    ] if recipient_target else []
    supplier_role = _role_relations(structure, supplier_matches, "supplier")
    supplier_as_recipient = _role_relations(structure, supplier_matches, "recipient")
    recipient_role = _role_relations(structure, recipient_matches, "recipient")
    distinct_values = {candidate.normalized_value for candidate in tax_candidates}

    if supplier_target and recipient_target and supplier_target == recipient_target:
        supplier_result = _diagnostic(
            "contradiction",
            "supplier_tax_id_matches_registered_recipient",
            value_occurrences=len(supplier_matches),
            anchor_clusters=len(structure.anchors_by_type.get("supplier", ())),
            competing_associations=len(supplier_as_recipient),
            relation_type="recipient_identity",
            evidence_class="known_recipient",
        )
    elif supplier_role:
        best = max(supplier_role, key=lambda item: item.score)
        relation_type, evidence_class = _relation_summary(best)
        supplier_result = _diagnostic(
            "confirmed",
            "supplier_tax_id_in_supplier_region",
            value_occurrences=len(supplier_matches),
            anchor_clusters=len(structure.anchors_by_type.get("supplier", ())),
            valid_associations=len(supplier_role),
            competing_associations=len(supplier_as_recipient),
            relation_type=relation_type,
            evidence_class=evidence_class,
        )
    elif (
        supplier_matches
        and recipient_matches
        and supplier_target != recipient_target
        and distinct_values.issubset({supplier_target, recipient_target})
    ):
        supplier_result = _diagnostic(
            "confirmed",
            "supplier_tax_id_resolved_against_known_recipient",
            value_occurrences=len(supplier_matches),
            valid_associations=1,
            relation_type="identity_exclusion",
            evidence_class="known_recipient",
        )
    elif supplier_as_recipient:
        supplier_result = _diagnostic(
            "contradiction",
            "supplier_tax_id_in_recipient_region",
            value_occurrences=len(supplier_matches),
            competing_associations=len(supplier_as_recipient),
            relation_type="recipient_identity",
            evidence_class=max(supplier_as_recipient, key=lambda item: item.score).evidence_class,
        )
    else:
        supplier_result = _diagnostic(
            "review",
            "supplier_tax_id_unattributed" if supplier_matches else "supplier_tax_id_not_found",
            value_occurrences=len(supplier_matches),
            anchor_clusters=len(structure.anchors_by_type.get("supplier", ())),
        )

    if recipient_matches and (recipient_role or recipient_target == registered_target):
        recipient_result = _diagnostic(
            "confirmed",
            "recipient_tax_id_confirmed",
            value_occurrences=len(recipient_matches),
            anchor_clusters=len(structure.anchors_by_type.get("recipient", ())),
            valid_associations=max(len(recipient_role), 1),
            relation_type="recipient_identity",
            evidence_class="known_recipient" if recipient_target == registered_target else "strong_label",
        )
    else:
        recipient_result = _diagnostic(
            "review",
            "recipient_tax_id_not_found" if not recipient_matches else "recipient_tax_id_unattributed",
            value_occurrences=len(recipient_matches),
            anchor_clusters=len(structure.anchors_by_type.get("recipient", ())),
        )

    supplier_jurisdiction = (
        "ES"
        if supplier_result["status"] == "confirmed" and is_spanish_tax_id(supplier_target)
        else None
    )
    return supplier_result, recipient_result, supplier_jurisdiction


def _resolve_provider_name(
    structure: StructuralDocument,
    provider_name: Any,
    supplier_tax_result: dict[str, Any],
) -> dict[str, Any]:
    target = normalize_entity(provider_name)
    candidates = [
        candidate
        for candidate in structure.candidates_by_type.get("entity", ())
        if candidate.normalized_value == target
    ] if target else []
    supplier_relations = _role_relations(structure, candidates, "supplier")
    recipient_relations = _role_relations(structure, candidates, "recipient")
    if supplier_relations and not recipient_relations:
        best = max(supplier_relations, key=lambda item: item.score)
        relation_type, evidence_class = _relation_summary(best)
        return _diagnostic(
            "confirmed",
            "provider_name_in_supplier_region",
            value_occurrences=len(candidates),
            anchor_clusters=len(structure.anchors_by_type.get("supplier", ())),
            valid_associations=len(supplier_relations),
            relation_type=relation_type,
            evidence_class=evidence_class,
        )
    if recipient_relations and not supplier_relations:
        return _diagnostic(
            "contradiction",
            "provider_name_in_recipient_region",
            value_occurrences=len(candidates),
            competing_associations=len(recipient_relations),
            relation_type="recipient_identity",
            evidence_class=max(recipient_relations, key=lambda item: item.score).evidence_class,
        )
    if candidates and supplier_tax_result.get("status") == "confirmed":
        return _diagnostic(
            "review",
            "provider_name_present_but_role_unproven",
            value_occurrences=len(candidates),
            anchor_clusters=len(structure.anchors_by_type.get("supplier", ())),
        )
    return _diagnostic(
        "review",
        "provider_name_not_found" if not candidates else "provider_name_role_ambiguous",
        value_occurrences=len(candidates),
        competing_associations=len(recipient_relations),
    )


def resolve_identity(
    structure: StructuralDocument,
    normalized: dict[str, Any],
    *,
    registered_company_tax_id: Any,
) -> tuple[dict[str, dict[str, Any]], str | None]:
    supplier_tax, recipient_tax, jurisdiction = _resolve_tax_identity(
        structure,
        normalized.get("supplier_tax_id"),
        normalized.get("customer_tax_id"),
        registered_company_tax_id,
    )
    fields = {
        "document_type": _resolve_document_type(structure),
        "supplier_tax_id": supplier_tax,
        "recipient_tax_id": recipient_tax,
        "provider_name": _resolve_provider_name(
            structure, normalized.get("provider_name"), supplier_tax
        ),
    }
    fields["invoice_number"] = resolve_invoice_number(
        structure, normalized.get("invoice_number")
    )
    fields["invoice_date"] = resolve_invoice_date(
        structure,
        normalized.get("invoice_date"),
        spanish_supplier=jurisdiction == "ES",
    )
    return fields, jurisdiction


def _decimal_string(value: Any) -> str | None:
    try:
        return format(
            Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP),
            ".2f",
        )
    except (InvalidOperation, TypeError, ValueError):
        return None


def _resolve_money_field(
    structure: StructuralDocument, field: str, proposed_value: Any
) -> dict[str, Any]:
    target = _decimal_string(proposed_value)
    anchors = set(_MONEY_FIELD_ANCHORS[field])
    anchor_count = sum(len(structure.anchors_by_type.get(item, ())) for item in anchors)
    matches = [
        candidate
        for candidate in structure.candidates_by_type.get("money", ())
        if candidate.normalized_value == target
    ] if target is not None else []
    positive = _relations(structure, matches, anchors)
    best, dominant = _dominant_relation(positive)
    candidate_map = _candidate_map(structure)
    alternatives = []
    for anchor_type in anchors:
        for anchor in structure.anchors_by_type.get(anchor_type, ()):
            for relation in structure.relations_by_anchor.get(anchor.anchor_id, ()):
                candidate = candidate_map.get(relation.candidate_id)
                if (
                    relation.evidence_class in _STRONG_EVIDENCE
                    and candidate is not None
                    and candidate.candidate_type == "money"
                    and candidate.normalized_value != target
                ):
                    alternatives.append(relation)
    relation_type, evidence_class = _relation_summary(best)
    if best is not None and dominant and not alternatives:
        return _diagnostic(
            "confirmed",
            f"{field}_documented",
            value_occurrences=len(matches),
            anchor_clusters=anchor_count,
            valid_associations=len(positive),
            relation_type=relation_type,
            evidence_class=evidence_class,
        )
    if positive:
        return _diagnostic(
            "review",
            f"{field}_associations_compete",
            value_occurrences=len(matches),
            anchor_clusters=anchor_count,
            valid_associations=len(positive),
            competing_associations=len(alternatives) + max(0, len(positive) - 1),
            relation_type=relation_type,
            evidence_class=evidence_class,
        )
    if alternatives:
        return _diagnostic(
            "contradiction",
            f"explicit_{field}_differs",
            value_occurrences=len(matches),
            anchor_clusters=anchor_count,
            competing_associations=len(alternatives),
            relation_type="incompatible_anchor_value",
            evidence_class=max(alternatives, key=lambda item: item.score).evidence_class,
        )
    return _diagnostic(
        "review",
        f"{field}_evidence_missing",
        value_occurrences=len(matches),
        anchor_clusters=anchor_count,
    )


def _resolve_currency(
    structure: StructuralDocument, proposed_value: Any
) -> dict[str, Any]:
    target = str(proposed_value or "").strip().upper()
    currencies = list(structure.candidates_by_type.get("currency", ()))
    matches = [candidate for candidate in currencies if candidate.normalized_value == target]
    distinct = {candidate.normalized_value for candidate in currencies}
    if matches and distinct == {target}:
        return _diagnostic(
            "confirmed",
            "currency_consistent_in_document",
            value_occurrences=len(matches),
            valid_associations=len(matches),
            relation_type="document_currency",
            evidence_class="unique_document_value",
        )
    if matches:
        return _diagnostic(
            "review",
            "multiple_document_currencies",
            value_occurrences=len(matches),
            valid_associations=len(matches),
            competing_associations=len(distinct - {target}),
        )
    if currencies:
        return _diagnostic(
            "contradiction",
            "document_currency_differs",
            competing_associations=len(currencies),
            evidence_class="explicit_currency",
        )
    return _diagnostic("review", "currency_evidence_missing")


def _normalized_rate(value: Any) -> str | None:
    try:
        return format(Decimal(str(value)).normalize(), "f")
    except (InvalidOperation, TypeError, ValueError):
        return None


def _resolve_vat_breakdown(
    structure: StructuralDocument, proposed_lines: Any
) -> dict[str, Any]:
    lines = [line for line in proposed_lines or [] if isinstance(line, dict)]
    anchor_count = len(structure.anchors_by_type.get("vat_breakdown", ())) + len(
        structure.anchors_by_type.get("vat_amount", ())
    )
    used_candidates = set()
    valid_rows = 0
    ambiguous_rows = 0
    for line in lines:
        rate = _normalized_rate(line.get("rate"))
        base = _decimal_string(line.get("base"))
        vat = _decimal_string(line.get("vat_amount"))
        if rate is None or base is None or vat is None:
            ambiguous_rows += 1
            continue
        rate_candidates = [
            candidate
            for candidate in structure.candidates_by_type.get("percentage", ())
            if candidate.normalized_value == rate
        ]
        matched = False
        for rate_candidate in rate_candidates:
            base_candidates = [
                candidate
                for candidate in structure.candidates_by_type.get("money", ())
                if candidate.segment_id == rate_candidate.segment_id
                and candidate.normalized_value == base
                and candidate.candidate_id not in used_candidates
            ]
            vat_candidates = [
                candidate
                for candidate in structure.candidates_by_type.get("money", ())
                if candidate.segment_id == rate_candidate.segment_id
                and candidate.normalized_value == vat
                and candidate.candidate_id not in used_candidates
            ]
            pair = next(
                (
                    (base_candidate, vat_candidate)
                    for base_candidate in base_candidates
                    for vat_candidate in vat_candidates
                    if base_candidate.candidate_id != vat_candidate.candidate_id
                ),
                None,
            )
            if pair is None:
                continue
            used_candidates.update(
                {
                    rate_candidate.candidate_id,
                    pair[0].candidate_id,
                    pair[1].candidate_id,
                }
            )
            valid_rows += 1
            matched = True
            break
        if not matched:
            ambiguous_rows += 1
    if lines and valid_rows == len(lines):
        return _diagnostic(
            "confirmed",
            "vat_breakdown_rows_documented",
            value_occurrences=valid_rows,
            anchor_clusters=anchor_count,
            valid_associations=valid_rows,
            relation_type="tax_rows",
            evidence_class="distinct_row_candidates",
        )
    if valid_rows:
        return _diagnostic(
            "review",
            "vat_breakdown_partially_documented",
            value_occurrences=valid_rows,
            anchor_clusters=anchor_count,
            valid_associations=valid_rows,
            competing_associations=ambiguous_rows,
            relation_type="tax_rows",
            evidence_class="partial_rows",
        )
    return _diagnostic(
        "review",
        "vat_breakdown_evidence_missing",
        anchor_clusters=anchor_count,
        competing_associations=ambiguous_rows,
    )


def resolve_fiscal(
    structure: StructuralDocument, normalized: dict[str, Any]
) -> dict[str, dict[str, Any]]:
    fields = {
        "currency": _resolve_currency(structure, normalized.get("currency")),
        "base_amount": _resolve_money_field(
            structure, "base_amount", normalized.get("base_amount")
        ),
        "vat_amount": _resolve_money_field(
            structure, "vat_amount", normalized.get("vat_amount")
        ),
        "withholding_amount": _resolve_money_field(
            structure, "withholding_amount", normalized.get("withholding_amount")
        ),
        "other_taxes": _resolve_money_field(
            structure, "other_taxes", normalized.get("other_taxes")
        ),
        "total_amount": _resolve_money_field(
            structure, "total_amount", normalized.get("total_amount")
        ),
        "vat_breakdown": _resolve_vat_breakdown(
            structure, normalized.get("vat_breakdown")
        ),
    }
    return fields


def _to_decimal(value: Any) -> Decimal | None:
    try:
        return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, TypeError, ValueError):
        return None


def resolve_consistency(
    structure: StructuralDocument, normalized: dict[str, Any]
) -> dict[str, dict[str, Any]]:
    base = _to_decimal(normalized.get("base_amount"))
    vat = _to_decimal(normalized.get("vat_amount"))
    withholding = _to_decimal(normalized.get("withholding_amount")) or Decimal("0.00")
    other_taxes = _to_decimal(normalized.get("other_taxes")) or Decimal("0.00")
    total = _to_decimal(normalized.get("total_amount"))
    if base is None or vat is None or total is None:
        accounting = _diagnostic("review", "accounting_values_missing")
    else:
        expected = (base + vat + other_taxes - abs(withholding)).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        accounting = _diagnostic(
            "confirmed" if abs(expected - total) <= Decimal("0.01") else "contradiction",
            "accounting_equation_consistent" if abs(expected - total) <= Decimal("0.01") else "accounting_equation_mismatch",
            valid_associations=1 if abs(expected - total) <= Decimal("0.01") else 0,
            competing_associations=0 if abs(expected - total) <= Decimal("0.01") else 1,
            relation_type="arithmetic",
            evidence_class="consistency_only",
        )

    tax_lines = [line for line in normalized.get("vat_breakdown") or [] if isinstance(line, dict)]
    line_vat = [_to_decimal(line.get("vat_amount")) for line in tax_lines]
    if vat is None or not tax_lines or any(value is None for value in line_vat):
        vat_total = _diagnostic("review", "vat_breakdown_totals_missing")
    else:
        breakdown_total = sum(line_vat, Decimal("0.00"))
        matches = abs(breakdown_total - vat) <= Decimal("0.01")
        vat_total = _diagnostic(
            "confirmed" if matches else "contradiction",
            "vat_breakdown_total_consistent" if matches else "vat_breakdown_total_mismatch",
            value_occurrences=len(tax_lines),
            valid_associations=len(tax_lines) if matches else 0,
            competing_associations=0 if matches else 1,
            relation_type="arithmetic",
            evidence_class="consistency_only",
        )

    currency = str(normalized.get("currency") or "").upper()
    currencies = {
        candidate.normalized_value
        for candidate in structure.candidates_by_type.get("currency", ())
    }
    if not currencies:
        currency_consistency = _diagnostic("review", "currency_markers_missing")
    elif currencies == {currency}:
        currency_consistency = _diagnostic(
            "confirmed",
            "currency_markers_consistent",
            value_occurrences=sum(
                1
                for candidate in structure.candidates_by_type.get("currency", ())
                if candidate.normalized_value == currency
            ),
            valid_associations=1,
            relation_type="document_consistency",
            evidence_class="consistency_only",
        )
    else:
        currency_consistency = _diagnostic(
            "contradiction",
            "currency_markers_conflict",
            competing_associations=len(currencies - {currency}),
            relation_type="document_consistency",
            evidence_class="consistency_only",
        )
    return {
        "accounting_equation": accounting,
        "vat_breakdown_total": vat_total,
        "currency_consistency": currency_consistency,
    }


def compare_v2_with_legacy(
    canonical_result: dict[str, Any], legacy: dict[str, Any]
) -> dict[str, dict[str, Any]]:
    comparisons = {}
    for field in ("invoice_number", "invoice_date"):
        legacy_field = legacy.get(field) if isinstance(legacy.get(field), dict) else {}
        canonical_field = (canonical_result.get("fields") or {}).get(field) or {}
        legacy_status = str(legacy_field.get("status") or "unknown")
        canonical_status = str(canonical_field.get("status") or "unknown")
        comparisons[field] = {
            "legacy_status": legacy_status,
            "legacy_reason": str(legacy_field.get("reason") or "") or None,
            "canonical_status": canonical_status,
            "canonical_reason": str(canonical_field.get("reason_code") or "") or None,
            "transition": (
                "same_status"
                if legacy_status == canonical_status
                else f"legacy_{legacy_status}_to_canonical_{canonical_status}"
            ),
        }
    return comparisons


def verify_canonical_document_v2(
    layout: DocumentLayout,
    normalized: dict[str, Any],
    *,
    registered_company_tax_id: Any = None,
    registered_company_name: Any = None,
) -> dict[str, Any]:
    structure_started = time.monotonic()
    segments, token_by_id = build_semantic_segments(layout)
    anchors = build_anchor_clusters(segments, token_by_id)
    structure_ms = round((time.monotonic() - structure_started) * 1000, 3)

    candidates_started = time.monotonic()
    candidates = extract_field_candidates(segments, anchors, token_by_id)
    regions, segment_region = build_regions(
        segments,
        anchors,
        candidates,
        registered_company_tax_id=registered_company_tax_id,
        registered_company_name=registered_company_name,
    )
    candidate_extraction_ms = round(
        (time.monotonic() - candidates_started) * 1000, 3
    )

    graph_started = time.monotonic()
    relations = build_evidence_relations(
        segments, anchors, candidates, segment_region
    )
    structure = assemble_structural_document(
        segments,
        anchors,
        regions,
        candidates,
        relations,
        token_by_id,
        segment_region,
    )
    evidence_graph_ms = round((time.monotonic() - graph_started) * 1000, 3)

    verification_started = time.monotonic()
    identity, supplier_jurisdiction = resolve_identity(
        structure,
        normalized,
        registered_company_tax_id=registered_company_tax_id,
    )
    fiscal = resolve_fiscal(structure, normalized)
    consistency = resolve_consistency(structure, normalized)
    verification_ms = round((time.monotonic() - verification_started) * 1000, 3)

    candidate_counts = {
        candidate_type: len(values)
        for candidate_type, values in structure.candidates_by_type.items()
    }
    region_counts = {}
    for region in structure.regions:
        region_counts[region.region_type] = region_counts.get(region.region_type, 0) + 1
    return {
        "version": VERIFIER_VERSION,
        "fields": {**identity, **fiscal},
        "consistency": consistency,
        "supplier_jurisdiction": supplier_jurisdiction,
        "structure": {
            "segments": len(structure.segments),
            "anchor_clusters": len(structure.anchors),
            "regions": len(structure.regions),
            "candidates": len(structure.candidates),
            "relations": len(structure.relations),
            "candidate_counts": candidate_counts,
            "region_counts": region_counts,
        },
        "timings_ms": {
            "structure_ms": structure_ms,
            "candidate_extraction_ms": candidate_extraction_ms,
            "evidence_graph_ms": evidence_graph_ms,
            "verification_ms": verification_ms,
        },
    }
