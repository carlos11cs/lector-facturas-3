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
from services.canonical_document_semantics import (
    FiscalStructure,
    PartyCluster,
    build_fiscal_structure,
    build_party_clusters,
    party_cluster_matches_name,
    party_cluster_tax_values,
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


def resolve_identity(
    structure: StructuralDocument,
    normalized: dict[str, Any],
    *,
    registered_company_tax_id: Any,
    registered_company_name: Any = None,
    party_clusters: tuple[PartyCluster, ...] | None = None,
) -> tuple[dict[str, dict[str, Any]], str | None, dict[str, Any]]:
    clusters = party_clusters or build_party_clusters(
        structure,
        registered_company_tax_id=registered_company_tax_id,
        registered_company_name=registered_company_name,
    )
    supplier_target = normalize_tax_id(normalized.get("supplier_tax_id"))
    recipient_target = normalize_tax_id(
        registered_company_tax_id or normalized.get("customer_tax_id")
    )
    registered_name_target = normalize_entity(registered_company_name)
    provider_name_target = normalize_entity(normalized.get("provider_name"))

    known_recipients = [cluster for cluster in clusters if cluster.known_recipient_match]
    labelled_recipients = [cluster for cluster in clusters if cluster.role == "recipient"]
    recipient_cluster = (
        known_recipients[0]
        if len(known_recipients) == 1
        else labelled_recipients[0]
        if not known_recipients and len(labelled_recipients) == 1
        else None
    )

    def matches_supplier_proposal(cluster: PartyCluster) -> bool:
        return bool(
            supplier_target
            and supplier_target in party_cluster_tax_values(structure, cluster)
        ) or party_cluster_matches_name(
            structure, cluster, normalized.get("provider_name")
        )

    supplier_pool = [
        cluster
        for cluster in clusters
        if cluster is not recipient_cluster and cluster.role == "supplier"
    ]
    if len(supplier_pool) > 1:
        matching = [cluster for cluster in supplier_pool if matches_supplier_proposal(cluster)]
        supplier_pool = matching if len(matching) == 1 else supplier_pool
    if len(supplier_pool) == 1:
        supplier_cluster = supplier_pool[0]
        supplier_resolution_reason = "explicit_supplier_cluster"
    else:
        exclusion_matches = [
            cluster
            for cluster in clusters
            if cluster is not recipient_cluster and matches_supplier_proposal(cluster)
        ]
        if recipient_cluster is not None and len(exclusion_matches) == 1:
            supplier_cluster = exclusion_matches[0]
            supplier_resolution_reason = "supplier_resolved_against_known_recipient"
        else:
            supplier_cluster = None
            supplier_resolution_reason = "supplier_cluster_ambiguous"

    supplier_values = (
        party_cluster_tax_values(structure, supplier_cluster)
        if supplier_cluster is not None
        else ()
    )
    recipient_values = (
        party_cluster_tax_values(structure, recipient_cluster)
        if recipient_cluster is not None
        else ()
    )
    all_tax_candidates = list(structure.candidates_by_type.get("tax_id", ()))
    supplier_occurrences = sum(
        candidate.normalized_value == supplier_target for candidate in all_tax_candidates
    ) if supplier_target else 0
    recipient_occurrences = sum(
        candidate.normalized_value == recipient_target for candidate in all_tax_candidates
    ) if recipient_target else 0

    if supplier_target and recipient_target and supplier_target == recipient_target:
        supplier_tax = _diagnostic(
            "contradiction",
            "supplier_tax_id_matches_registered_recipient",
            value_occurrences=supplier_occurrences,
            relation_type="party_cluster",
            evidence_class="known_recipient",
        )
    elif supplier_target and supplier_target in supplier_values:
        supplier_tax = _diagnostic(
            "confirmed",
            "supplier_tax_id_in_resolved_supplier_cluster",
            value_occurrences=supplier_occurrences,
            valid_associations=1,
            relation_type="party_cluster",
            evidence_class=supplier_resolution_reason,
        )
    elif supplier_target and supplier_target in recipient_values:
        supplier_tax = _diagnostic(
            "review",
            "supplier_tax_id_in_mixed_recipient_cluster",
            value_occurrences=supplier_occurrences,
            competing_associations=1,
            relation_type="party_cluster",
            evidence_class="cluster_contamination",
        )
    elif supplier_cluster is not None and supplier_values and supplier_target:
        supplier_tax = _diagnostic(
            "contradiction" if len(set(supplier_values)) == 1 else "review",
            "resolved_supplier_tax_id_differs" if len(set(supplier_values)) == 1 else "supplier_cluster_has_multiple_tax_ids",
            value_occurrences=supplier_occurrences,
            competing_associations=len(set(supplier_values)),
            relation_type="party_cluster",
            evidence_class=supplier_resolution_reason,
        )
    else:
        supplier_tax = _diagnostic(
            "review",
            "supplier_cluster_unresolved" if supplier_cluster is None else "supplier_tax_id_not_in_cluster",
            value_occurrences=supplier_occurrences,
            competing_associations=max(len(supplier_pool), 0),
            relation_type="party_cluster" if supplier_cluster else "none",
        )

    if recipient_target and recipient_target in recipient_values:
        recipient_tax = _diagnostic(
            "confirmed",
            "recipient_tax_id_in_resolved_recipient_cluster",
            value_occurrences=recipient_occurrences,
            valid_associations=1,
            relation_type="party_cluster",
            evidence_class="known_recipient" if recipient_cluster and recipient_cluster.known_recipient_match else "strong_label",
        )
    elif recipient_cluster is not None and recipient_values and recipient_target:
        recipient_tax = _diagnostic(
            "contradiction" if len(set(recipient_values)) == 1 else "review",
            "resolved_recipient_tax_id_differs" if len(set(recipient_values)) == 1 else "recipient_cluster_has_multiple_tax_ids",
            value_occurrences=recipient_occurrences,
            competing_associations=len(set(recipient_values)),
            relation_type="party_cluster",
        )
    else:
        recipient_tax = _diagnostic(
            "review",
            "recipient_cluster_unresolved" if recipient_cluster is None else "recipient_tax_id_not_in_cluster",
            value_occurrences=recipient_occurrences,
            competing_associations=max(len(known_recipients) - 1, 0),
        )

    if supplier_cluster is not None and party_cluster_matches_name(
        structure, supplier_cluster, normalized.get("provider_name")
    ):
        provider_name = _diagnostic(
            "confirmed",
            "provider_name_in_resolved_supplier_cluster",
            value_occurrences=1,
            valid_associations=1,
            relation_type="party_cluster",
            evidence_class=supplier_resolution_reason,
        )
    elif recipient_cluster is not None and party_cluster_matches_name(
        structure, recipient_cluster, normalized.get("provider_name")
    ):
        provider_name = _diagnostic(
            "contradiction"
            if provider_name_target
            and registered_name_target
            and provider_name_target == registered_name_target
            else "review",
            "provider_name_matches_registered_recipient"
            if provider_name_target
            and registered_name_target
            and provider_name_target == registered_name_target
            else "provider_name_in_mixed_recipient_cluster",
            value_occurrences=1,
            competing_associations=1,
            relation_type="party_cluster",
            evidence_class="known_recipient"
            if provider_name_target == registered_name_target
            else "cluster_contamination",
        )
    else:
        provider_name = _diagnostic(
            "review",
            "provider_name_not_in_resolved_supplier_cluster" if supplier_cluster else "supplier_cluster_unresolved",
            value_occurrences=0,
            competing_associations=max(len(supplier_pool) - 1, 0),
            relation_type="party_cluster" if supplier_cluster else "none",
        )

    jurisdiction = (
        "ES"
        if supplier_tax["status"] == "confirmed" and is_spanish_tax_id(supplier_target)
        else None
    )
    fields = {
        "document_type": _resolve_document_type(structure),
        "supplier_tax_id": supplier_tax,
        "recipient_tax_id": recipient_tax,
        "provider_name": provider_name,
    }
    fields["invoice_number"] = resolve_invoice_number(
        structure, normalized.get("invoice_number")
    )
    fields["invoice_date"] = resolve_invoice_date(
        structure,
        normalized.get("invoice_date"),
        spanish_supplier=jurisdiction == "ES",
    )
    aggregate = {
        "supplier_status": "confirmed" if supplier_cluster is not None else "review",
        "recipient_status": "confirmed" if recipient_cluster is not None else "review",
        "supplier_tax_id_status": supplier_tax["status"],
        "jurisdiction_status": "confirmed" if jurisdiction else "review",
        "party_cluster_count": len(clusters),
        "ambiguous_party_clusters": sum(cluster.role == "ambiguous" for cluster in clusters),
    }
    return fields, jurisdiction, aggregate


def _decimal_string(value: Any) -> str | None:
    try:
        return format(
            Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP),
            ".2f",
        )
    except (InvalidOperation, TypeError, ValueError):
        return None


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


def _fiscal_candidate_value(
    structure: StructuralDocument, candidate_id: str
) -> str | None:
    candidate = _candidate_map(structure).get(candidate_id)
    return candidate.normalized_value if candidate is not None else None


def _totals_field_values(
    structure: StructuralDocument,
    fiscal_structure: FiscalStructure,
    field: str,
) -> list[str]:
    values = []
    for block in fiscal_structure.totals_blocks:
        if field in block.ambiguous_fields:
            continue
        for candidate_field, candidate_id in block.field_candidates:
            if candidate_field != field:
                continue
            value = _fiscal_candidate_value(structure, candidate_id)
            if value is not None:
                values.append(value)
    return values


def _fiscal_table_rows(
    structure: StructuralDocument, fiscal_structure: FiscalStructure
) -> list[tuple[str, str, str]]:
    rows = []
    for row in fiscal_structure.table.rows:
        rate = _fiscal_candidate_value(structure, row.rate_candidate_id)
        base = _fiscal_candidate_value(structure, row.base_candidate_id)
        tax = _fiscal_candidate_value(structure, row.tax_candidate_id)
        normalized_rate = _normalized_rate(rate)
        if normalized_rate is not None and base is not None and tax is not None:
            rows.append((normalized_rate, base, tax))
    return rows


def _sum_decimal_strings(values: Iterable[str]) -> str | None:
    try:
        return format(sum((Decimal(value) for value in values), Decimal("0.00")), ".2f")
    except (InvalidOperation, TypeError, ValueError):
        return None


def _structured_money_values(
    structure: StructuralDocument,
    fiscal_structure: FiscalStructure,
    field: str,
) -> list[str]:
    values = _totals_field_values(structure, fiscal_structure, field)
    if fiscal_structure.table.status == "confirmed" and field in {"base_amount", "vat_amount"}:
        table_rows = _fiscal_table_rows(structure, fiscal_structure)
        position = 1 if field == "base_amount" else 2
        total = _sum_decimal_strings(row[position] for row in table_rows)
        if total is not None:
            values.append(total)
    return values


def _definitive_structured_money_values(
    structure: StructuralDocument,
    fiscal_structure: FiscalStructure,
    field: str,
) -> list[str]:
    values = []
    for block in fiscal_structure.totals_blocks:
        if block.status != "confirmed" or field in block.ambiguous_fields:
            continue
        for candidate_field, candidate_id in block.field_candidates:
            if candidate_field == field:
                value = _fiscal_candidate_value(structure, candidate_id)
                if value is not None:
                    values.append(value)
    if fiscal_structure.table.status == "confirmed" and field in {
        "base_amount",
        "vat_amount",
    }:
        position = 1 if field == "base_amount" else 2
        total = _sum_decimal_strings(
            row[position] for row in _fiscal_table_rows(structure, fiscal_structure)
        )
        if total is not None:
            values.append(total)
    return values


def _resolve_structured_money_field(
    structure: StructuralDocument,
    fiscal_structure: FiscalStructure,
    field: str,
    proposed_value: Any,
) -> dict[str, Any]:
    target = _decimal_string(proposed_value)
    values = _structured_money_values(structure, fiscal_structure, field)
    distinct = set(values)
    definitive = set(
        _definitive_structured_money_values(structure, fiscal_structure, field)
    )
    anchor_count = len(structure.anchors_by_type.get(field, ()))
    if target is not None and distinct == {target}:
        return _diagnostic(
            "confirmed",
            f"{field}_in_reconstructed_fiscal_structure",
            value_occurrences=len(values),
            anchor_clusters=anchor_count,
            valid_associations=len(values),
            relation_type="fiscal_structure",
            evidence_class="unique_structural_value",
        )
    if target is not None and len(distinct) == 1 and distinct == definitive:
        return _diagnostic(
            "contradiction",
            f"reconstructed_{field}_differs",
            value_occurrences=0,
            anchor_clusters=anchor_count,
            competing_associations=1,
            relation_type="fiscal_structure",
            evidence_class="unique_structural_value",
        )
    if target is not None and len(distinct) == 1:
        return _diagnostic(
            "review",
            f"provisional_{field}_differs",
            value_occurrences=0,
            anchor_clusters=anchor_count,
            competing_associations=1,
            relation_type="fiscal_structure",
            evidence_class="provisional_structural_value",
        )
    if len(distinct) > 1:
        return _diagnostic(
            "review",
            f"reconstructed_{field}_values_compete",
            value_occurrences=sum(value == target for value in values),
            anchor_clusters=anchor_count,
            valid_associations=sum(value == target for value in values),
            competing_associations=len(distinct),
            relation_type="fiscal_structure",
            evidence_class="competing_structural_values",
        )
    return _diagnostic(
        "review",
        f"{field}_structure_missing",
        anchor_clusters=anchor_count,
        relation_type="fiscal_structure",
    )


def _resolve_structured_vat_breakdown(
    structure: StructuralDocument,
    fiscal_structure: FiscalStructure,
    proposed_lines: Any,
) -> dict[str, Any]:
    proposed = []
    invalid = 0
    for line in proposed_lines or ():
        if not isinstance(line, dict):
            invalid += 1
            continue
        rate = _normalized_rate(line.get("rate"))
        base = _decimal_string(line.get("base"))
        tax = _decimal_string(line.get("vat_amount"))
        if rate is None or base is None or tax is None:
            invalid += 1
            continue
        proposed.append((rate, base, tax))
    documented = _fiscal_table_rows(structure, fiscal_structure)
    anchor_count = fiscal_structure.table.header_count
    if (
        fiscal_structure.table.status == "confirmed"
        and invalid == 0
        and proposed
        and sorted(proposed) == sorted(documented)
    ):
        return _diagnostic(
            "confirmed",
            "vat_breakdown_matches_reconstructed_table",
            value_occurrences=len(documented),
            anchor_clusters=anchor_count,
            valid_associations=len(documented),
            relation_type="fiscal_table_rows",
            evidence_class="exclusive_row_assignment",
        )
    if fiscal_structure.table.status == "confirmed" and documented and proposed:
        return _diagnostic(
            "contradiction",
            "reconstructed_vat_breakdown_differs",
            value_occurrences=0,
            anchor_clusters=anchor_count,
            competing_associations=max(len(documented), len(proposed)),
            relation_type="fiscal_table_rows",
            evidence_class="exclusive_row_assignment",
        )
    return _diagnostic(
        "review",
        "vat_breakdown_structure_ambiguous" if documented else "vat_breakdown_structure_missing",
        value_occurrences=len(documented),
        anchor_clusters=anchor_count,
        valid_associations=len(documented),
        competing_associations=fiscal_structure.table.ambiguous_rows + invalid,
        relation_type="fiscal_table_rows" if documented else "none",
        evidence_class="partial_row_assignment" if documented else "none",
    )


def _has_sufficient_optional_tax_coverage(
    structure: StructuralDocument, fiscal_structure: FiscalStructure
) -> bool:
    confirmed_blocks = [
        block for block in fiscal_structure.totals_blocks if block.status == "confirmed"
    ]
    if len(confirmed_blocks) == 1:
        return True
    return bool(
        fiscal_structure.table.status == "confirmed"
        and len(set(_totals_field_values(structure, fiscal_structure, "total_amount"))) == 1
    )


def _zero_optional_tax_is_consistent(
    normalized: dict[str, Any], field: str
) -> bool:
    base = _to_decimal(normalized.get("base_amount"))
    vat = _to_decimal(normalized.get("vat_amount"))
    total = _to_decimal(normalized.get("total_amount"))
    if base is None or vat is None or total is None:
        return False
    withholding = (
        Decimal("0.00")
        if field == "withholding_amount"
        else abs(_to_decimal(normalized.get("withholding_amount")) or Decimal("0.00"))
    )
    other = (
        Decimal("0.00")
        if field == "other_taxes"
        else _to_decimal(normalized.get("other_taxes")) or Decimal("0.00")
    )
    expected = (base + vat + other - withholding).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )
    return abs(expected - total) <= Decimal("0.01")


def _resolve_optional_tax(
    structure: StructuralDocument,
    fiscal_structure: FiscalStructure,
    normalized: dict[str, Any],
    field: str,
) -> dict[str, Any]:
    proposed = _to_decimal(normalized.get(field)) or Decimal("0.00")
    raw_values = _totals_field_values(structure, fiscal_structure, field)
    if field == "withholding_amount":
        explicit = {
            format(abs(Decimal(value)), ".2f") for value in raw_values
        }
        target = format(abs(proposed), ".2f")
    else:
        explicit = set(raw_values)
        target = format(proposed, ".2f")
    anchor_count = len(structure.anchors_by_type.get(field, ()))
    if explicit == {target}:
        return _diagnostic(
            "confirmed",
            f"{field}_explicitly_documented",
            value_occurrences=len(raw_values),
            anchor_clusters=anchor_count,
            valid_associations=len(raw_values),
            relation_type="totals_block",
            evidence_class="explicit_optional_tax",
        )
    if len(explicit) == 1:
        return _diagnostic(
            "contradiction",
            f"explicit_{field}_differs",
            anchor_clusters=anchor_count,
            competing_associations=1,
            relation_type="totals_block",
            evidence_class="explicit_optional_tax",
        )
    if len(explicit) > 1:
        return _diagnostic(
            "review",
            f"explicit_{field}_values_compete",
            anchor_clusters=anchor_count,
            competing_associations=len(explicit),
            relation_type="totals_block",
            evidence_class="competing_structural_values",
        )
    if (
        proposed == Decimal("0.00")
        and _has_sufficient_optional_tax_coverage(structure, fiscal_structure)
        and _zero_optional_tax_is_consistent(normalized, field)
    ):
        return _diagnostic(
            "not_applicable",
            f"{field}_absent_with_sufficient_fiscal_coverage",
            anchor_clusters=anchor_count,
            valid_associations=1,
            relation_type="fiscal_structure_absence",
            evidence_class="coverage_and_consistency",
        )
    return _diagnostic(
        "review",
        f"{field}_absence_not_proven" if proposed == Decimal("0.00") else f"{field}_structure_missing",
        anchor_clusters=anchor_count,
        relation_type="fiscal_structure",
    )


def resolve_fiscal(
    structure: StructuralDocument,
    normalized: dict[str, Any],
    *,
    fiscal_structure: FiscalStructure | None = None,
) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    reconstructed = fiscal_structure or build_fiscal_structure(structure)
    fields = {
        "currency": _resolve_currency(structure, normalized.get("currency")),
        "base_amount": _resolve_structured_money_field(
            structure, reconstructed, "base_amount", normalized.get("base_amount")
        ),
        "vat_amount": _resolve_structured_money_field(
            structure, reconstructed, "vat_amount", normalized.get("vat_amount")
        ),
        "withholding_amount": _resolve_optional_tax(
            structure, reconstructed, normalized, "withholding_amount"
        ),
        "other_taxes": _resolve_optional_tax(
            structure, reconstructed, normalized, "other_taxes"
        ),
        "total_amount": _resolve_structured_money_field(
            structure, reconstructed, "total_amount", normalized.get("total_amount")
        ),
        "vat_breakdown": _resolve_structured_vat_breakdown(
            structure, reconstructed, normalized.get("vat_breakdown")
        ),
    }
    confirmed_blocks = sum(
        block.status == "confirmed" for block in reconstructed.totals_blocks
    )
    totals_ambiguities = sum(
        len(block.ambiguous_fields) for block in reconstructed.totals_blocks
    )
    totals_field_count = sum(
        len(block.field_candidates) for block in reconstructed.totals_blocks
    )
    aggregate = {
        "fiscal_table_status": reconstructed.table.status,
        "fiscal_row_count": len(reconstructed.table.rows),
        "totals_block_status": "confirmed" if confirmed_blocks == 1 and totals_ambiguities == 0 else "review",
        "totals_field_count": totals_field_count,
        "ambiguous_fiscal_rows": reconstructed.table.ambiguous_rows + totals_ambiguities,
    }
    return fields, aggregate


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
    party_clusters = build_party_clusters(
        structure,
        registered_company_tax_id=registered_company_tax_id,
        registered_company_name=registered_company_name,
    )
    reconstructed_fiscal = build_fiscal_structure(structure)
    identity, supplier_jurisdiction, party_resolution = resolve_identity(
        structure,
        normalized,
        registered_company_tax_id=registered_company_tax_id,
        registered_company_name=registered_company_name,
        party_clusters=party_clusters,
    )
    fiscal, fiscal_structure = resolve_fiscal(
        structure,
        normalized,
        fiscal_structure=reconstructed_fiscal,
    )
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
        "party_resolution": party_resolution,
        "fiscal_structure": fiscal_structure,
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
