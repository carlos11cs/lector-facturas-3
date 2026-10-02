"""Ephemeral party and fiscal structures derived from canonical evidence."""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from itertools import permutations
from typing import Any, Iterable

from services.canonical_document_structure import (
    AnchorCluster,
    FieldCandidate,
    StructuralDocument,
    normalize_entity,
    normalize_tax_id,
    normalize_word,
)


_STRONG_RELATIONS = {"strong_same_row", "strong_below"}
_PARTY_ROLES = {"supplier", "recipient"}
_TOTAL_ANCHOR_FIELDS = {
    "base_amount": "base_amount",
    "vat_amount": "vat_amount",
    "withholding_amount": "withholding_amount",
    "other_taxes": "other_taxes",
    "subtotal": "subtotal",
    "total_amount": "total_amount",
    "amount_due": "amount_due",
}


@dataclass(frozen=True)
class PartyCluster:
    cluster_id: str
    page: int
    segment_ids: tuple[str, ...]
    entity_candidate_ids: tuple[str, ...]
    tax_id_candidate_ids: tuple[str, ...]
    anchor_ids: tuple[str, ...]
    role_hints: tuple[str, ...]
    known_recipient_match: bool
    role: str


@dataclass(frozen=True)
class FiscalRow:
    row_id: str
    page: int
    rate_candidate_id: str
    base_candidate_id: str
    tax_candidate_id: str


@dataclass(frozen=True)
class FiscalTable:
    status: str
    rows: tuple[FiscalRow, ...]
    ambiguous_rows: int
    header_count: int


@dataclass(frozen=True)
class TotalsBlock:
    block_id: str
    page: int
    anchor_ids: tuple[str, ...]
    field_candidates: tuple[tuple[str, str], ...]
    ambiguous_fields: tuple[str, ...]
    status: str


@dataclass(frozen=True)
class FiscalStructure:
    table: FiscalTable
    totals_blocks: tuple[TotalsBlock, ...]


def _candidate_map(structure: StructuralDocument) -> dict[str, FieldCandidate]:
    return {candidate.candidate_id: candidate for candidate in structure.candidates}


def _anchor_map(structure: StructuralDocument) -> dict[str, AnchorCluster]:
    return {anchor.anchor_id: anchor for anchor in structure.anchors}


def _segment_map(structure: StructuralDocument):
    return {segment.segment_id: segment for segment in structure.segments}


def _height(box) -> float:
    return max(float(box[3]) - float(box[1]), 1.0)


def _center_x(box) -> float:
    return (float(box[0]) + float(box[2])) / 2


def _horizontal_gap(first, second) -> float:
    if first[2] < second[0]:
        return float(second[0]) - float(first[2])
    if second[2] < first[0]:
        return float(first[0]) - float(second[2])
    return 0.0


def _vertical_gap(first, second) -> float:
    if first[3] < second[1]:
        return float(second[1]) - float(first[3])
    if second[3] < first[1]:
        return float(first[1]) - float(second[3])
    return 0.0


def _column_related(first, second) -> bool:
    overlap = min(first[2], second[2]) - max(first[0], second[0])
    scale = max(_height(first), _height(second))
    return overlap >= 0 or abs(_center_x(first) - _center_x(second)) <= scale * 6


def _candidate_near_cluster(
    candidate: FieldCandidate,
    members: list[FieldCandidate],
) -> float | None:
    scores = []
    for member in members:
        if member.page != candidate.page:
            continue
        scale = max(_height(member.bbox), _height(candidate.bbox))
        if member.row_id == candidate.row_id:
            gap = _horizontal_gap(member.bbox, candidate.bbox)
            if gap <= scale * 10:
                scores.append(100 - gap / scale)
        else:
            gap = _vertical_gap(member.bbox, candidate.bbox)
            if gap <= scale * 3 and _column_related(member.bbox, candidate.bbox):
                scores.append(80 - gap / scale)
    return max(scores) if scores else None


def _cluster_name_matches(
    structure: StructuralDocument,
    candidate_ids: Iterable[str],
    name: Any,
) -> bool:
    target = normalize_entity(name)
    if not target:
        return False
    candidates = _candidate_map(structure)
    entities = sorted(
        (
            candidates[candidate_id]
            for candidate_id in candidate_ids
            if candidate_id in candidates
        ),
        key=lambda item: (item.page, item.bbox[1], item.bbox[0]),
    )
    values = [candidate.normalized_value for candidate in entities]
    if target in values:
        return True
    if len(target) >= 6 and any(target in value for value in values):
        return True
    role_prefixes = {
        "PROVEEDOR",
        "EMISOR",
        "SUPPLIER",
        "ISSUER",
        "CLIENTE",
        "RECEPTOR",
        "DESTINATARIO",
        "CUSTOMER",
        "BILLTO",
    }
    for start in range(len(values)):
        combined = ""
        for end in range(start, min(len(values), start + 4)):
            combined += values[end]
            if combined == target:
                return True
            if target.endswith(combined) and target[: -len(combined)] in role_prefixes:
                return True
            if (
                len(combined) >= 6
                and combined in target
                and len(combined) / len(target) >= 0.6
            ):
                return True
            if len(combined) > len(target):
                break
    return False


def party_cluster_matches_name(
    structure: StructuralDocument, cluster: PartyCluster, name: Any
) -> bool:
    return _cluster_name_matches(structure, cluster.entity_candidate_ids, name)


def party_cluster_tax_values(
    structure: StructuralDocument, cluster: PartyCluster
) -> tuple[str, ...]:
    candidates = _candidate_map(structure)
    return tuple(
        dict.fromkeys(
            candidates[candidate_id].normalized_value
            for candidate_id in cluster.tax_id_candidate_ids
            if candidate_id in candidates
        )
    )


def build_party_clusters(
    structure: StructuralDocument,
    *,
    registered_company_tax_id: Any = None,
    registered_company_name: Any = None,
) -> tuple[PartyCluster, ...]:
    candidates = [
        candidate
        for candidate in structure.candidates
        if candidate.candidate_type in {"entity", "tax_id"}
    ]
    anchors = _anchor_map(structure)
    registered_tax = normalize_tax_id(registered_company_tax_id)
    role_scores: dict[str, dict[str, list[tuple[float, str]]]] = {}
    for candidate in candidates:
        for relation in structure.relations_by_candidate.get(candidate.candidate_id, ()):
            anchor = anchors.get(relation.anchor_id)
            if (
                anchor is None
                or anchor.anchor_type not in _PARTY_ROLES
                or relation.evidence_class not in _STRONG_RELATIONS
            ):
                continue
            role_scores.setdefault(candidate.candidate_id, {}).setdefault(
                anchor.anchor_type, []
            ).append((relation.score, anchor.anchor_id))

    buckets: dict[str, list[FieldCandidate]] = {}
    bucket_roles: dict[str, set[str]] = {}
    bucket_anchors: dict[str, set[str]] = {}
    known_recipient_ids = set()
    for candidate in candidates:
        known_recipient = bool(
            candidate.candidate_type == "tax_id"
            and registered_tax
            and candidate.normalized_value == registered_tax
        ) or bool(
            candidate.candidate_type == "entity"
            and registered_company_name
            and candidate.normalized_value == normalize_entity(registered_company_name)
        )
        scores = role_scores.get(candidate.candidate_id, {})
        best_by_role = {
            role: max(values, key=lambda item: item[0])
            for role, values in scores.items()
        }
        ordered = sorted(
            best_by_role.items(), key=lambda item: item[1][0], reverse=True
        )
        resolved_role = None
        if ordered and (
            len(ordered) == 1 or ordered[0][1][0] >= ordered[1][1][0] + 15
        ):
            resolved_role = ordered[0][0]
        if (
            not known_recipient
            and resolved_role == "recipient"
            and candidate.candidate_type == "entity"
            and _cluster_name_matches(
                structure, (candidate.candidate_id,), registered_company_name
            )
        ):
            known_recipient = True
        if known_recipient:
            known_recipient_ids.add(candidate.candidate_id)
            # A known recipient value without a role relation is attached in
            # the geometry phases below. Creating a separate bucket here
            # would split the recipient name from its tax identifier.
            if resolved_role is None:
                continue
            resolved_role = "recipient"
        if resolved_role is None:
            continue
        top_score = best_by_role.get(resolved_role, (None, None))[0]
        role_anchor_ids = sorted(
            anchor_id
            for score, anchor_id in scores.get(resolved_role, ())
            if top_score is not None and score >= top_score - 15
        )
        key = (
            f"role:{resolved_role}:{candidate.page}:{','.join(role_anchor_ids)}"
            if role_anchor_ids
            else f"known:{resolved_role}:{candidate.page}"
        )
        bucket_roles.setdefault(key, set()).add(resolved_role)
        bucket_anchors.setdefault(key, set()).update(role_anchor_ids)
        buckets.setdefault(key, []).append(candidate)

    assigned = {
        candidate.candidate_id for values in buckets.values() for candidate in values
    }
    unassigned = [candidate for candidate in candidates if candidate.candidate_id not in assigned]

    def attach_unique(
        pending: list[FieldCandidate],
        *,
        candidate_type: str,
        require_entity_seed: bool,
    ) -> list[FieldCandidate]:
        # Each phase scores against a fixed snapshot. This permits a wrapped
        # party name followed by its tax ID without allowing arbitrary
        # transitive growth through the rest of the page.
        seeds = {key: tuple(members) for key, members in buckets.items()}
        attachments: dict[str, list[FieldCandidate]] = {}
        for candidate in pending:
            if candidate.candidate_type != candidate_type:
                continue
            scored = []
            for key, members in seeds.items():
                role = next(iter(bucket_roles.get(key, ())), None)
                if candidate.candidate_id in known_recipient_ids and role != "recipient":
                    continue
                entity_members = [
                    member for member in members if member.candidate_type == "entity"
                ]
                if require_entity_seed and not entity_members:
                    continue
                score = _candidate_near_cluster(candidate, list(members))
                if score is None:
                    continue
                if candidate_type == "entity" and not any(
                    member.row_id == candidate.row_id
                    or (
                        member.page == candidate.page
                        and _vertical_gap(member.bbox, candidate.bbox)
                        <= max(_height(member.bbox), _height(candidate.bbox)) * 2
                        and _column_related(member.bbox, candidate.bbox)
                    )
                    for member in entity_members
                ):
                    continue
                scored.append((score, key))
            scored.sort(reverse=True)
            if scored and (len(scored) == 1 or scored[0][0] >= scored[1][0] + 15):
                attachments.setdefault(scored[0][1], []).append(candidate)
        attached_ids = {
            candidate.candidate_id
            for values in attachments.values()
            for candidate in values
        }
        for key, values in attachments.items():
            buckets[key].extend(values)
        return [
            candidate
            for candidate in pending
            if candidate.candidate_id not in attached_ids
        ]

    unassigned = attach_unique(
        unassigned,
        candidate_type="entity",
        require_entity_seed=True,
    )
    unassigned = attach_unique(
        unassigned,
        candidate_type="tax_id",
        require_entity_seed=True,
    )

    for candidate in sorted(unassigned, key=lambda item: (item.page, item.bbox[1], item.bbox[0])):
        compatible = [
            (score, key)
            for key, members in buckets.items()
            if key.startswith("unassigned:")
            and any(
                member.candidate_type != candidate.candidate_type
                or member.row_id == candidate.row_id
                for member in members
            )
            and (score := _candidate_near_cluster(candidate, members)) is not None
        ]
        compatible.sort(reverse=True)
        key = compatible[0][1] if compatible else f"unassigned:{len(buckets) + 1}"
        buckets.setdefault(key, []).append(candidate)

    clusters = []
    for index, (key, members) in enumerate(
        sorted(
            buckets.items(),
            key=lambda item: (
                min(candidate.page for candidate in item[1]),
                min(candidate.bbox[1] for candidate in item[1]),
                min(candidate.bbox[0] for candidate in item[1]),
            ),
        ),
        1,
    ):
        roles = set(bucket_roles.get(key, set()))
        for member in members:
            roles.update(role_scores.get(member.candidate_id, {}).keys())
        known_recipient = any(
            member.candidate_id in known_recipient_ids for member in members
        )
        if known_recipient:
            role = "recipient"
        elif roles == {"supplier"}:
            role = "supplier"
        elif roles == {"recipient"}:
            role = "recipient"
        else:
            role = "ambiguous"
        related_anchor_ids = {
            relation.anchor_id
            for member in members
            for relation in structure.relations_by_candidate.get(member.candidate_id, ())
            if relation.anchor_id in anchors
            and anchors[relation.anchor_id].anchor_type in _PARTY_ROLES
            and relation.evidence_class in _STRONG_RELATIONS
        }
        clusters.append(
            PartyCluster(
                cluster_id=f"party-{index}",
                page=members[0].page,
                segment_ids=tuple(dict.fromkeys(member.segment_id for member in members)),
                entity_candidate_ids=tuple(
                    member.candidate_id
                    for member in members
                    if member.candidate_type == "entity"
                ),
                tax_id_candidate_ids=tuple(
                    member.candidate_id
                    for member in members
                    if member.candidate_type == "tax_id"
                ),
                anchor_ids=tuple(sorted(related_anchor_ids | bucket_anchors.get(key, set()))),
                role_hints=tuple(sorted(roles)),
                known_recipient_match=known_recipient,
                role=role,
            )
        )
    return tuple(clusters)


def _header_columns(structure: StructuralDocument):
    rows: dict[str, list] = {}
    for segment in structure.segments:
        rows.setdefault(segment.row_id, []).extend(
            structure.token_by_id[token_id]
            for token_id in segment.token_ids
            if token_id in structure.token_by_id
        )
    row_records = []
    for row_id, tokens in rows.items():
        ordered = sorted(tokens, key=lambda token: token.bbox[0])
        if not ordered:
            continue
        row_records.append(
            {
                "row_id": row_id,
                "page": ordered[0].page,
                "y": min(token.bbox[1] for token in ordered),
                "bottom": max(token.bbox[3] for token in ordered),
                "height": max(_height(token.bbox) for token in ordered),
                "tokens": ordered,
            }
        )
    row_records.sort(key=lambda row: (row["page"], row["y"]))

    def cues(tokens):
        terms = [normalize_word(token.original_text) for token in tokens]
        base = [
            _center_x(token.bbox)
            for token, term in zip(tokens, terms)
            if term in {"BASE", "TAXABLE"}
        ]
        rate = [
            _center_x(token.bbox)
            for token, term in zip(tokens, terms)
            if term in {"TIPO", "RATE", "PORCENTAJE", "PERCENTAGE"}
            or token.original_text.strip().upper() in {"%", "%IVA", "IVA%", "%VAT", "VAT%"}
        ]
        tax = [
            _center_x(token.bbox)
            for token, term in zip(tokens, terms)
            if term in {"CUOTA", "TAX"}
        ]
        if not tax and any(term in {"IVA", "VAT"} for term in terms):
            tax = [
                _center_x(token.bbox)
                for token, term in zip(tokens, terms)
                if term in {"IMPORTE", "AMOUNT"}
            ]
        return base, rate, tax

    headers = []
    seen = set()
    for index, first in enumerate(row_records):
        band = [first]
        for following in row_records[index + 1 : index + 3]:
            if following["page"] != first["page"]:
                break
            if following["y"] - band[-1]["bottom"] > max(
                following["height"], band[-1]["height"]
            ) * 2.5:
                break
            band.append(following)
        for width in range(1, len(band) + 1):
            selected_rows = band[:width]
            base = []
            rate = []
            tax = []
            for row in selected_rows:
                row_base, row_rate, row_tax = cues(row["tokens"])
                base.extend(row_base)
                rate.extend(row_rate)
                tax.extend(row_tax)
            if not (base and rate and tax):
                continue
            columns = (rate[0], base[0], tax[-1])
            key = (
                first["page"],
                round(min(row["y"] for row in selected_rows), 1),
                tuple(round(value, 1) for value in columns),
            )
            if key in seen:
                break
            seen.add(key)
            headers.append(
                {
                    "row_ids": tuple(row["row_id"] for row in selected_rows),
                    "page": first["page"],
                    "y": min(row["y"] for row in selected_rows),
                    "bottom": max(row["bottom"] for row in selected_rows),
                    "height": max(row["height"] for row in selected_rows),
                    "columns": columns,
                }
            )
            break
    return headers


def build_fiscal_table(structure: StructuralDocument) -> FiscalTable:
    headers = _header_columns(structure)
    rates_by_row: dict[str, list[FieldCandidate]] = {}
    money_by_row: dict[str, list[FieldCandidate]] = {}
    for candidate in structure.candidates_by_type.get("percentage", ()):
        rates_by_row.setdefault(candidate.row_id, []).append(candidate)
    for candidate in structure.candidates_by_type.get("money", ()):
        money_by_row.setdefault(candidate.row_id, []).append(candidate)
    segment_by_id = _segment_map(structure)
    rows = []
    ambiguous = 0
    used_candidates = set()
    for header in headers:
        candidate_row_ids = sorted(
            {
                candidate.row_id
                for candidate in structure.candidates
                if candidate.page == header["page"]
                and candidate.row_id not in header["row_ids"]
                and candidate.bbox[1] >= header["bottom"] - 1
                and candidate.bbox[1] - header["bottom"] <= header["height"] * 20
            },
            key=lambda row_id: min(
                segment.bbox[1]
                for segment in structure.segments
                if segment.row_id == row_id
            ),
        )
        for row_id in candidate_row_ids:
            rates = [
                candidate
                for candidate in rates_by_row.get(row_id, ())
                if candidate.candidate_id not in used_candidates
            ]
            money = [
                candidate
                for candidate in money_by_row.get(row_id, ())
                if candidate.candidate_id not in used_candidates
            ]
            rate_token_ids = {
                token_id for candidate in rates for token_id in candidate.token_ids
            }
            money = [
                candidate
                for candidate in money
                if rate_token_ids.isdisjoint(candidate.token_ids)
            ]
            if not rates and not money:
                continue
            rate_x, base_x, tax_x = header["columns"]
            if not rates and len(money) >= 3:
                inferred_rates = []
                for candidate in money:
                    try:
                        numeric = Decimal(candidate.normalized_value)
                    except (InvalidOperation, ValueError):
                        continue
                    if Decimal("0") <= numeric <= Decimal("100"):
                        inferred_rates.append(candidate)
                inferred_rates.sort(
                    key=lambda candidate: abs(_center_x(candidate.bbox) - rate_x)
                )
                if inferred_rates:
                    if (
                        len(inferred_rates) == 1
                        or abs(_center_x(inferred_rates[0].bbox) - rate_x) + 5
                        < abs(_center_x(inferred_rates[1].bbox) - rate_x)
                    ):
                        rates = [inferred_rates[0]]
                        money = [
                            candidate
                            for candidate in money
                            if candidate.candidate_id != inferred_rates[0].candidate_id
                        ]
            if len(rates) != 1 or len(money) < 2:
                if rates:
                    ambiguous += 1
                continue
            ranked = sorted(
                [
                    (
                    abs(_center_x(base.bbox) - base_x)
                    + abs(_center_x(tax.bbox) - tax_x),
                    base,
                    tax,
                    )
                for base, tax in permutations(money, 2)
                if base.candidate_id != tax.candidate_id
                ],
                key=lambda item: item[0],
            )
            if not ranked:
                ambiguous += 1
                continue
            if len(ranked) > 1 and abs(ranked[0][0] - ranked[1][0]) < 1:
                ambiguous += 1
                continue
            _, base, tax = ranked[0]
            rate = rates[0]
            used_candidates.update(
                {rate.candidate_id, base.candidate_id, tax.candidate_id}
            )
            rows.append(
                FiscalRow(
                    row_id=row_id,
                    page=segment_by_id[rate.segment_id].page,
                    rate_candidate_id=rate.candidate_id,
                    base_candidate_id=base.candidate_id,
                    tax_candidate_id=tax.candidate_id,
                )
            )
    return FiscalTable(
        status="confirmed" if rows and ambiguous == 0 else "review",
        rows=tuple(rows),
        ambiguous_rows=ambiguous,
        header_count=len(headers),
    )


def _anchor_rows(structure: StructuralDocument, anchor: AnchorCluster) -> set[str]:
    segments = _segment_map(structure)
    return {
        segments[segment_id].row_id
        for segment_id in anchor.segment_ids
        if segment_id in segments
    }


def _totals_anchor_groups(structure: StructuralDocument):
    table_rows = {
        row_id
        for anchor in structure.anchors_by_type.get("vat_breakdown", ())
        for row_id in _anchor_rows(structure, anchor)
    }
    anchors = [
        anchor
        for anchor in structure.anchors
        if anchor.anchor_type in _TOTAL_ANCHOR_FIELDS
        and not _anchor_rows(structure, anchor).intersection(table_rows)
    ]
    groups: list[list[AnchorCluster]] = []
    for anchor in sorted(anchors, key=lambda item: (item.page, item.bbox[1], item.bbox[0])):
        compatible = []
        for index, group in enumerate(groups):
            if group[0].page != anchor.page:
                continue
            if any(
                _vertical_gap(existing.bbox, anchor.bbox)
                <= max(_height(existing.bbox), _height(anchor.bbox)) * 6
                and (
                    abs(existing.bbox[1] - anchor.bbox[1])
                    <= max(_height(existing.bbox), _height(anchor.bbox))
                    or _column_related(existing.bbox, anchor.bbox)
                )
                for existing in group
            ):
                compatible.append(index)
        if len(compatible) == 1:
            groups[compatible[0]].append(anchor)
        else:
            groups.append([anchor])
    return groups


def build_totals_blocks(structure: StructuralDocument) -> tuple[TotalsBlock, ...]:
    candidates = _candidate_map(structure)
    blocks = []
    for index, anchors in enumerate(_totals_anchor_groups(structure), 1):
        relations = []
        for anchor in anchors:
            for relation in structure.relations_by_anchor.get(anchor.anchor_id, ()):
                candidate = candidates.get(relation.candidate_id)
                if (
                    candidate is None
                    or candidate.candidate_type != "money"
                    or relation.evidence_class not in _STRONG_RELATIONS
                ):
                    continue
                scale = max(_height(anchor.bbox), _height(candidate.bbox))
                center_distance = abs(_center_x(anchor.bbox) - _center_x(candidate.bbox))
                if relation.evidence_class == "strong_same_row":
                    score = 220 - relation.horizontal_distance / scale
                else:
                    score = 150 - center_distance / scale * 8 - relation.vertical_distance / scale
                relations.append((score, anchor, candidate))

        by_candidate: dict[str, list] = {}
        for item in relations:
            by_candidate.setdefault(item[2].candidate_id, []).append(item)
        assigned = []
        ambiguous_fields = set()
        for candidate_relations in by_candidate.values():
            candidate_relations.sort(key=lambda item: item[0], reverse=True)
            if (
                len(candidate_relations) > 1
                and candidate_relations[0][0] < candidate_relations[1][0] + 10
                and candidate_relations[0][1].anchor_type
                != candidate_relations[1][1].anchor_type
            ):
                ambiguous_fields.update(
                    _TOTAL_ANCHOR_FIELDS[item[1].anchor_type]
                    for item in candidate_relations[:2]
                )
                continue
            assigned.append(candidate_relations[0])

        by_field: dict[str, list] = {}
        for score, anchor, candidate in assigned:
            field = _TOTAL_ANCHOR_FIELDS[anchor.anchor_type]
            by_field.setdefault(field, []).append((score, candidate))
        field_candidates = []
        for field, values in by_field.items():
            values.sort(key=lambda item: item[0], reverse=True)
            distinct = {item[1].normalized_value for item in values}
            if len(distinct) > 1 and len(values) > 1 and values[0][0] < values[1][0] + 10:
                ambiguous_fields.add(field)
                continue
            field_candidates.append((field, values[0][1].candidate_id))
        resolved_fields = {field for field, _ in field_candidates}
        complete = (
            "total_amount" in resolved_fields
            and {"base_amount", "vat_amount"}.issubset(resolved_fields)
            and not ambiguous_fields
        )
        blocks.append(
            TotalsBlock(
                block_id=f"totals-{index}",
                page=anchors[0].page,
                anchor_ids=tuple(anchor.anchor_id for anchor in anchors),
                field_candidates=tuple(field_candidates),
                ambiguous_fields=tuple(sorted(ambiguous_fields)),
                status="confirmed" if complete else "review",
            )
        )
    return tuple(blocks)


def build_fiscal_structure(structure: StructuralDocument) -> FiscalStructure:
    return FiscalStructure(
        table=build_fiscal_table(structure),
        totals_blocks=build_totals_blocks(structure),
    )
