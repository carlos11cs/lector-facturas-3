#!/usr/bin/env python3
"""Read-only benchmark report for persisted Invoice Engine V2 shadow runs.

The script deliberately queries only operational/accounting fields already
stored by the shadow benchmark and V1 telemetry. It never calls OpenAI, S3 or
the invoice analysis code, and it does not print document text, prompts,
filenames, PDFs or result payloads.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import sys
from collections import Counter, defaultdict
from typing import Any, Iterable, Mapping, Sequence

from sqlalchemy import MetaData, Table, create_engine, func, select
from sqlalchemy.engine import Engine


_COMPARISON_FIELD_LABELS = {
    "provider_match": "provider",
    "invoice_number_match": "invoice_number",
    "invoice_date_match": "invoice_date",
    "tax_base_match": "base_amount",
    "withholding_match": "withholding_amount",
    "total_match": "total_amount",
    "due_date_match": "payment_dates",
}
_REVIEWABLE_V2_STATUSES = {"failed", "error"}
_PENDING_V2_STATUSES = {"queued", "processing", "retrying"}
_TOKEN_FIELDS = (
    ("input_tokens", "Input"),
    ("output_tokens", "Output"),
    ("reasoning_tokens", "Reasoning"),
    ("total_tokens", "Total"),
)


def _normalize_database_url(value: str) -> str:
    database_url = (value or "").strip()
    if database_url.startswith("postgres://"):
        return database_url.replace("postgres://", "postgresql://", 1)
    return database_url


def _safe_json_dict(value: Any) -> dict[str, Any]:
    if isinstance(value, dict):
        return value
    if not isinstance(value, str) or not value.strip():
        return {}
    try:
        parsed = json.loads(value)
    except (TypeError, ValueError):
        return {}
    return parsed if isinstance(parsed, dict) else {}


def _safe_json_list(value: Any) -> list[Any]:
    if isinstance(value, list):
        return value
    if not isinstance(value, str) or not value.strip():
        return []
    try:
        parsed = json.loads(value)
    except (TypeError, ValueError):
        return []
    return parsed if isinstance(parsed, list) else []


def _first_value(*values: Any) -> Any:
    for value in values:
        if value is not None and value != "":
            return value
    return None


def _canonical_v1_result(value: Any) -> dict[str, Any]:
    """Read only the persisted normalized V1 fields needed for VAT diagnosis."""
    result = _safe_json_dict(value)
    structured = _safe_json_dict(result.get("structured_extraction"))
    totals = _safe_json_dict(structured.get("totals"))
    structured_taxes = structured.get("taxes")
    if not isinstance(structured_taxes, list):
        structured_taxes = []
    structured_breakdown = [
        {
            "base": line.get("taxable_base"),
            "rate": line.get("vat_rate"),
            "vat_amount": line.get("vat_amount"),
        }
        for line in structured_taxes
        if isinstance(line, dict)
    ]
    return {
        "vat_amount": _first_value(result.get("vat_amount"), totals.get("vat_amount")),
        "vat_breakdown": (
            result.get("vat_breakdown")
            if result.get("vat_breakdown") is not None
            else structured_breakdown
        ),
    }


def _canonical_v2_result(value: Any) -> dict[str, Any]:
    result = _safe_json_dict(value)
    return {
        "vat_amount": result.get("vat_amount"),
        "vat_breakdown": result.get("vat_breakdown"),
    }


def _amount_matches(left: Any, right: Any, tolerance: float = 0.01) -> bool:
    if left is None or right is None:
        return left is None and right is None
    try:
        return abs(float(left) - float(right)) <= tolerance
    except (TypeError, ValueError):
        return False


def _vat_breakdown_matches(left: Any, right: Any) -> bool:
    def normalized_lines(value: Any) -> list[tuple[float, float, float]] | None:
        if value is None:
            value = []
        if not isinstance(value, list):
            return None
        lines = []
        for item in value:
            if not isinstance(item, dict):
                return None
            try:
                base = round(float(item.get("base")), 2)
                vat_amount = round(float(item.get("vat_amount")), 2)
            except (TypeError, ValueError):
                return None
            if base == 0 and vat_amount == 0:
                continue
            try:
                rate = round(float(item.get("rate")), 2)
            except (TypeError, ValueError):
                return None
            lines.append((rate, base, vat_amount))
        return sorted(lines)

    left_lines = normalized_lines(left)
    right_lines = normalized_lines(right)
    if left_lines is None or right_lines is None or len(left_lines) != len(right_lines):
        return False
    return all(
        _amount_matches(left_line[0], right_line[0])
        and _amount_matches(left_line[1], right_line[1])
        and _amount_matches(left_line[2], right_line[2])
        for left_line, right_line in zip(left_lines, right_lines)
    )


def _comparison_differences(row: Mapping[str, Any]) -> set[str]:
    """Map persisted comparison metadata to review fields without displaying JSON."""
    comparison = _safe_json_dict(row.get("comparison_json"))
    differences = {
        _COMPARISON_FIELD_LABELS[field]
        for field in comparison.get("different_fields", [])
        if field in _COMPARISON_FIELD_LABELS
    }
    if comparison.get("supplier_tax_id_match") is False:
        differences.add("supplier_tax_id")

    # The persisted historical comparator records VAT as one combined field.
    # Split it here from already persisted normalized results for a useful
    # diagnostic, while retaining the original comparator unchanged.
    if "vat_match" in comparison.get("different_fields", []):
        v1 = _canonical_v1_result(row.get("v1_result_json"))
        v2 = _canonical_v2_result(row.get("v2_result_json"))
        if not _amount_matches(v1.get("vat_amount"), v2.get("vat_amount")):
            differences.add("vat_amount")
        if not _vat_breakdown_matches(v1.get("vat_breakdown"), v2.get("vat_breakdown")):
            differences.add("vat_breakdown")
        if "vat_amount" not in differences and "vat_breakdown" not in differences:
            differences.add("vat_amount_or_breakdown")
    return differences


def _numeric_values(rows: Iterable[Mapping[str, Any]], field: str) -> list[float]:
    values = []
    for row in rows:
        value = row.get(field)
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            continue
        values.append(float(value))
    return values


def _paired_values(
    rows: Iterable[Mapping[str, Any]], v1_field: str, v2_field: str
) -> list[tuple[float, float]]:
    pairs = []
    for row in rows:
        left = row.get(v1_field)
        right = row.get(v2_field)
        if any(isinstance(value, bool) or not isinstance(value, (int, float)) for value in (left, right)):
            continue
        pairs.append((float(left), float(right)))
    return pairs


def _percentile(values: Sequence[float], percentile: float) -> float | None:
    if not values:
        return None
    ordered = sorted(values)
    if len(ordered) == 1:
        return ordered[0]
    rank = (len(ordered) - 1) * percentile
    lower = math.floor(rank)
    upper = math.ceil(rank)
    if lower == upper:
        return ordered[lower]
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (rank - lower)


def _summary(values: Sequence[float]) -> dict[str, float | int | None]:
    if not values:
        return {"count": 0, "mean": None, "p50": None, "p95": None}
    return {
        "count": len(values),
        "mean": sum(values) / len(values),
        "p50": _percentile(values, 0.50),
        "p95": _percentile(values, 0.95),
    }


def _percent(numerator: int | float, denominator: int | float) -> float | None:
    if not denominator:
        return None
    return (float(numerator) / float(denominator)) * 100


def _comparison_summary(pairs: Sequence[tuple[float, float]]) -> dict[str, float | int | None]:
    if not pairs:
        return {
            "count": 0,
            "v1_mean": None,
            "v2_mean": None,
            "absolute_reduction": None,
            "reduction_percent": None,
            "speedup": None,
        }
    v1_mean = sum(left for left, _ in pairs) / len(pairs)
    v2_mean = sum(right for _, right in pairs) / len(pairs)
    reduction = v1_mean - v2_mean
    return {
        "count": len(pairs),
        "v1_mean": v1_mean,
        "v2_mean": v2_mean,
        "absolute_reduction": reduction,
        "reduction_percent": _percent(reduction, v1_mean),
        "speedup": (v1_mean / v2_mean) if v2_mean > 0 else None,
    }


def build_benchmark_report(rows: Sequence[Mapping[str, Any]], scope: Mapping[str, Any]) -> dict[str, Any]:
    """Build all aggregate diagnostics from selected persisted shadow runs."""
    normalized_rows = [dict(row) for row in rows]
    total = len(normalized_rows)
    completed_rows = [row for row in normalized_rows if row.get("status") == "completed"]
    validation_passed_rows = [
        row for row in normalized_rows if row.get("validation_status") == "passed"
    ]
    validation_failed_rows = [
        row for row in normalized_rows if row.get("validation_status") == "failed"
    ]
    strict_match_rows = [row for row in normalized_rows if row.get("strict_accounting_match") is True]
    strict_mismatch_rows = [row for row in normalized_rows if row.get("strict_accounting_match") is False]
    comparable_rows = [
        row
        for row in normalized_rows
        if row.get("eligible") is True
        and row.get("status") == "completed"
        and row.get("strict_accounting_match") in (True, False)
    ]
    validation_passed_comparable_rows = [
        row for row in comparable_rows if row.get("validation_status") == "passed"
    ]
    strict_validation_passed_rows = [
        row
        for row in validation_passed_comparable_rows
        if row.get("strict_accounting_match") is True
    ]
    pending_rows = [
        row for row in normalized_rows if str(row.get("status") or "") in _PENDING_V2_STATUSES
    ]
    technical_error_rows = [
        row
        for row in normalized_rows
        if str(row.get("status") or "") in _REVIEWABLE_V2_STATUSES or row.get("error_type")
    ]
    safe_fast_path_rows = [
        row
        for row in normalized_rows
        if row.get("eligible") is True
        and row.get("status") == "completed"
        and row.get("validation_status") == "passed"
        and row.get("strict_accounting_match") is True
    ]
    accounting_safe_rows = [
        row
        for row in normalized_rows
        if row.get("eligible") is True
        and row.get("status") == "completed"
        and row.get("accounting_safety_status") == "passed"
    ]
    accounting_failed_rows = [
        row
        for row in normalized_rows
        if row.get("eligible") is True
        and row.get("status") == "completed"
        and row.get("accounting_safety_status") == "failed"
    ]
    accounting_evaluated_rows = accounting_safe_rows + accounting_failed_rows
    accounting_safe_comparable_rows = [
        row for row in comparable_rows if row.get("accounting_safety_status") == "passed"
    ]
    strict_accounting_safe_rows = [
        row
        for row in accounting_safe_comparable_rows
        if row.get("strict_accounting_match") is True
    ]
    metadata_confirmed_rows = [
        row
        for row in normalized_rows
        if row.get("eligible") is True
        and row.get("status") == "completed"
        and row.get("metadata_quality_status") == "confirmed"
    ]
    metadata_review_rows = [
        row
        for row in normalized_rows
        if row.get("eligible") is True
        and row.get("status") == "completed"
        and row.get("metadata_quality_status") == "review_required"
    ]
    metadata_failed_rows = [
        row
        for row in normalized_rows
        if row.get("eligible") is True
        and row.get("status") == "completed"
        and row.get("metadata_quality_status") == "failed"
    ]
    metadata_evaluated_rows = (
        metadata_confirmed_rows + metadata_review_rows + metadata_failed_rows
    )
    fully_confirmed_rows = [
        row
        for row in normalized_rows
        if row.get("eligible") is True
        and row.get("status") == "completed"
        and row.get("accounting_safety_status") == "passed"
        and row.get("metadata_quality_status") == "confirmed"
        and row.get("strict_accounting_match") is True
    ]
    eligible_completed_rows = [
        row
        for row in normalized_rows
        if row.get("eligible") is True and row.get("status") == "completed"
    ]
    fallback_accounting_rows = [
        row for row in eligible_completed_rows if row.get("accounting_safety_status") == "failed"
    ]
    fallback_metadata_rows = [
        row
        for row in eligible_completed_rows
        if row.get("accounting_safety_status") == "passed"
        and row.get("metadata_quality_status") != "confirmed"
    ]
    fallback_unassessed_rows = [
        row
        for row in eligible_completed_rows
        if row.get("accounting_safety_status") not in ("passed", "failed")
        or (
            row.get("accounting_safety_status") == "passed"
            and row.get("metadata_quality_status") == "confirmed"
            and row.get("strict_accounting_match") is not True
        )
    ]

    def discrepancies_for(rows_to_group: Iterable[Mapping[str, Any]]) -> dict[str, dict[str, Any]]:
        scoped_rows = list(rows_to_group)
        grouped: dict[str, set[int]] = defaultdict(set)
        for row in scoped_rows:
            job_id = row.get("job_id")
            if not isinstance(job_id, int):
                continue
            for field in _comparison_differences(row):
                grouped[field].add(job_id)
        denominator = len(scoped_rows)
        return {
            field: {
                "count": len(job_ids),
                "percent": _percent(len(job_ids), denominator),
                "job_ids": sorted(job_ids),
            }
            for field, job_ids in sorted(grouped.items())
        }

    validation_error_counts: Counter[str] = Counter()
    for row in normalized_rows:
        for code in _safe_json_list(row.get("validation_errors_json")):
            if isinstance(code, str) and code.strip():
                validation_error_counts[code.strip()] += 1

    ineligible_reasons = Counter(
        str(row.get("eligibility_reason") or "unknown")
        for row in normalized_rows
        if row.get("eligible") is False
    )
    # Failed, queued and skipped V2 runs are operationally important above, but
    # they are not comparable extraction latencies. Keep performance figures on
    # the same completed V2 corpus and state missing V1 telemetry explicitly.
    v1_latency = _summary(_numeric_values(completed_rows, "v1_processing_ms"))
    v2_latency = _summary(_numeric_values(completed_rows, "total_ms"))
    v1_openai = _summary(_numeric_values(completed_rows, "v1_openai_ms"))
    v2_openai = _summary(_numeric_values(completed_rows, "openai_ms"))

    token_report = {}
    for field, label in _TOKEN_FIELDS:
        v1_field = f"v1_{field}"
        token_report[field] = {
            "label": label,
            "v1": _summary(_numeric_values(completed_rows, v1_field)),
            "v2": _summary(_numeric_values(completed_rows, field)),
            "paired": _comparison_summary(_paired_values(completed_rows, v1_field, field)),
        }

    review_ids = {
        int(row["job_id"])
        for row in normalized_rows
        if isinstance(row.get("job_id"), int)
        and (
            row.get("validation_status") == "failed"
            or row.get("strict_accounting_match") is False
            or row.get("metadata_quality_status") in ("review_required", "failed")
            or row.get("accounting_safety_status") == "failed"
            or str(row.get("status") or "") in _REVIEWABLE_V2_STATUSES
            or bool(row.get("error_type"))
        )
    }
    review_rows = [
        {
            "job_id": row["job_id"],
            "status": row.get("status") or "unknown",
            "validation_status": row.get("validation_status") or "unknown",
            "strict_accounting_match": row.get("strict_accounting_match"),
            "accounting_safety_status": row.get("accounting_safety_status") or "n/a",
            "metadata_quality_status": row.get("metadata_quality_status") or "n/a",
        }
        for row in normalized_rows
        if row.get("job_id") in review_ids
    ]

    return {
        "scope": dict(scope),
        "volume": {
            "total": total,
            "eligible": sum(row.get("eligible") is True for row in normalized_rows),
            "ineligible": sum(row.get("eligible") is False for row in normalized_rows),
            "completed": len(completed_rows),
            "validation_passed": len(validation_passed_rows),
            "validation_failed": len(validation_failed_rows),
            "strict_match": len(strict_match_rows),
            "strict_mismatch": len(strict_mismatch_rows),
            "pending": len(pending_rows),
            "technical_error": len(technical_error_rows),
            "accounting_safety_passed": len(accounting_safe_rows),
            "accounting_safety_failed": len(accounting_failed_rows),
            "accounting_safety_unassessed": len(
                [
                    row
                    for row in normalized_rows
                    if row.get("eligible") is True
                    and row.get("status") == "completed"
                    and row.get("accounting_safety_status") not in ("passed", "failed")
                ]
            ),
            "metadata_confirmed": len(metadata_confirmed_rows),
            "metadata_review_required": len(metadata_review_rows),
            "metadata_failed": len(metadata_failed_rows),
            "metadata_unassessed": len(
                [
                    row
                    for row in normalized_rows
                    if row.get("eligible") is True
                    and row.get("status") == "completed"
                    and row.get("metadata_quality_status")
                    not in ("confirmed", "review_required", "failed")
                ]
            ),
            "fully_confirmed_fast_path_candidate": len(fully_confirmed_rows),
            "accounting_safe_metadata_review": len(metadata_review_rows),
            "comparable": len(comparable_rows),
            "fallback_accounting": len(fallback_accounting_rows),
            "fallback_metadata": len(fallback_metadata_rows),
            "fallback_unassessed_or_mismatch": len(fallback_unassessed_rows),
            "fallback_not_eligible": sum(row.get("eligible") is False for row in normalized_rows),
        },
        "rates": {
            "eligibility": _percent(
                sum(row.get("eligible") is True for row in normalized_rows), total
            ),
            "validation_passed_of_completed": _percent(
                len(validation_passed_rows), len(completed_rows)
            ),
            "strict_match_of_completed": _percent(len(strict_match_rows), len(completed_rows)),
            "strict_match_of_validation_passed": _percent(
                len(strict_validation_passed_rows), len(validation_passed_comparable_rows)
            ),
            "strict_match_of_accounting_safe": _percent(
                len(strict_accounting_safe_rows), len(accounting_safe_comparable_rows)
            ),
            "safe_fast_path_candidate": _percent(len(safe_fast_path_rows), total),
            "requires_v1_fallback": _percent(total - len(safe_fast_path_rows), total),
            "accounting_safety_passed_of_completed": _percent(
                len(accounting_safe_rows), len(accounting_evaluated_rows)
            ),
            "metadata_confirmed_of_evaluated": _percent(
                len(metadata_confirmed_rows), len(metadata_evaluated_rows)
            ),
            "metadata_confirmed_of_accounting_safe": _percent(
                len(
                    [
                        row
                        for row in accounting_safe_rows
                        if row.get("metadata_quality_status") == "confirmed"
                    ]
                ),
                len(accounting_safe_rows),
            ),
            "fully_confirmed_fast_path_candidate": _percent(
                len(fully_confirmed_rows), total
            ),
            "fully_confirmed_fast_path_candidate_of_eligible_completed": _percent(
                len(fully_confirmed_rows), len(eligible_completed_rows)
            ),
        },
        "safe_fast_path_candidate_count": len(safe_fast_path_rows),
        "fully_confirmed_fast_path_candidate_count": len(fully_confirmed_rows),
        "latency": {
            "v1_processing_ms": v1_latency,
            "v2_total_ms": v2_latency,
            "paired": _comparison_summary(
                _paired_values(completed_rows, "v1_processing_ms", "total_ms")
            ),
            "v1_openai_ms": v1_openai,
            "v2_openai_ms": v2_openai,
            "openai_paired": _comparison_summary(
                _paired_values(completed_rows, "v1_openai_ms", "openai_ms")
            ),
        },
        "tokens": token_report,
        "discrepancies": {
            "validation_failed": {
                "count": len(validation_failed_rows),
                "job_ids": sorted(
                    row["job_id"]
                    for row in validation_failed_rows
                    if isinstance(row.get("job_id"), int)
                ),
                "fields": discrepancies_for(validation_failed_rows),
            },
            "validation_passed_strict_mismatch": {
                "count": len(
                    [
                        row
                        for row in strict_mismatch_rows
                        if row.get("validation_status") == "passed"
                    ]
                ),
                "fields": discrepancies_for(
                    [
                        row
                        for row in strict_mismatch_rows
                        if row.get("validation_status") == "passed"
                    ]
                ),
            },
        },
        "validation_errors": dict(sorted(validation_error_counts.items())),
        "eligibility_reasons": dict(sorted(ineligible_reasons.items())),
        "review_rows": sorted(review_rows, key=lambda row: row["job_id"]),
    }


def _reflect_benchmark_tables(engine: Engine) -> tuple[Table, Table, Table]:
    metadata = MetaData()
    try:
        jobs = Table("invoice_analysis_jobs", metadata, autoload_with=engine)
        metrics = Table("invoice_analysis_metrics", metadata, autoload_with=engine)
        shadow_runs = Table("invoice_analysis_shadow_runs", metadata, autoload_with=engine)
    except Exception as error:
        raise RuntimeError(
            "No se han encontrado las tablas de análisis de facturas requeridas. "
            "Ejecuta el script en el Shell del servicio con DATABASE_URL configurada."
        ) from error
    return jobs, metrics, shadow_runs


def load_benchmark_rows(
    engine: Engine,
    *,
    shadow_version: str,
    batch_id: str | None = None,
    job_min: int | None = None,
    job_max: int | None = None,
    latest: int | None = None,
) -> list[dict[str, Any]]:
    """Load a scoped benchmark corpus without selecting filenames or source data."""
    jobs, metrics, shadow_runs = _reflect_benchmark_tables(engine)
    latest_metric_id = (
        select(func.max(metrics.c.id))
        .where(metrics.c.job_id == shadow_runs.c.job_id)
        .correlate(shadow_runs)
        .scalar_subquery()
    )
    query = (
        select(
            shadow_runs.c.id.label("shadow_run_id"),
            shadow_runs.c.job_id,
            shadow_runs.c.batch_id.label("shadow_batch_id"),
            jobs.c.batch_id.label("job_batch_id"),
            shadow_runs.c.eligible,
            shadow_runs.c.eligibility_reason,
            shadow_runs.c.status,
            shadow_runs.c.validation_status,
            shadow_runs.c.strict_accounting_match,
            shadow_runs.c.accounting_safety_status,
            shadow_runs.c.accounting_safety_issues_json,
            shadow_runs.c.metadata_quality_status,
            shadow_runs.c.metadata_issues_json,
            shadow_runs.c.invoice_number_evidence_status,
            shadow_runs.c.comparison_json,
            shadow_runs.c.validation_errors_json,
            shadow_runs.c.error_type,
            shadow_runs.c.preprocessing_ms,
            shadow_runs.c.openai_ms,
            shadow_runs.c.parsing_ms,
            shadow_runs.c.validation_ms,
            shadow_runs.c.total_ms,
            shadow_runs.c.input_tokens,
            shadow_runs.c.output_tokens,
            shadow_runs.c.reasoning_tokens,
            shadow_runs.c.total_tokens,
            shadow_runs.c.result_json.label("v2_result_json"),
            shadow_runs.c.created_at,
            jobs.c.result_json.label("v1_result_json"),
            metrics.c.status.label("v1_status"),
            metrics.c.processing_ms.label("v1_processing_ms"),
            metrics.c.openai_ms.label("v1_openai_ms"),
            metrics.c.input_tokens.label("v1_input_tokens"),
            metrics.c.output_tokens.label("v1_output_tokens"),
            metrics.c.reasoning_tokens.label("v1_reasoning_tokens"),
            metrics.c.total_tokens.label("v1_total_tokens"),
        )
        .select_from(
            shadow_runs.outerjoin(jobs, jobs.c.id == shadow_runs.c.job_id).outerjoin(
                metrics, metrics.c.id == latest_metric_id
            )
        )
        .where(shadow_runs.c.shadow_version == shadow_version)
    )
    if batch_id:
        query = query.where(
            func.coalesce(shadow_runs.c.batch_id, jobs.c.batch_id) == batch_id
        )
    if job_min is not None:
        query = query.where(shadow_runs.c.job_id >= job_min)
    if job_max is not None:
        query = query.where(shadow_runs.c.job_id <= job_max)
    if latest is not None:
        query = query.order_by(shadow_runs.c.created_at.desc(), shadow_runs.c.id.desc()).limit(latest)
    else:
        query = query.order_by(shadow_runs.c.job_id.asc(), shadow_runs.c.id.asc())

    with engine.connect() as conn:
        return [dict(row) for row in conn.execute(query).mappings().all()]


def _format_number(value: Any, digits: int = 1) -> str:
    if value is None:
        return "n/a"
    if isinstance(value, int):
        return str(value)
    return f"{float(value):,.{digits}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _format_percent(value: float | None) -> str:
    return "n/a" if value is None else f"{value:.1f}%"


def _format_job_ids(job_ids: Sequence[int]) -> str:
    return ", ".join(str(job_id) for job_id in job_ids) if job_ids else "ninguno"


def render_benchmark_report(report: Mapping[str, Any]) -> str:
    """Render a concise console report without leaking source document details."""
    scope = report["scope"]
    volume = report["volume"]
    rates = report["rates"]
    latency = report["latency"]
    lines = [
        "Ledged Invoice Engine V2 - benchmark persistido",
        "=" * 54,
        f"Versión shadow: {scope['shadow_version']}",
        f"Filtro: {scope.get('description') or 'todos los shadow runs de la versión'}",
        "",
        "VOLUMEN",
        f"  Total documentos: {volume['total']}",
        f"  Elegibles V2: {volume['eligible']} | No elegibles: {volume['ineligible']}",
        f"  Completados V2: {volume['completed']} | Pendientes: {volume['pending']} | Error técnico: {volume['technical_error']}",
        f"  Validación: passed={volume['validation_passed']} | failed={volume['validation_failed']}",
        f"  Strict accounting: true={volume['strict_match']} | false={volume['strict_mismatch']}",
        "",
        "ACCOUNTING SAFETY (V5+)",
        f"  Passed: {volume['accounting_safety_passed']} | Failed: {volume['accounting_safety_failed']} | Sin evaluar: {volume['accounting_safety_unassessed']}",
        f"  Passed sobre evaluados: {_format_percent(rates['accounting_safety_passed_of_completed'])}",
        "",
        "METADATA QUALITY (V5+)",
        f"  Confirmed: {volume['metadata_confirmed']} | Review required: {volume['metadata_review_required']} | Failed: {volume['metadata_failed']} | Sin evaluar: {volume['metadata_unassessed']}",
        f"  Confirmed sobre evaluados: {_format_percent(rates['metadata_confirmed_of_evaluated'])}",
        f"  Confirmed entre accounting-safe: {_format_percent(rates['metadata_confirmed_of_accounting_safe'])}",
        f"  Contablemente seguros con metadata review: {volume['accounting_safe_metadata_review']}",
        f"  Fully confirmed fast path candidate (solo benchmark): {report['fully_confirmed_fast_path_candidate_count']} ({_format_percent(rates['fully_confirmed_fast_path_candidate'])} del total; {_format_percent(rates['fully_confirmed_fast_path_candidate_of_eligible_completed'])} de elegibles completados)",
        "",
        "TASAS",
        f"  Elegibilidad: {_format_percent(rates['eligibility'])}",
        f"  Validación passed sobre V2 completados: {_format_percent(rates['validation_passed_of_completed'])}",
        f"  Strict match sobre V2 completados: {_format_percent(rates['strict_match_of_completed'])}",
        f"  Strict match entre validation passed comparables: {_format_percent(rates['strict_match_of_validation_passed'])}",
        f"  Strict match entre accounting-safe comparables: {_format_percent(rates['strict_match_of_accounting_safe'])}",
        f"  Safe fast path candidate (solo benchmark): {report['safe_fast_path_candidate_count']} ({_format_percent(rates['safe_fast_path_candidate'])})",
        f"  Requerirían fallback a V1 (conceptual): {_format_percent(rates['requires_v1_fallback'])}",
        f"  Fallback por accounting: {volume['fallback_accounting']} | por metadata: {volume['fallback_metadata']} | no elegibles: {volume['fallback_not_eligible']} | sin evaluar o strict mismatch: {volume['fallback_unassessed_or_mismatch']}",
        "",
        "LATENCIA (ms)",
        "  Fuente                         n       media       p50       p95",
            f"  V1 processing_ms      {latency['v1_processing_ms']['count']:>8} {_format_number(latency['v1_processing_ms']['mean']):>11} {_format_number(latency['v1_processing_ms']['p50']):>9} {_format_number(latency['v1_processing_ms']['p95']):>9}",
            f"  V2 total_ms           {latency['v2_total_ms']['count']:>8} {_format_number(latency['v2_total_ms']['mean']):>11} {_format_number(latency['v2_total_ms']['p50']):>9} {_format_number(latency['v2_total_ms']['p95']):>9}",
            f"  V1 openai_ms          {latency['v1_openai_ms']['count']:>8} {_format_number(latency['v1_openai_ms']['mean']):>11} {_format_number(latency['v1_openai_ms']['p50']):>9} {_format_number(latency['v1_openai_ms']['p95']):>9}",
            f"  V2 openai_ms          {latency['v2_openai_ms']['count']:>8} {_format_number(latency['v2_openai_ms']['mean']):>11} {_format_number(latency['v2_openai_ms']['p50']):>9} {_format_number(latency['v2_openai_ms']['p95']):>9}",
    ]
    paired = latency["paired"]
    lines.extend(
        [
            f"  Comparación pareada V1/V2: n={paired['count']} | reducción media={_format_number(paired['absolute_reduction'])} ms | reducción={_format_percent(paired['reduction_percent'])} | speedup={_format_number(paired['speedup'], 2)}x",
            f"  OpenAI pareado V1/V2: n={latency['openai_paired']['count']} | reducción media={_format_number(latency['openai_paired']['absolute_reduction'])} ms | reducción={_format_percent(latency['openai_paired']['reduction_percent'])}",
            "",
            "TOKENS (promedios pareados cuando ambas rutas informan el dato)",
            "  Tipo           V1 n       V1    V2 n       V2  Parejas     ahorro       %",
        ]
    )
    for field, _label in _TOKEN_FIELDS:
        token_report = report["tokens"][field]
        token = token_report["paired"]
        lines.append(
            f"  {token_report['label']:<12} {token_report['v1']['count']:>5} {_format_number(token_report['v1']['mean']):>8} {token_report['v2']['count']:>7} {_format_number(token_report['v2']['mean']):>8} {token['count']:>8} {_format_number(token['absolute_reduction']):>10} {_format_percent(token['reduction_percent']):>8}"
        )

    discrepancies = report["discrepancies"]
    lines.extend(["", "DISCREPANCIAS"])
    failed = discrepancies["validation_failed"]
    lines.append(
        f"  Validation failed: {failed['count']} | jobs: {_format_job_ids(failed['job_ids'])}"
    )
    for field, details in failed["fields"].items():
        lines.append(
            f"    {field}: {details['count']} ({_format_percent(details['percent'])}) | jobs: {_format_job_ids(details['job_ids'])}"
        )
    strict = discrepancies["validation_passed_strict_mismatch"]
    lines.append(f"  Validation passed + strict mismatch: {strict['count']}")
    for field, details in strict["fields"].items():
        lines.append(
            f"    {field}: {details['count']} ({_format_percent(details['percent'])}) | jobs: {_format_job_ids(details['job_ids'])}"
        )
    if not failed["fields"] and not strict["fields"]:
        lines.append("    Sin campos de discrepancia persistidos en el alcance seleccionado.")

    lines.extend(["", "MOTIVOS DE VALIDACIÓN"])
    if report["validation_errors"]:
        lines.extend(
            f"  {code}: {count}" for code, count in report["validation_errors"].items()
        )
    else:
        lines.append("  Ninguno")

    lines.extend(["", "ELEGIBILIDAD: motivos de exclusión"])
    if report["eligibility_reasons"]:
        lines.extend(
            f"  {reason}: {count}" for reason, count in report["eligibility_reasons"].items()
        )
    else:
        lines.append("  Ninguno")

    lines.extend(["", "CASOS A REVISAR"])
    if report["review_rows"]:
        for row in report["review_rows"]:
            strict_value = row["strict_accounting_match"]
            strict_label = "true" if strict_value is True else "false" if strict_value is False else "n/a"
            lines.append(
                f"  Job {row['job_id']}: status={row['status']}, validation={row['validation_status']}, strict={strict_label}, accounting={row['accounting_safety_status']}, metadata={row['metadata_quality_status']}"
            )
    else:
        lines.append("  Ninguno")
    lines.extend(
        [
            "",
            "Nota: el informe no imprime archivos, nombres, PDFs, texto extraído, prompts ni payloads de resultados.",
            "Las métricas V1 sin telemetría histórica se muestran como n/a y no se estiman.",
        ]
    )
    return "\n".join(lines)


def _positive_int(value: str) -> int:
    try:
        parsed = int(value)
    except ValueError as error:
        raise argparse.ArgumentTypeError("Debe ser un entero positivo.") from error
    if parsed < 1:
        raise argparse.ArgumentTypeError("Debe ser un entero positivo.")
    return parsed


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Informe read-only para benchmarks persistidos de Invoice Engine V2."
    )
    parser.add_argument("--version", required=True, help="Versión shadow, por ejemplo v2-sol-text-v4.")
    parser.add_argument("--batch-id", help="Restringe el informe a un lote persistido.")
    parser.add_argument("--job-min", type=_positive_int, help="Job ID mínimo inclusivo.")
    parser.add_argument("--job-max", type=_positive_int, help="Job ID máximo inclusivo.")
    parser.add_argument("--latest", type=_positive_int, help="Últimos N shadow runs tras aplicar filtros.")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.job_min is not None and args.job_max is not None and args.job_min > args.job_max:
        raise SystemExit("--job-min no puede ser mayor que --job-max.")
    database_url = _normalize_database_url(os.getenv("DATABASE_URL", ""))
    if not database_url:
        raise SystemExit(
            "DATABASE_URL no está configurada. Ejecuta el script desde el Shell de Render "
            "o exporta una URL PostgreSQL local sin incluirla en el historial de comandos."
        )
    engine = create_engine(database_url, pool_pre_ping=True, future=True)
    try:
        rows = load_benchmark_rows(
            engine,
            shadow_version=args.version.strip(),
            batch_id=(args.batch_id or "").strip() or None,
            job_min=args.job_min,
            job_max=args.job_max,
            latest=args.latest,
        )
    finally:
        engine.dispose()
    scope_parts = []
    if args.batch_id:
        scope_parts.append(f"batch_id={args.batch_id}")
    if args.job_min is not None or args.job_max is not None:
        scope_parts.append(f"job_ids={args.job_min or '-'}..{args.job_max or '-'}")
    if args.latest is not None:
        scope_parts.append(f"latest={args.latest}")
    report = build_benchmark_report(
        rows,
        {
            "shadow_version": args.version.strip(),
            "description": ", ".join(scope_parts) if scope_parts else None,
        },
    )
    print(render_benchmark_report(report))
    return 0


if __name__ == "__main__":
    sys.exit(main())
