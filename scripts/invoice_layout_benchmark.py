#!/usr/bin/env python3
"""Local-only comparison of the current text view and canonical PDF words.

Does not call invoice analysis, OpenAI, S3 or a database. Only anonymous counts,
timings and fixed diagnostic codes may leave this process. No files are written.
"""

from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import fields, is_dataclass
from difflib import SequenceMatcher
import hashlib
import json
from pathlib import Path
import re
from statistics import median
import sys
import time

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from services.document_layout import (
    BOUNDARY_TOLERANCE, LAYOUT_VERSION, MAX_PAGES, OVERLAP_TOLERANCE,
    DocumentLayout, LayoutError, extract_native_document, serialize_document_layout,
)


MAX_FILE_BYTES = 20 * 1024 * 1024
MAX_SIMILARITY_CHARS = 60000
MAX_SIMILARITY_WORK = 25000000
DATE_PATTERN = re.compile(r"(?<!\d)\d{1,4}[-/.]\d{1,2}[-/.]\d{2,4}(?!\d)")
MONEY_PATTERN = re.compile(r"(?<!\d)(?:\d{1,3}(?:[. ]\d{3})+|\d+)[,.]\d{2}(?!\d)")
IDENTIFIER_PATTERN = re.compile(r"(?<!\w)\w+(?:[-/]\w+)*(?!\w)")
CURRENCY_PATTERN = re.compile(r"\u20ac|\$|\u00a3|\b(?:EUR|USD|GBP)\b", re.I)
PERCENT_PATTERN = re.compile(r"(?<!\d)\d+(?:[.,]\d+)?[ \t]*%")
# Shape-only counter. Bounded Spanish forms avoid consuming the next numeric
# column as part of a tax ID; this is not fiscal identity validation.
TAX_ID_PATTERN = re.compile(
    r"(?<![A-Z0-9])(?:ES[ .-]*)?(?:"
    r"[ABCDEFGHJNPQRSUVW][ .-]*\d(?:[ .-]*\d){6}[ .-]*[0-9A-J]|"
    r"[XYZ][ .-]*\d(?:[ .-]*\d){6}[ .-]*[A-Z]|"
    r"\d(?:[ .-]*\d){7}[ .-]*[A-Z]"
    r")(?![A-Z0-9])", re.I,
)
SAFE_ERRORS = {
    "file_limit_exceeded", "page_limit_exceeded", "word_limit_exceeded",
    "encrypted_pdf", "pdf_extraction_failed", "invalid_page_geometry",
    "invalid_page_rotation", "invalid_native_word", "duplicate_token_ids",
    "invalid_token_projection", "incomplete_token_projection",
    "geometry_work_limit_exceeded",
}


def geometry_metrics(layout) -> dict:
    """Structural counts only: a channel is NOT a supplier/recipient region."""
    multicolumn_pages = zones = 0
    for page in layout.pages:
        by_zone = {}
        for group in page.reading_groups:
            if group.kind == "column":
                by_zone.setdefault(group.zone, []).append(group)
        # Require multiple vertically coexisting multi-line channels, not two
        # single words placed at opposite ends of a page.
        detected = 0
        for groups in by_zone.values():
            multiline = [g for g in groups if len(g.rows) > 1]
            if any(max(a.rows[0].bbox[1], b.rows[0].bbox[1]) < min(a.rows[-1].bbox[3], b.rows[-1].bbox[3])
                   for i, a in enumerate(multiline) for b in multiline[i + 1:]):
                detected += 1
        zones += detected
        multicolumn_pages += bool(detected)
    gaps = [right.bbox[0] - left.bbox[2] for row in layout.rows
            for left, right in zip(row.segments, row.segments[1:])]
    flags = Counter(flag for token in layout.tokens for flag in token.geometry_flags)
    return {
        "multicolumn_geometry_pages": multicolumn_pages,
        "multicolumn_geometry_zones": zones,
        "multi_segment_rows": sum(len(row.segments) > 1 for row in layout.rows),
        "rows_with_multiple_source_lines": sum(len(row.source_lines) > 1 for row in layout.rows),
        "source_lines_split_across_rows": sum(
            count > 1 for count in Counter((row.page, source) for row in layout.rows for source in row.source_lines).values()
        ),
        "horizontal_segment_gaps": len(gaps),
        "horizontal_gap_points": {
            "min": round(min(gaps), 3) if gaps else None,
            "median": round(median(gaps), 3) if gaps else None,
            "max": round(max(gaps), 3) if gaps else None,
        },
        "overlapping_tokens": flags["overlap"],
        "same_band_overlapping_tokens": flags["same_band_overlap"],
        "cross_band_overlapping_tokens": flags["cross_band_overlap"],
        "overlap_pairs": sum(p.overlap_pairs for p in layout.pages),
        "duplicate_geometry_tokens": flags["duplicate_geometry"],
        "duplicate_geometry_pairs": sum(p.duplicate_geometry_pairs for p in layout.pages),
        "outside_cropbox_tokens": flags["outside_cropbox"],
        "outside_mediabox_tokens": flags["outside_mediabox"],
        "boundary_roundoff_tokens": flags["boundary_roundoff"],
        "nonhorizontal_tokens": flags["nonhorizontal_orientation"],
        "unknown_orientation_tokens": flags["unknown_orientation"],
        "rotated_pages": sum(bool(p.rotation) for p in layout.pages),
        "isolated_tokens": sum(bool(t.isolation_reasons) for t in layout.tokens),
    }


def current_text_view(raw_pages: list[str]) -> str:
    # Use the real legacy compactor, including its OCR-style amount repair.
    # Importing this module does not instantiate a model client or run analysis.
    from services.ai_invoice_service import (
        _compact_native_pdf_page_text, _get_invoice_v2_fast_text_max_chars,
    )

    parts = []
    for number, raw in enumerate(raw_pages, 1):
        compact = _compact_native_pdf_page_text(raw)
        if compact:
            parts.append(f"[P\u00c1GINA {number}]\n{compact}")
    text = "\n\n".join(parts).strip()
    limit = _get_invoice_v2_fast_text_max_chars()
    if len(text) > limit:
        marker = "\n\n[CONTENIDO INTERMEDIO OMITIDO POR L\u00cdMITE DE TAMA\u00d1O]\n\n"
        head = max(1, int((limit - len(marker)) * 0.6))
        tail = max(1, limit - len(marker) - head)
        text = text[:head] + marker + text[-tail:]
    return text


def _characters(text: str) -> Counter:
    return Counter(character for character in text if not character.isspace())


def _identifiers(text: str) -> Counter:
    return Counter(
        value for value in IDENTIFIER_PATTERN.findall(text)
        if any(c.isdigit() for c in value) and any(c.isalpha() for c in value)
    )


def _lexemes(text: str) -> dict[str, Counter]:
    return {
        "words": Counter(re.findall(r"\S+", text)),
        "identifiers": _identifiers(text),
        "dates": Counter(DATE_PATTERN.findall(text)),
        "monetary_strings": Counter(MONEY_PATTERN.findall(text)),
        "currency_symbols_codes": Counter(CURRENCY_PATTERN.findall(text)),
        "percentages": Counter(PERCENT_PATTERN.findall(text)),
        "tax_id_like": Counter(TAX_ID_PATTERN.findall(text)),
    }


def _preservation(reference: str, candidate: str) -> dict:
    left, right = _lexemes(reference), _lexemes(candidate)
    return {
        field: {
            "reference": sum(values.values()),
            "candidate": sum(right[field].values()),
            "preserved": sum((values & right[field]).values()),
            "rate_pct": round(100 * sum((values & right[field]).values()) / sum(values.values()), 2)
            if values else None,
        }
        for field, values in left.items()
    }


def _python_bytes(value, seen=None) -> int:
    """Approximate retained Python objects, not process RSS or peak memory."""
    seen = set() if seen is None else seen
    if id(value) in seen:
        return 0
    seen.add(id(value))
    size = sys.getsizeof(value)
    if is_dataclass(value):
        if hasattr(value, "__dict__"):
            size += _python_bytes(vars(value), seen)
        else:
            size += sum(_python_bytes(getattr(value, f.name), seen) for f in fields(value))
    elif isinstance(value, dict):
        size += sum(_python_bytes(k, seen) + _python_bytes(v, seen) for k, v in value.items())
    elif isinstance(value, (tuple, list, set)):
        size += sum(_python_bytes(item, seen) for item in value)
    return size


def diagnose_legacy_money_changes(raw_pages: list[str], layout) -> dict:
    """Reproduce the earlier 19-difference experiment, not the new serializer.

    Signed stage deltas telescope exactly to each missing multiplicity. No
    inference that a regex money match is the invoice's true monetary value.
    """
    from services.ai_invoice_service import (
        _build_fast_text_document_layout, _compact_native_pdf_page_text,
        _normalize_ocr_amount_text,
    )
    words = [[(*token.bbox, token.original_text, token.block, token.line, token.word)
              for token in page.tokens] for page in layout.pages]
    legacy = _build_fast_text_document_layout(words)
    legacy_pages = ["\n".join(row["text"] for row in page["rows"]) for page in legacy["pages"]]

    def counts(pages, transform=lambda text: text):
        return Counter(value for text in pages for value in MONEY_PATTERN.findall(transform(text)))

    current_raw, spatial_raw = counts(raw_pages), counts(legacy_pages)
    current_ocr = counts(raw_pages, _normalize_ocr_amount_text)
    spatial_ocr = counts(legacy_pages, _normalize_ocr_amount_text)
    current_final = counts(raw_pages, _compact_native_pdf_page_text)
    spatial_final = counts(legacy_pages, _compact_native_pdf_page_text)
    differences = []
    for index, (value, count) in enumerate(sorted((current_final - spatial_final).items()), 1):
        deltas = {
            "raw_layout_boundary_delta": current_raw[value] - spatial_raw[value],
            "current_ocr_delta": current_ocr[value] - current_raw[value],
            "spatial_ocr_delta": spatial_raw[value] - spatial_ocr[value],
            "current_compaction_delta": current_final[value] - current_ocr[value],
            "spatial_compaction_delta": spatial_ocr[value] - spatial_final[value],
        }
        causes = []
        containing_shapes = sorted({
            re.sub(r"\d", "N", candidate)
            for candidate in spatial_raw if candidate != value and value in candidate
        })
        if deltas["raw_layout_boundary_delta"]:
            # All raw token characters survive. The regex can span spaces but
            # not newlines: row assembly changes its boundaries and grouping.
            causes.append("D_lexical_join_to_neighboring_numeric_token" if containing_shapes
                          else "B_E_lexical_boundaries_require_review")
        if deltas["current_ocr_delta"] or deltas["spatial_ocr_delta"]:
            causes.append("C_ocr_amount_normalization")
        if deltas["current_compaction_delta"] or deltas["spatial_compaction_delta"]:
            causes.append("F_compaction_or_deduplication")
        differences.append({
            "difference": index, "masked_shape": re.sub(r"\d", "N", value),
            "missing_occurrences": count, "causes": causes,
            "stage_deltas": deltas, "accounted_for": sum(deltas.values()) == count,
            "containing_raw_money_shapes": containing_shapes,
        })
    return {
        "missing_occurrences": sum(item["missing_occurrences"] for item in differences),
        "differences": differences,
        "all_accounted_for": all(item["accounted_for"] for item in differences),
        "scope": "legacy_v14_row_join_plus_legacy_compactor_not_canonical_serializer",
    }


def compare_text_views(current: str, derived: str) -> dict:
    """Lexical diagnostics only: not ground truth or field-level accuracy."""
    normalized = [" ".join(value.split()) for value in (current, derived)]
    similarity = None
    if (max(map(len, normalized)) <= MAX_SIMILARITY_CHARS
            and len(normalized[0]) * len(normalized[1]) <= MAX_SIMILARITY_WORK):
        similarity = round(100 * SequenceMatcher(None, *normalized, autojunk=False).ratio(), 2)
    missing = _characters(current) - _characters(derived)
    return {
        "current_chars": len(current), "derived_chars": len(derived),
        "current_lines": len(current.splitlines()), "derived_lines": len(derived.splitlines()),
        "current_approx_tokens": round(len(current) / 4),
        "derived_approx_tokens": round(len(derived) / 4),
        "sequence_similarity_pct": similarity,
        "nonspace_chars_missing": sum(missing.values()),
        "nonspace_chars_added": sum((_characters(derived) - _characters(current)).values()),
        "symbols_missing": sum(n for c, n in missing.items() if not c.isalnum()),
        "identifier_lexemes_missing": sum((_identifiers(current) - _identifiers(derived)).values()),
        "date_lexemes_missing": sum((Counter(DATE_PATTERN.findall(current)) - Counter(DATE_PATTERN.findall(derived))).values()),
        "money_lexemes_missing": sum((Counter(MONEY_PATTERN.findall(current)) - Counter(MONEY_PATTERN.findall(derived))).values()),
    }


def compare_pdf_bytes(pdf_bytes: bytes) -> dict:
    # Exclude the one-off import of legacy helpers from per-document timing.
    current_text_view([])
    if len(pdf_bytes) > MAX_FILE_BYTES:
        raise LayoutError("file_limit_exceeded")
    started = time.perf_counter()
    extraction = extract_native_document(pdf_bytes)
    layout = extraction.layout
    layout_ms = (time.perf_counter() - started) * 1000
    started = time.perf_counter()
    view = serialize_document_layout(layout)
    serialization_ms = (time.perf_counter() - started) * 1000
    started = time.perf_counter()
    raw_pages = list(extraction.native_text_pages)
    current = current_text_view(raw_pages)
    current_ms = (time.perf_counter() - started) * 1000
    tokens = {t.token_id: t for page in layout.pages for t in page.tokens}
    native = "".join(token.text for token in tokens.values())
    raw = "\n".join(raw_pages)
    return {
        "pages": len(layout.pages), "words": len(tokens),
        "rows": sum(len(page.rows) for page in layout.pages),
        "segments": sum(len(row.segments) for page in layout.pages for row in page.rows),
        "source_blocks": len(layout.blocks),
        "source_lines": len({(t.page, t.block, t.line) for t in layout.tokens}),
        "native_token_spans_preserved": all(
            view.text[span.start:span.end] == tokens[span.token_id].text for span in view.spans
        ) and len(view.spans) == len(tokens),
        "reversible_token_spans": sum(
            view.text[s.start:s.end] == tokens[s.token_id].text
            and view.token_ids_for_range(s.start, s.end) == (s.token_id,) for s in view.spans
        ),
        "geometry": geometry_metrics(layout),
        "raw_nonspace_chars_missing_in_words": sum((_characters(raw) - _characters(native)).values()),
        "current_text_truncated": "[CONTENIDO INTERMEDIO OMITIDO POR" in current,
        "diagnostics": sorted({code for page in layout.pages for code in page.diagnostics}),
        "current_compaction_ms": round(current_ms, 3),
        "shared_extraction_and_layout_ms": round(layout_ms, 3),
        "serialization_ms": round(serialization_ms, 3),
        "approx_layout_python_bytes": _python_bytes(layout),
        "approx_serialized_view_python_bytes": _python_bytes(view),
        "source_token_lexeme_preservation": _preservation(
            "\n".join(t.original_text for t in layout.tokens),
            "\n".join(view.text[span.start:span.end] for span in view.spans),
        ),
        "current_text_lexeme_preservation": _preservation(current, view.text),
        "legacy_money_diagnostics": diagnose_legacy_money_changes(raw_pages, layout),
        **compare_text_views(current, view.text),
    }


def select_local_pdfs(root: Path, limit: int, *, digital_invoices: bool = False) -> list[Path]:
    """Deterministic sample, preferring distinct directories, not suppliers."""
    candidates = sorted(
        (path for path in root.rglob("*") if path.suffix.lower() == ".pdf" and path.is_file() and not path.is_symlink()),
        key=lambda path: hashlib.sha256(str(path.relative_to(root)).encode()).digest(),
    )
    selected, parents = [], set()
    if digital_invoices:
        import fitz

        # Same eligibility and ordering used by the earlier 20-PDF audit.
        eligible = []
        for path in candidates:
            # The original audit sampled rglob('*.pdf'), case-sensitively.
            if path.suffix != ".pdf":
                continue
            if path.stat().st_size > MAX_FILE_BYTES:
                continue
            try:
                with fitz.open(path) as document:
                    if document.needs_pass or not 1 <= len(document) <= 10:
                        continue
                    text = "".join(p.get_text("text", sort=True) for p in document)
                if len(text.strip()) >= 100 and re.search(r"factura|invoice", text, re.I):
                    eligible.append(path)
            except Exception:
                continue
        candidates = eligible
    for path in candidates:
        if path.parent not in parents:
            selected.append(path)
            parents.add(path.parent)
            if len(selected) == limit:
                return selected
    for path in candidates:
        if path not in selected:
            selected.append(path)
            if len(selected) == limit:
                break
    return selected


def build_report(paths: list[Path]) -> dict:
    documents = []
    for index, path in enumerate(paths, 1):
        record = {"document": f"document_{index:02d}"}
        try:
            if path.stat().st_size > MAX_FILE_BYTES:
                raise LayoutError("file_limit_exceeded")
            record.update(compare_pdf_bytes(path.read_bytes()))
            record["status"] = "measured"
        except LayoutError as exc:
            record.update(status="error", error_code=str(exc) if str(exc) in SAFE_ERRORS else "local_comparison_failed")
        except Exception:
            # Exception messages may contain document content or private paths.
            record.update(status="error", error_code="local_comparison_failed")
        documents.append(record)
    measured = [r for r in documents if r["status"] == "measured"]
    return {
        "layout_version": LAYOUT_VERSION,
        "scope": "offline_only_no_acceptance_decision",
        "token_estimate": "characters_divided_by_4_not_model_tokens",
        "geometry_scope": "geometric_channels_not_semantic_regions_or_tables",
        "preservation_scope": "original_word_fragments_recovered_through_exact_token_spans",
        "geometry_tolerance_points": {"boundary": BOUNDARY_TOLERANCE, "overlap": OVERLAP_TOLERANCE},
        "limits": {"pages": MAX_PAGES, "file_bytes": MAX_FILE_BYTES},
        "documents": documents,
        "summary": {
            "total": len(documents), "measured": len(measured),
            "errors": len(documents) - len(measured),
            "documents_with_native_words": sum(r["words"] > 0 for r in measured),
            "token_projection_failures": sum(not r["native_token_spans_preserved"] for r in measured),
            "serialization_reversibility_rate_pct": round(
                100 * sum(r["reversible_token_spans"] for r in measured) / sum(r["words"] for r in measured), 2
            ) if sum(r["words"] for r in measured) else None,
            "geometry": {
                key: sum(r["geometry"][key] for r in measured)
                for key in geometry_metrics(DocumentLayout(()))
                if key != "horizontal_gap_points"
            },
            "legacy_missing_money_occurrences": sum(r["legacy_money_diagnostics"]["missing_occurrences"] for r in measured),
            "legacy_differences_accounted_for": all(r["legacy_money_diagnostics"]["all_accounted_for"] for r in measured),
            "source_token_lexeme_preservation": {
                field: {
                    key: sum(r["source_token_lexeme_preservation"][field][key] for r in measured)
                    for key in ("reference", "candidate", "preserved")
                }
                for field in _lexemes("")
            },
            **{key: sum(r[key] for r in measured) for key in (
                "current_chars", "derived_chars", "current_lines", "derived_lines",
                "raw_nonspace_chars_missing_in_words", "money_lexemes_missing",
                "identifier_lexemes_missing", "date_lexemes_missing",
            )},
        },
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", required=True, type=Path, help="Local PDF directory (read-only)")
    parser.add_argument("--limit", type=int, default=20)
    parser.add_argument("--json", action="store_true", help="Anonymous operational JSON, never document data")
    parser.add_argument("--digital-invoices", action="store_true", help="Match the earlier digital-invoice sample criteria")
    args = parser.parse_args(argv)
    if not args.root.is_dir() or not 1 <= args.limit <= 1000:
        parser.error("A local directory and a limit between 1 and 1000 are required")
    report = build_report(select_local_pdfs(args.root, args.limit, digital_invoices=args.digital_invoices))
    if args.json:
        print(json.dumps(report, ensure_ascii=True, indent=2))
    else:
        print(f"Offline canonical layout: {LAYOUT_VERSION}")
        print("Document       Words  Chars(old/new)  Lines(old/new)  Similarity  Token spans")
        for record in report["documents"]:
            if record["status"] == "error":
                print(f"{record['document']}: {record['error_code']}")
                continue
            similarity = record["sequence_similarity_pct"]
            print(
                f"{record['document']}  {record['words']:6d}  "
                f"{record['current_chars']}/{record['derived_chars']}  "
                f"{record['current_lines']}/{record['derived_lines']}  "
                f"{similarity if similarity is not None else 'not measured'}  "
                f"{'preserved' if record['native_token_spans_preserved'] else 'FAILED'}"
            )
        print(json.dumps(report["summary"], ensure_ascii=True))
        print("Lexical comparison only; not accounting accuracy or permission to accept V2.")
    return 1 if report["summary"]["errors"] or not report["documents"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
