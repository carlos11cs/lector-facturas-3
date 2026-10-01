# Canonical DocumentLayout: offline experiment

This experiment is NOT imported by `app.py`, `worker.py`, or either invoice
pipeline. It changes no verifier, prompt, model, schema, prepared text, queue,
concurrency, shadow version, dependency, environment variable or database.
It does not call OpenAI, S3 or PostgreSQL and never processes a document through
the official analysis pipeline.

## V15 preservation

The previous tracked V15 diff was preserved in the named stash
`v15-spatial-conflict-resolution-wip`:

- Stash object: `b55b86e4ac3cf7115f9bd927e58f753b0f6205d8`.
- SHA-256 of the binary diff before stashing and of the stash patch:
  `8dbcaf5901a4eeaeb604311aa6084cf394cfc69e410acd58b1ce3984fad213cd`.
- `git status --short` was empty after stashing and before restoring only the
  separate offline experiment. V15 was not reapplied or deleted.
- The separate initial experiment backup is retained as
  `canonical-document-layout-offline-wip`, object
  `4607c4ba0716db5e85005b498c68ba8d9992df3b`.

No commit, push or deployment was performed. Stash operations create Git stash
objects, not a new commit on main.

## Representation

`services/document_layout.py` defines immutable evidence and derived views:

```text
DocumentLayout(version, pages)
  Page(width, height, cropbox, mediabox, rotation, coordinate_reference)
    Tokens: exact text, bbox, normalized bbox, block/line/word, direction,
            writing mode and structural flags
    Geometric rows
      Segments: token IDs, bbox and original block/line provenance
    Original blocks
    Diagnostics
```

Tokens are the only raw evidence. Rows, segments and serializer reading groups are
recomputable geometry views; no supplier, recipient, totals, VAT or table role is
introduced. `regions` remains empty. Original text and coordinates are never
normalized, clipped or overwritten.

Token IDs use page/block/line/word/occurrence. Duplicate source positions receive
different IDs. Each token retains its PyMuPDF block/line/word provenance even when
geometry places fragments from different source lines into the same visual row.
Conversely, source metadata cannot force geometrically separate text into one row.

PyMuPDF 1.23.8 creates one `TextPage` per page. Words, the optional legacy text
reference and line direction metadata all come from that same object. Direction is
accepted only when the word bbox is inside the corresponding line bbox. Missing or
invalid direction is explicit, never silently assumed horizontal.

Coordinates returned by text extraction are retained in unrotated cropbox-relative
top-left point space. Page rotation, CropBox, MediaBox and normalized coordinates
are retained separately. MediaBox bounds are translated into the same coordinate
space before validation. A 0.01-point tolerance covers floating-point boundary
noise; values beyond it are marked, kept exact and isolated from normal grouping.

Overlap detection uses a bounded sweep and a 0.25-point tolerance. Every overlap
is reported. Cross-line font-box intersections stay in their already distinct
rows; same-band overlaps, practically duplicate geometry, non-horizontal/unknown
orientation, and out-of-bounds tokens are serialized as isolated evidence. Nothing
is deduplicated. Work is capped at 500,000 geometric comparisons and fails with a
safe code instead of returning partial data.

## Multi-column serialization

`serialize_document_layout()` returns exact text spans plus `token_ids_for_range()`.
Every token appears once and every token span maps back to the original token ID.
Generated delimiters intentionally map to no token.

Persistent horizontal gutters create derived reading channels. A spanning row
closes the current zone instead of joining its columns. Each channel is serialized
contiguously between explicit `[COLUMN ...]` / `[/COLUMN]` markers, so two
independent columns cannot become `left1 right1 left2 right2`. Rows and segments
remain present in the canonical object, so column serialization does not destroy
the alternate row view used by borderless tables.

Page markers preserve page boundaries; rotation and isolated evidence receive
explicit markers. Tabs preserve multiple segments inside one non-column row. There
is no HTML decoding, amount repair, OCR repair, line deduplication or truncation.
Invalid, missing or repeated token references fail explicitly.

This remains a candidate OFFLINE view. It is not imported by production and does
not replace `prepared["text"]`.

## Running the report

```bash
.venv311/bin/python scripts/invoice_layout_benchmark.py \
  --root /path/to/local/invoices --limit 20 --digital-invoices --json
```

The report is read-only and contains anonymous labels, counts, fixed diagnostic
codes and timings. It never prints filenames, PDF/text content, identifiers,
amounts, coordinates, prompts, credentials or exception details.

Selection is deterministic for a fixed directory state: lowercase `.pdf`, one to
ten pages, at least 100 native characters and an invoice keyword, ordered by a hash
of relative path and preferring distinct directories. During this phase another
local PDF became eligible and would have displaced the original twentieth item.
Final figures therefore use the original twenty-path snapshot; its anonymous path
set fingerprint is
`35f0d4341e8ff4bc0ddc54651006e7cbbd759c95d4103d76971916e7dad6b55f`.
No private path is persisted. This demonstrates why corpus membership must be
pinned before comparing future benchmark runs.

## Final 20-document audit

Python 3.11.16 / PyMuPDF 1.23.8; 20 PDFs, 30 pages, 6,965 words, 791 original
blocks, 2,614 original lines, 920 geometric rows and 2,646 segments.

### Exact preservation

| Original token category | Original | Recovered via spans | Rate |
| --- | ---: | ---: | ---: |
| Word tokens | 6,965 | 6,965 | 100% |
| Identifier-like fragments | 218 | 218 | 100% |
| Dates | 89 | 89 | 100% |
| Monetary strings | 443 | 443 | 100% |
| Currency symbols/codes | 122 | 122 | 100% |
| Percentages within tokens | 28 | 28 | 100% |
| Tax-ID-like fragments | 49 | 49 | 100% |

All 6,965 spans are individually reversible: 100%. Raw-text-versus-word non-space
character loss is zero. These are preservation measurements, not extraction
accuracy or proof that a token is the correct accounting field.

The current compact view has 44,580 characters / 2,542 lines / approximately
11,145 tokens (`chars / 4`). The structural serializer has 52,374 characters /
1,968 lines / approximately 13,093 tokens. Its extra characters are explicit
structural markers. Mean sequential similarity is 70.022%; reordering complete
columns intentionally lowers this non-safety metric.

### Geometry

| Diagnostic | Corpus result |
| --- | ---: |
| Pages with persistent multi-channel geometry | 27 / 30 |
| Multi-channel zones | 60 |
| Rows with multiple segments | 472 |
| Horizontal gaps between segments | 1,726 |
| Horizontal gap min / median / p95 / max | 0.257 / 27.123 / 210.222 / 419.244 pt |
| Source lines represented in more than one geometric row | 5 |
| Tokens participating in any bbox overlap | 1,381 |
| Overlap pairs | 1,363 |
| Same-band overlapping tokens | 2 |
| Practically duplicate-geometry tokens / pairs | 2 / 1 |
| Partially outside CropBox / MediaBox | 2 / 2 |
| Non-horizontal orientation | 26 |
| Unknown orientation | 0 |
| Rotated PDF pages | 0 |
| Tokens isolated from normal grouping | 30 |

The high general overlap count is not 1,381 duplicated words: 1,362 pairs are
cross-band font boxes whose ascenders/descenders intersect while remaining on
different rows. Only one same-band pair has duplicate geometry; both occurrences
survive and are isolated. The two out-of-page bboxes exceed the page by 1.554 and
0.661 points, not floating-point noise; they are retained unchanged and isolated.

The 26 non-horizontal tokens consist of 25 at -90 degrees and one at -45 degrees,
across five documents. They are preserved in separate rotated segments. No page in
this corpus has a page-level rotation, so 90/180/270-degree page handling is covered
by adversarial fixtures rather than inferred from this sample.

`27 pages` is deliberately called multi-channel geometry, not semantic columns:
tables and side-by-side boxes can have the same structure. The canonical rows and
segments remain available. Synthetic two-column tests prove contiguous left-channel
then right-channel serialization, including equal-Y, staggered, spanning-header and
spanning-footer cases; no supplier/customer interpretation is involved.

## Legacy 19-difference diagnosis and OCR recommendation

The earlier 19 differences remain fully accounted for: 13 are caused by
`_normalize_ocr_amount_text`; six arise when neighboring numeric tokens acquire a
space accepted by the money regex. No source digit or token is lost by the new
serializer.

`_normalize_ocr_amount_text` currently HTML-decodes, normalizes NBSP and applies a
regex whose `\s` can cross a line or tab. There is no valid general reason to apply
that destructive repair to reliable native PDF text. The recommended future design
combines A and B: preserve canonical evidence, and expose a separate traced repair
view only for proven OCR/scanned input. Each repair should link to source token IDs,
remain reversible/auditable and never replace native tokens. A malformed native PDF
should be diagnosed or fall back, not be silently rewritten. Production behavior is
unchanged in this experiment.

## Performance and memory

Mean of per-document medians over three local runs:

- Shared TextPage extraction, direction metadata and layout: 17.85 ms/PDF.
- Structural serialization and span index: 1.44 ms/PDF.
- Legacy compaction reference: 0.34 ms/PDF.
- Approximate retained Python layout: 348,446 bytes/PDF (340 KiB).
- Approximate serialized view and spans: 102,545 bytes/PDF (100 KiB).

These memory figures include retained Python objects, not MuPDF native allocations,
RSS or peak process memory. Benchmark lexical comparisons and the old 19-difference
diagnosis are extra offline costs and are not proposed request-path work.

## Tests and unrelated expired fixtures

The offline suite covers independent equal-Y/staggered columns, spanning headers
and footers, visual rows split across blocks, label-above-value, borderless tables,
BASE/IVA/TOTAL, multiple VAT rates, duplicate and non-adjacent overlap, partially
outside bboxes, CropBox/MediaBox distinction, boundary tolerance, page rotation,
vertical/reversed/unknown direction, repeated headers and amounts, fragmented IDs,
separate currency symbols, two pages, exact span/range reversibility, stable IDs,
resource limits, privacy and absence of network/pipeline calls.

Eight existing tests share an absolute `expires_at=2026-10-01T10:00:00` and fail
after that instant because jobs are filtered as expired. They were not modified:

1. `test_two_claim_attempts_cannot_take_the_same_job`
2. `test_parallel_sqlite_claimers_only_receive_one_copy`
3. `test_several_jobs_can_be_claimed_without_duplication`
4. `test_expired_lease_is_reclaimed_with_a_new_token`
5. `test_stale_worker_cannot_finish_a_recovered_job`
6. `test_renewal_requires_the_current_token_and_is_measured`
7. `test_company_fairness_reserves_capacity_for_another_company`
8. `test_batch_serialization_and_paginated_endpoint_are_scoped_to_company`

A normal full run executes 369 tests: 361 pass and those eight produce three
failures plus five errors. A controlled verification freezes only that module's
application clock at `2026-09-30T12:00:00Z`; it does not modify production or test
source. The focused offline module passes 57/57 tests; the controlled full suite
passes 369/369 with zero failures, errors or skips. Network sockets were disabled
and no PostgreSQL was used; database tests use their existing local/in-memory
SQLite fixtures.

## Shadow gate

The representation meets this corpus gate: 100% relevant-fragment preservation,
100% reversible token spans, explicit page/segment/channel boundaries, no column
interleaving in adversarial fixtures, explicit rotation/orientation handling, no
silent token loss and no destructive normalization.

It is therefore suitable for a NEW, explicitly versioned SHADOW experiment in
which V2 receives canonical-derived text. This is approval to measure a candidate
input only, not to migrate production, the verifier or V1, and not permission to
make it the real Fast Path. A broader corpus, including page-rotated and heavily
cropped PDFs, remains a residual risk to measure in that future shadow.
