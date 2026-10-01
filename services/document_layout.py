"""Ephemeral PDF representation for canonical-layout experiments.

The canonical shadow input imports this module without changing V1 or the
legacy V2 path. Rows and segments describe geometry, not fiscal roles, tables,
reading-order certainty or an acceptance decision.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field, replace
import math
from statistics import median
from typing import Iterable


LAYOUT_VERSION = "canonical-words-experiment-v2"
MAX_PAGES = 100
MAX_WORDS = 50000
MAX_GEOMETRY_COMPARISONS = 500000
# Points, not pixels. Only floating-point boundary noise is tolerated.
BOUNDARY_TOLERANCE = 0.01
OVERLAP_TOLERANCE = 0.25
ISOLATING_FLAGS = frozenset({
    "duplicate_geometry", "same_band_overlap", "outside_cropbox", "outside_mediabox",
    "unknown_orientation", "nonhorizontal_orientation",
})
BBox = tuple[float, float, float, float]


class LayoutError(ValueError):
    """An operational code, never document text or a file path."""


@dataclass(frozen=True)
class LayoutToken:
    token_id: str
    page: int
    bbox: BBox = field(repr=False)
    original_text: str = field(repr=False)
    block: int
    line: int
    word: int
    direction: tuple[float, float] | None = None
    writing_mode: int | None = None
    normalized_bbox: BBox | None = field(default=None, repr=False)
    geometry_flags: tuple[str, ...] = ()

    @property
    def text(self) -> str:
        return self.original_text

    @property
    def isolation_reasons(self) -> tuple[str, ...]:
        return tuple(flag for flag in self.geometry_flags if flag in ISOLATING_FLAGS)


@dataclass(frozen=True)
class LayoutSegment:
    token_ids: tuple[str, ...]
    bbox: BBox = field(repr=False)
    source_lines: tuple[tuple[int, int], ...]


@dataclass(frozen=True)
class LayoutRow:
    page: int
    segments: tuple[LayoutSegment, ...]
    bbox: BBox = field(repr=False)
    source_lines: tuple[tuple[int, int], ...]
    isolation_reasons: tuple[str, ...] = ()

    @property
    def token_ids(self) -> tuple[str, ...]:
        return tuple(token_id for segment in self.segments for token_id in segment.token_ids)


@dataclass(frozen=True)
class LayoutBlock:
    page: int
    block: int
    token_ids: tuple[str, ...]
    bbox: BBox = field(repr=False)


@dataclass(frozen=True)
class LayoutRegion:
    region_id: str
    page: int
    token_ids: tuple[str, ...]
    bbox: BBox = field(repr=False)
    classification: str | None = None


@dataclass(frozen=True)
class LayoutReadingGroup:
    """A geometric channel, never a semantic region or table."""

    kind: str
    zone: int
    column: int
    rows: tuple[LayoutRow, ...] = field(repr=False)


@dataclass(frozen=True)
class LayoutPage:
    page: int
    width: float
    height: float
    rotation: int
    tokens: tuple[LayoutToken, ...] = field(repr=False)
    rows: tuple[LayoutRow, ...] = field(repr=False)
    diagnostics: tuple[str, ...]
    blocks: tuple[LayoutBlock, ...] = field(repr=False)
    regions: tuple[LayoutRegion, ...] = field(default=(), repr=False)
    writing_direction_verified: bool = False
    cropbox: BBox | None = field(default=None, repr=False)
    mediabox: BBox | None = field(default=None, repr=False)
    media_bounds: BBox | None = field(default=None, repr=False)
    coordinate_reference: str = "unrotated_cropbox_top_left_points"
    overlap_pairs: int = 0
    duplicate_geometry_pairs: int = 0

    @property
    def reading_groups(self) -> tuple[LayoutReadingGroup, ...]:
        return _reading_groups(self.rows)


@dataclass(frozen=True)
class DocumentLayout:
    pages: tuple[LayoutPage, ...] = field(repr=False)
    version: str = LAYOUT_VERSION

    @property
    def tokens(self) -> tuple[LayoutToken, ...]:
        return tuple(token for page in self.pages for token in page.tokens)

    @property
    def rows(self) -> tuple[LayoutRow, ...]:
        return tuple(row for page in self.pages for row in page.rows)

    @property
    def blocks(self) -> tuple[LayoutBlock, ...]:
        return tuple(block for page in self.pages for block in page.blocks)

    @property
    def regions(self) -> tuple[LayoutRegion, ...]:
        return tuple(region for page in self.pages for region in page.regions)


@dataclass(frozen=True)
class NativeExtraction:
    layout: DocumentLayout = field(repr=False)
    native_text_pages: tuple[str, ...] = field(repr=False)


@dataclass(frozen=True)
class TokenSpan:
    token_id: str
    start: int
    end: int


@dataclass(frozen=True)
class LayoutTextView:
    text: str = field(repr=False)
    spans: tuple[TokenSpan, ...] = field(repr=False)
    layout_version: str

    def token_ids_for_range(self, start: int, end: int) -> tuple[str, ...]:
        """Generated delimiters have no source token; token slices are exact."""
        if not 0 <= start <= end <= len(self.text):
            raise ValueError("invalid_text_range")
        return tuple(span.token_id for span in self.spans
                     if start < end and span.start < end and span.end > start)


def _bbox(tokens: Iterable[LayoutToken]) -> BBox:
    boxes = [token.bbox for token in tokens]
    return (
        min(box[0] for box in boxes), min(box[1] for box in boxes),
        max(box[2] for box in boxes), max(box[3] for box in boxes),
    )


def _same_band(anchor: LayoutToken, token: LayoutToken) -> bool:
    # The first token remains the anchor: overlapping pairs must not cause a
    # transitive chain to merge successively lower lines into one row.
    first, second = anchor.bbox, token.bbox
    height = min(first[3] - first[1], second[3] - second[1])
    overlap = min(first[3], second[3]) - max(first[1], second[1])
    center_distance = abs((first[1] + first[3] - second[1] - second[3]) / 2)
    return overlap >= height * 0.6 and center_distance <= height * 0.35


def _segments(tokens: list[LayoutToken]) -> tuple[LayoutSegment, ...]:
    groups: list[list[LayoutToken]] = []
    for token in sorted(tokens, key=lambda item: (item.bbox[0], item.bbox[1], item.token_id)):
        if groups:
            previous = groups[-1][-1]
            gap = token.bbox[0] - previous.bbox[2]
            height = min(token.bbox[3] - token.bbox[1], previous.bbox[3] - previous.bbox[1])
            if (
                (token.block, token.line) == (previous.block, previous.line)
                and -0.1 * height <= gap <= 2 * height
            ):
                groups[-1].append(token)
                continue
        groups.append([token])
    return tuple(LayoutSegment(
        tuple(t.token_id for t in group), _bbox(group),
        tuple(sorted({(t.block, t.line) for t in group})),
    ) for group in groups)


def _x_channels(rows: list[LayoutRow], gap: float) -> list[tuple[float, float]]:
    return _merge_intervals([segment.bbox[::2] for row in rows for segment in row.segments], gap)


def _merge_intervals(intervals: list[tuple[float, float]], gap: float) -> list[tuple[float, float]]:
    channels: list[tuple[float, float]] = []
    for left, right in sorted(intervals):
        if channels and left - channels[-1][1] < gap:
            channels[-1] = (channels[-1][0], max(right, channels[-1][1]))
        else:
            channels.append((left, right))
    return channels


def _reading_groups(rows: tuple[LayoutRow, ...]) -> tuple[LayoutReadingGroup, ...]:
    """Keep persistent empty gutters; consume a channel before its neighbour.

    A spanning row closes a zone rather than bridging its columns. Isolated
    tokens never participate in gutter detection. There is deliberately no
    table/party inference; row and segment IDs remain available in the layout.
    """
    normal = [row for row in rows if not row.isolation_reasons]
    isolated = [row for row in rows if row.isolation_reasons]
    groups = []
    index = 0
    zone = 0
    while index < len(normal):
        zone += 1
        selection = [normal[index]]
        # A fixed scale per zone avoids shrinking a gutter threshold mid-run.
        gap = max(12.0, 1.5 * median(s.bbox[3] - s.bbox[1] for s in selection[0].segments))
        channels = _x_channels(selection, gap)
        end = index + 1
        while end < len(normal):
            following = normal[end]
            row_height = max(selection[-1].bbox[3] - selection[-1].bbox[1], following.bbox[3] - following.bbox[1])
            if following.bbox[1] - selection[-1].bbox[3] > 3 * row_height:
                break
            combined = _merge_intervals([*channels, *(s.bbox[::2] for s in following.segments)], gap)
            if len(combined) < 2:
                break
            selection.append(following)
            channels = combined
            end += 1
        if len(channels) > 1 and len(selection) > 1:
            for column, (left, right) in enumerate(channels, 1):
                channel_rows = []
                for row in selection:
                    segments = tuple(s for s in row.segments if left <= s.bbox[0] and s.bbox[2] <= right)
                    if segments:
                        channel_rows.append(replace(
                            row, segments=segments,
                            bbox=(min(s.bbox[0] for s in segments), min(s.bbox[1] for s in segments),
                                  max(s.bbox[2] for s in segments), max(s.bbox[3] for s in segments)),
                            source_lines=tuple(sorted({source for s in segments for source in s.source_lines})),
                        ))
                groups.append(LayoutReadingGroup("column", zone, column, tuple(channel_rows)))
        else:
            groups.append(LayoutReadingGroup("rows", zone, 0, tuple(selection)))
        index = end
    # Exceptional words are fully preserved but cannot imply proximity links
    # with otherwise ordinary horizontal content (or with each other).
    for row in isolated:
        zone += 1
        groups.append(LayoutReadingGroup("isolated", zone, 0, (row,)))
    return tuple(groups)


def _outside(box: BBox, bounds: BBox, tolerance: float = 0.0) -> bool:
    return (box[0] < bounds[0] - tolerance or box[1] < bounds[1] - tolerance
            or box[2] > bounds[2] + tolerance or box[3] > bounds[3] + tolerance)


def _geometry_flags(tokens: list[LayoutToken], width: float, height: float, media_bounds: BBox):
    flags = {token.token_id: set() for token in tokens}
    active: list[LayoutToken] = []
    overlaps = duplicates = comparisons = 0
    for token in sorted(tokens, key=lambda t: (t.bbox[1], t.bbox[0], t.token_id)):
        own = flags[token.token_id]
        crop_bounds = (0.0, 0.0, width, height)
        if _outside(token.bbox, crop_bounds):
            own.add("outside_cropbox" if _outside(token.bbox, crop_bounds, BOUNDARY_TOLERANCE)
                    else "boundary_roundoff")
        if _outside(token.bbox, media_bounds, BOUNDARY_TOLERANCE):
            own.add("outside_mediabox")
        if token.direction is None or token.writing_mode is None:
            own.add("unknown_orientation")
        elif token.writing_mode != 0 or abs(token.direction[0] - 1) > 1e-4 or abs(token.direction[1]) > 1e-4:
            own.add("nonhorizontal_orientation")
        active = [other for other in active if other.bbox[3] > token.bbox[1] + OVERLAP_TOLERANCE]
        for other in active:
            comparisons += 1
            if comparisons > MAX_GEOMETRY_COMPARISONS:
                raise LayoutError("geometry_work_limit_exceeded")
            x_overlap = min(token.bbox[2], other.bbox[2]) - max(token.bbox[0], other.bbox[0])
            y_overlap = min(token.bbox[3], other.bbox[3]) - max(token.bbox[1], other.bbox[1])
            if x_overlap > OVERLAP_TOLERANCE and y_overlap > OVERLAP_TOLERANCE:
                overlaps += 1
                own.add("overlap")
                flags[other.token_id].add("overlap")
                # Ascender/descender boxes can intersect between distinct
                # lines. They remain flagged, but are already separate rows;
                # only same-band overlap can create a false horizontal link.
                kind = "same_band_overlap" if _same_band(token, other) else "cross_band_overlap"
                own.add(kind)
                flags[other.token_id].add(kind)
                if all(abs(a - b) <= OVERLAP_TOLERANCE for a, b in zip(token.bbox, other.bbox)):
                    duplicates += 1
                    own.add("duplicate_geometry")
                    flags[other.token_id].add("duplicate_geometry")
        active.append(token)
    return [replace(t, geometry_flags=tuple(sorted(flags[t.token_id]))) for t in tokens], overlaps, duplicates


def build_layout_page(
    words: Iterable[tuple], *, page: int, width: float, height: float, rotation: int = 0,
    line_directions: dict | None = None, cropbox: BBox | None = None,
    mediabox: BBox | None = None, media_bounds: BBox | None = None,
) -> LayoutPage:
    """Keep every native word; reject corrupt geometry instead of dropping it."""
    if page < 1 or not all(math.isfinite(n) and n > 0 for n in (width, height)):
        raise LayoutError("invalid_page_geometry")
    if rotation not in (0, 90, 180, 270):
        raise LayoutError("invalid_page_rotation")
    for bounds in (cropbox, mediabox, media_bounds):
        if bounds is not None and (
            len(bounds) != 4 or not all(math.isfinite(value) for value in bounds)
            or bounds[2] <= bounds[0] or bounds[3] <= bounds[1]
        ):
            raise LayoutError("invalid_page_geometry")
    records = []
    for word in words:
        if len(records) >= MAX_WORDS:
            raise LayoutError("word_limit_exceeded")
        try:
            x0, y0, x1, y1, text, block, line, index = word
            box = tuple(float(value) for value in (x0, y0, x1, y1))
            if (
                not isinstance(text, str) or not text
                or not all(math.isfinite(value) for value in box)
                or box[2] <= box[0] or box[3] <= box[1]
                or any(type(value) is not int or value < 0 for value in (block, line, index))
            ):
                raise ValueError
        except (TypeError, ValueError, OverflowError):
            raise LayoutError("invalid_native_word") from None
        records.append((block, line, index, box, text))

    # Stable source IDs, independent of the ordering of get_text(words).
    occurrences: Counter = Counter()
    tokens = []
    for block, line, index, box, text in sorted(records):
        key = (block, line, index)
        occurrences[key] += 1
        token_id = f"p{page}:b{block}:l{line}:w{index}:n{occurrences[key]}"
        direction, writing_mode = (line_directions or {}).get((block, line), (None, None))
        if (not isinstance(direction, (tuple, list)) or len(direction) != 2
                or not all(isinstance(v, (int, float)) and math.isfinite(v) for v in direction)
                or abs(math.hypot(*direction) - 1) > 1e-3 or writing_mode not in (0, 1)):
            direction, writing_mode = None, None
        tokens.append(LayoutToken(
            token_id, page, box, text, block, line, index,
            tuple(direction) if direction is not None else None, writing_mode,
            (box[0] / width, box[1] / height, box[2] / width, box[3] / height),
        ))

    cropbox = cropbox or (0.0, 0.0, width, height)
    mediabox = mediabox or cropbox
    media_bounds = media_bounds or (0.0, 0.0, width, height)
    tokens, overlaps, duplicates = _geometry_flags(tokens, width, height, media_bounds)

    bands: list[list[LayoutToken]] = []
    active: list[list[LayoutToken]] = []
    diagnostics = set()
    isolated = []
    for token in sorted(tokens, key=lambda item: (item.bbox[1], item.bbox[0], item.token_id)):
        if token.isolation_reasons:
            isolated.append(token)
            continue
        active = [band for band in active if band[0].bbox[3] >= token.bbox[1]]
        candidates = [band for band in active if _same_band(band[0], token)]
        if len(candidates) == 1:
            candidates[0].append(token)
        else:
            if candidates:
                diagnostics.add("row_association_ambiguous")
            band = [token]
            bands.append(band)
            active.append(band)
    rows = tuple(sorted([
        LayoutRow(page, _segments(band), _bbox(band), tuple(sorted({(t.block, t.line) for t in band})))
        for band in bands
    ] + [LayoutRow(page, _segments([t]), t.bbox, ((t.block, t.line),), t.isolation_reasons)
         for t in isolated], key=lambda row: (row.bbox[1], row.bbox[0], row.token_ids)))
    if rotation:
        diagnostics.add("rotated_page")
    if not tokens:
        diagnostics.add("no_native_words")
    if any(count > 1 for count in occurrences.values()):
        diagnostics.add("duplicate_source_positions")
    if any(len(row.segments) > 1 for row in rows):
        diagnostics.add("multiple_segments_in_row")
    source_line_rows = Counter(source for row in rows for source in row.source_lines)
    if any(count > 1 for count in source_line_rows.values()):
        diagnostics.add("source_line_split_across_rows")
    if any("outside_cropbox" in t.geometry_flags for t in tokens):
        diagnostics.add("words_outside_page")
    if overlaps:
        diagnostics.add("overlapping_words")
    diagnostics.update(flag for token in tokens for flag in token.geometry_flags)
    by_block: dict[int, list[LayoutToken]] = {}
    for token in tokens:
        by_block.setdefault(token.block, []).append(token)
    blocks = tuple(LayoutBlock(page, number, tuple(t.token_id for t in group), _bbox(group))
                   for number, group in sorted(by_block.items()))
    return LayoutPage(
        page, width, height, rotation, tuple(tokens), rows, tuple(sorted(diagnostics)), blocks,
        writing_direction_verified=bool(tokens) and all(t.direction is not None for t in tokens),
        cropbox=cropbox, mediabox=mediabox, media_bounds=media_bounds,
        overlap_pairs=overlaps, duplicate_geometry_pairs=duplicates,
    )


def extract_document_layout(pdf_bytes: bytes) -> DocumentLayout:
    """One words extraction per page. No OCR, model, storage or persistence."""
    return extract_native_document(pdf_bytes, include_native_text=False).layout


def extract_native_document(pdf_bytes: bytes, *, include_native_text: bool = True) -> NativeExtraction:
    """Both diagnostic projections can share one TextPage and identical flags.

    Native text is a comparison reference only, not part of the canonical data.
    No TextPage or PDF object survives the extraction call.
    """
    import fitz

    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
            if document.needs_pass:
                raise LayoutError("encrypted_pdf")
            if len(document) > MAX_PAGES:
                raise LayoutError("page_limit_exceeded")
            pages = []
            native_text_pages = []
            count = 0
            for number, pdf_page in enumerate(document, 1):
                text_page = pdf_page.get_textpage(flags=fitz.TEXTFLAGS_WORDS)
                words = pdf_page.get_text("words", sort=False, textpage=text_page)
                # The same native TextPage supplies line direction only. Its
                # span text is neither an alternative source nor persisted.
                metadata = text_page.extractDICT()
                native_lines = {
                    (block["number"], index): line
                    for block in metadata["blocks"] if block.get("type") == 0
                    for index, line in enumerate(block.get("lines", ()))
                }
                directions = {}
                for word in words:
                    key = (word[5], word[6])
                    line = native_lines.get(key)
                    if (not line or "bbox" not in line
                            or _outside(word[:4], line["bbox"], BOUNDARY_TOLERANCE)):
                        directions[key] = (None, None)
                    else:
                        # A conflicting member invalidates the mapping for the
                        # complete source line rather than guessing a direction.
                        directions.setdefault(key, (line.get("dir"), line.get("wmode")))
                del metadata, native_lines
                if include_native_text:
                    native_text_pages.append(pdf_page.get_text("text", sort=True, textpage=text_page))
                del text_page
                count += len(words)
                if count > MAX_WORDS:
                    raise LayoutError("word_limit_exceeded")
                pages.append(build_layout_page(
                    words, page=number, width=pdf_page.cropbox.width,
                    height=pdf_page.cropbox.height, rotation=pdf_page.rotation,
                    line_directions=directions, cropbox=tuple(pdf_page.cropbox),
                    mediabox=tuple(pdf_page.mediabox),
                    media_bounds=(
                        pdf_page.mediabox.x0 - pdf_page.cropbox.x0,
                        pdf_page.mediabox.y0 - pdf_page.cropbox.y0,
                        pdf_page.mediabox.x1 - pdf_page.cropbox.x0,
                        pdf_page.mediabox.y1 - pdf_page.cropbox.y0,
                    ),
                ))
            return NativeExtraction(DocumentLayout(tuple(pages)), tuple(native_text_pages))
    except LayoutError:
        raise
    except Exception:
        raise LayoutError("pdf_extraction_failed") from None


def serialize_document_layout(layout: DocumentLayout) -> LayoutTextView:
    """A lossless token projection, NOT guaranteed semantic reading order.

    Explicit column channels precede isolated anomalous tokens. Tabs separate
    segments, newlines separate rows inside a channel. Native text is never
    normalized, deduplicated, HTML-decoded, numerically repaired or truncated.
    Every character of every token is traceable through the returned spans.
    """
    chunks: list[str] = []
    spans = []
    offset = 0
    seen = set()
    all_ids = [token.token_id for page in layout.pages for token in page.tokens]
    if len(all_ids) != len(set(all_ids)):
        raise LayoutError("duplicate_token_ids")

    def append(value: str) -> None:
        nonlocal offset
        chunks.append(value)
        offset += len(value)

    for index, page in enumerate(layout.pages):
        if index:
            append("\n\n")
        append(f"[P\u00c1GINA {page.page}]\n")
        if page.rotation:
            append(f"[PAGE_ROTATION {page.rotation}]\n")
        tokens = {token.token_id: token for token in page.tokens}
        for group_index, group in enumerate(page.reading_groups):
            if group_index:
                append("\n")
            if group.kind == "column":
                append(f"[COLUMN {group.zone}:{group.column}]\n")
            elif group.kind == "isolated":
                append("[ISOLATED " + ",".join(group.rows[0].isolation_reasons) + "]\n")
            for row_index, row in enumerate(group.rows):
                if row_index:
                    append("\n")
                for segment_index, segment in enumerate(row.segments):
                    if segment_index:
                        append("\t")
                    for token_index, token_id in enumerate(segment.token_ids):
                        if token_id in seen or token_id not in tokens:
                            raise LayoutError("invalid_token_projection")
                        if token_index:
                            append(" ")
                        token = tokens[token_id]
                        spans.append(TokenSpan(token_id, offset, offset + len(token.text)))
                        append(token.text)
                        seen.add(token_id)
            if group.kind == "column":
                append("\n[/COLUMN]")
    if seen != set(all_ids):
        raise LayoutError("incomplete_token_projection")
    return LayoutTextView("".join(chunks), tuple(spans), layout.version)
