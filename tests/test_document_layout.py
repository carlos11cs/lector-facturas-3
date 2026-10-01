import ast
from collections import Counter
from dataclasses import FrozenInstanceError, replace
import io
import json
from pathlib import Path
import random
import tempfile
import unittest
from unittest.mock import patch

import fitz

from scripts import invoice_layout_benchmark as benchmark
from services import document_layout as layout_service
from services import ai_invoice_service as invoice_service


def word(text, x=10, y=10, *, block=0, line=0, index=0, width=30, height=10):
    return (x, y, x + width, y + height, text, block, line, index)


def page(words, **kwargs):
    # Synthetic words explicitly describe horizontal writing. Missing native
    # direction must NOT acquire this assumption in the production-independent builder.
    kwargs.setdefault("line_directions", {(w[5], w[6]): ((1.0, 0.0), 0) for w in words if len(w) == 8})
    return layout_service.build_layout_page(words, page=1, width=600, height=800, **kwargs)


def pdf_bytes(text="FACTURA F-1\nBASE 100,00 EUR\nIVA 21,00 EUR\nTOTAL 121,00 EUR"):
    with fitz.open() as document:
        document.new_page().insert_text((72, 72), text)
        return document.tobytes()


class TestCanonicalDocumentLayout(unittest.TestCase):
    def test_original_tokens_and_coordinates_are_immutable(self):
        raw = word("&amp; 1.234,56\u20ac", width=105)
        result = page([raw])
        token = result.tokens[0]
        self.assertEqual(token.text, raw[4])
        self.assertEqual(token.bbox, tuple(float(n) for n in raw[:4]))
        self.assertEqual((token.page, token.block, token.line, token.word), (1, 0, 0, 0))
        with self.assertRaises(FrozenInstanceError):
            token.text = "modified"

    def test_each_native_character_has_a_lossless_text_span(self):
        values = ["N.\u00ba", "FACTURA", "2025IR", "0625", "-1.234,56", "\u20ac", "\ufb01", "&amp;"]
        result = page([word(value, x=10 + 40*i, index=i) for i, value in enumerate(values)])
        view = layout_service.serialize_document_layout(layout_service.DocumentLayout((result,)))
        tokens = {t.token_id: t for t in result.tokens}
        self.assertEqual(len(view.spans), len(values))
        for span in view.spans:
            self.assertEqual(view.text[span.start:span.end], tokens[span.token_id].text)
        self.assertEqual(Counter(t.text for t in result.tokens), Counter(values))
        self.assertIn("2025IR 0625", view.text)
        self.assertIn("&amp;", view.text)

    def test_repeated_lines_and_repeated_source_positions_are_not_removed(self):
        result = page([word("TOTAL"), word("TOTAL"), word("TOTAL", y=40, line=1)])
        view = layout_service.serialize_document_layout(layout_service.DocumentLayout((result,)))
        self.assertEqual(view.text.count("TOTAL"), 3)
        self.assertEqual(len({t.token_id for t in result.tokens}), 3)
        self.assertIn("duplicate_source_positions", result.diagnostics)
        self.assertIn("overlapping_words", result.diagnostics)

    def test_geometry_is_independent_of_incoming_word_order(self):
        words = [word("F-1", x=150, index=1), word("FACTURA"), word("TOTAL", y=40, line=1)]
        expected = page(words)
        for seed in range(8):
            shuffled = words[:]
            random.Random(seed).shuffle(shuffled)
            self.assertEqual(page(shuffled), expected)

    def test_columns_stay_separate_even_with_same_source_line(self):
        result = page([word("1.234", x=10), word("56,78", x=300, index=1)])
        self.assertEqual(len(result.rows), 1)
        self.assertEqual(len(result.rows[0].segments), 2)
        view = layout_service.serialize_document_layout(layout_service.DocumentLayout((result,)))
        self.assertIn("1.234\t56,78", view.text)
        self.assertNotIn("1.234,56", view.text)

    def test_distinct_party_source_lines_are_not_joined_into_a_segment(self):
        result = page([word("PROVEEDOR"), word("CLIENTE", x=50, block=1)])
        self.assertEqual(len(result.rows[0].segments), 2)

    def test_source_line_metadata_does_not_override_vertical_geometry(self):
        result = page([word("FACTURA"), word("PEDIDO", y=70, index=1)])
        self.assertEqual(len(result.rows), 2)

    def test_row_anchor_does_not_drift_with_overlapping_tokens(self):
        result = page([word(str(i), x=10+40*i, y=10+3*i, index=i) for i in range(5)])
        self.assertEqual(len(result.rows), 3)

    def test_two_plausible_bands_are_reported_not_arbitrarily_selected(self):
        result = page([
            word("A1", x=10, y=0, height=10),
            word("A2", x=60, y=0, height=18, index=1),
            word("A3", x=110, y=1, height=12, index=2),
        ])
        self.assertIn("row_association_ambiguous", result.diagnostics)
        self.assertEqual(len(result.rows), 3)

    def test_invalid_geometry_fails_instead_of_dropping_words(self):
        for raw in [word("F-1", width=0), word("F-1", x=float("nan")), word(""), (1, 2)]:
            with self.subTest(raw_shape=len(raw)):
                with self.assertRaisesRegex(layout_service.LayoutError, "invalid_native_word"):
                    page([raw])
        with self.assertRaisesRegex(layout_service.LayoutError, "invalid_page_geometry"):
            page([], media_bounds=(0, 0, 0, 800))

    def test_rotation_and_empty_pages_are_explicit(self):
        result = page([], rotation=90)
        self.assertEqual(result.rotation, 90)
        self.assertIn("rotated_page", result.diagnostics)
        self.assertIn("no_native_words", result.diagnostics)
        self.assertIn("[P\u00c1GINA 1]", layout_service.serialize_document_layout(layout_service.DocumentLayout((result,))).text)

    def test_no_truncation_and_page_boundaries_survive(self):
        first = page([word("X" * 31000)])
        second = layout_service.build_layout_page([word("SECOND")], page=2, width=600, height=800)
        view = layout_service.serialize_document_layout(layout_service.DocumentLayout((first, second)))
        self.assertIn("X" * 31000, view.text)
        self.assertIn("[P\u00c1GINA 2]", view.text)
        self.assertEqual(len(view.spans), 2)

    def test_incomplete_or_duplicate_projection_is_rejected(self):
        original = page([word("F-1")])
        for broken in [replace(original, rows=()), replace(original, rows=original.rows * 2)]:
            with self.assertRaises(layout_service.LayoutError):
                layout_service.serialize_document_layout(layout_service.DocumentLayout((broken,)))
        with self.assertRaisesRegex(layout_service.LayoutError, "duplicate_token_ids"):
            layout_service.serialize_document_layout(layout_service.DocumentLayout((original, original)))

    def test_repr_does_not_expose_document_text(self):
        result = page([word("SENSITIVE_DOCUMENT_VALUE")])
        document = layout_service.DocumentLayout((result,))
        view = layout_service.serialize_document_layout(document)
        for value in (result.tokens[0], result, document, view):
            self.assertNotIn("SENSITIVE_DOCUMENT_VALUE", repr(value))

    def test_pdf_extractor_only_requests_words_once_per_page(self):
        payload = pdf_bytes()
        calls = []
        original = fitz.Page.get_text

        def checked(page, option="text", **kwargs):
            calls.append(option)
            return original(page, option, **kwargs)

        with patch.object(fitz.Page, "get_text", checked), patch(
            "socket.socket.connect", side_effect=AssertionError("Network forbidden")
        ):
            result = layout_service.extract_document_layout(payload)
        self.assertEqual(calls, ["words"])
        self.assertGreater(len(result.pages[0].tokens), 0)

    def test_encrypted_and_corrupt_pdfs_return_only_safe_codes(self):
        with fitz.open() as document:
            document.new_page()
            encrypted = document.tobytes(encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw="owner", user_pw="user")
        for payload, code in [(encrypted, "encrypted_pdf"), (b"SECRET CONTENT", "pdf_extraction_failed")]:
            with self.assertRaisesRegex(layout_service.LayoutError, code):
                layout_service.extract_document_layout(payload)

    def test_document_resource_limits(self):
        with patch.object(layout_service, "MAX_WORDS", 1):
            with self.assertRaisesRegex(layout_service.LayoutError, "word_limit_exceeded"):
                page([word("A1"), word("A2", x=60, index=1)])
        with patch.object(layout_service, "MAX_PAGES", 0):
            with self.assertRaisesRegex(layout_service.LayoutError, "page_limit_exceeded"):
                layout_service.extract_document_layout(pdf_bytes())


class TestOfflineLayoutBenchmark(unittest.TestCase):
    def test_legacy_baseline_matches_production_preparation(self):
        payload = pdf_bytes("FACTURA F-1\n" + "\n".join(f"Documento digital de ejemplo {i}." for i in range(12)))
        with fitz.open(stream=payload, filetype="pdf") as document:
            raw = [p.get_text("text", sort=True) for p in document]
        prepared = invoice_service.prepare_invoice_v2_fast_text(payload, filename="demo.pdf")
        self.assertEqual(benchmark.current_text_view(raw), prepared["text"])

    def test_comparison_counts_not_actual_values(self):
        result = benchmark.compare_text_views("FACTURA PRIVATE123\n01/09/2026 123,45", "FACTURA OTHER123")
        self.assertEqual(result["identifier_lexemes_missing"], 1)
        self.assertEqual(result["date_lexemes_missing"], 1)
        self.assertEqual(result["money_lexemes_missing"], 1)
        self.assertNotIn("PRIVATE123", json.dumps(result))
        self.assertNotIn("123,45", json.dumps(result))

    def test_native_pdf_benchmark_never_calls_pipeline_or_network(self):
        payload = pdf_bytes()
        with patch.object(invoice_service, "analyze_invoice", side_effect=AssertionError("No V1")), patch.object(
            invoice_service, "analyze_invoice_v2_fast_text", side_effect=AssertionError("No V2")
        ), patch.object(invoice_service, "_get_client", side_effect=AssertionError("No OpenAI")), patch(
            "socket.socket.connect", side_effect=AssertionError("No network")
        ):
            result = benchmark.compare_pdf_bytes(payload)
        self.assertTrue(result["native_token_spans_preserved"])
        self.assertEqual(result["raw_nonspace_chars_missing_in_words"], 0)

    def test_selection_is_stable_and_prefers_distinct_directories(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for folder in ["a", "b"]:
                (root / folder).mkdir()
                for name in ["first.pdf", "second.PDF"]:
                    (root / folder / name).write_bytes(pdf_bytes())
            first = benchmark.select_local_pdfs(root, 2)
            self.assertEqual(first, benchmark.select_local_pdfs(root, 2))
            self.assertEqual(len({p.parent for p in first}), 2)

    def test_cli_is_read_only_and_reports_no_sensitive_content_or_paths(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "PRIVATE_FILENAME.pdf"
            path.write_bytes(pdf_bytes("PRIVATE_SUPPLIER\nSECRET123\n12345,67 EUR"))
            before = path.read_bytes()
            for json_output in (False, True):
                output = io.StringIO()
                with patch("sys.stdout", output):
                    exit_code = benchmark.main(["--root", directory] + (["--json"] if json_output else []))
                self.assertEqual(exit_code, 0)
                for secret in ["PRIVATE_FILENAME", "PRIVATE_SUPPLIER", "SECRET123", "12345,67", directory]:
                    self.assertNotIn(secret, output.getvalue())
                self.assertEqual(path.read_bytes(), before)
                self.assertEqual(list(root.iterdir()), [path])

    def test_invalid_pdf_error_does_not_expose_path_or_exception(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "SECRET.pdf"
            path.write_bytes(b"not a pdf SECRET")
            report = benchmark.build_report([path])
            self.assertEqual(report["summary"]["errors"], 1)
            self.assertNotIn("SECRET", json.dumps(report))
            self.assertNotIn(directory, json.dumps(report))

    def test_large_similarity_is_explicitly_not_measured(self):
        with patch.object(benchmark, "MAX_SIMILARITY_CHARS", 10):
            result = benchmark.compare_text_views("X" * 11, "Y" * 11)
        self.assertIsNone(result["sequence_similarity_pct"])

    def test_no_native_text_is_not_reported_as_eligible_or_accepted(self):
        result = benchmark.compare_pdf_bytes(pdf_bytes(""))
        self.assertEqual(result["words"], 0)
        self.assertIn("no_native_words", result["diagnostics"])
        self.assertNotIn("eligible", result)
        self.assertNotIn("accept_v2", json.dumps(result))

    def test_canonical_integration_is_limited_to_shadow_analysis_service(self):
        root = Path(__file__).resolve().parents[1]
        for relative in ("app.py", "worker.py"):
            tree = ast.parse((root / relative).read_text())
            imports = [
                node.module for node in ast.walk(tree) if isinstance(node, ast.ImportFrom)
            ] + [alias.name for node in ast.walk(tree) if isinstance(node, ast.Import) for alias in node.names]
            self.assertFalse(any("document_layout" in (name or "") for name in imports))
        service_tree = ast.parse((root / "services/ai_invoice_service.py").read_text())
        service_imports = [
            node.module
            for node in ast.walk(service_tree)
            if isinstance(node, ast.ImportFrom)
        ]
        self.assertIn("services.document_layout", service_imports)

    def test_shared_textpage_is_created_once_for_both_projections(self):
        original = fitz.Page.get_textpage
        calls = []

        def counted(page, **kwargs):
            calls.append(page.number)
            return original(page, **kwargs)

        with patch.object(fitz.Page, "get_textpage", counted):
            extraction = layout_service.extract_native_document(pdf_bytes())
        self.assertEqual(calls, [0])
        self.assertEqual(len(extraction.native_text_pages), 1)
        self.assertTrue(extraction.layout.tokens)

    def test_legacy_money_stage_deltas_account_for_every_difference(self):
        # A valid integer amount followed by a percentage-like value can be
        # joined by the legacy OCR repair across a newline. Canonical data is
        # intentionally NOT passed through that repair.
        result = page([word("1.234"), word("56,78", x=300, index=1)])
        document = layout_service.DocumentLayout((result,))
        raw_pages = ["1.234\n56,78"]
        self.assertEqual(invoice_service._normalize_ocr_amount_text(raw_pages[0]), "1.234,56,78")
        diagnostics = benchmark.diagnose_legacy_money_changes(raw_pages, document)
        self.assertTrue(diagnostics["all_accounted_for"])
        for difference in diagnostics["differences"]:
            self.assertEqual(sum(difference["stage_deltas"].values()), difference["missing_occurrences"])
        self.assertIn("1.234\t56,78", layout_service.serialize_document_layout(document).text)

    def test_unknown_layout_error_cannot_leak_private_text(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "demo.pdf"
            path.write_bytes(b"pdf")
            with patch.object(benchmark, "compare_pdf_bytes", side_effect=layout_service.LayoutError("SECRET")):
                report = benchmark.build_report([path])
            self.assertNotIn("SECRET", json.dumps(report))


class TestAdversarialLayoutPreservation(unittest.TestCase):
    def assert_preserved(self, layout):
        view = layout_service.serialize_document_layout(layout)
        tokens = {token.token_id: token for token in layout.tokens}
        self.assertEqual(len(view.spans), len(tokens))
        self.assertEqual(Counter(span.token_id for span in view.spans), Counter(tokens.keys()))
        for span in view.spans:
            self.assertEqual(view.text[span.start:span.end], tokens[span.token_id].original_text)
        return view

    def test_borderless_three_column_table_and_two_tax_rates(self):
        words = []
        for row_index, values in enumerate([
            ("TIPO", "BASE", "CUOTA"), ("21%", "100,00", "21,00"), ("10%", "100,00", "10,00")
        ]):
            for column, value in enumerate(values):
                words.append(word(value, x=20+170*column, y=20+20*row_index, block=column, line=row_index))
        document = layout_service.DocumentLayout((page(words),))
        view = self.assert_preserved(document)
        self.assertEqual(len(document.rows), 3)
        self.assertTrue(all(len(row.segments) == 3 for row in document.rows))
        self.assertEqual(view.text.count("100,00"), 2)
        self.assertEqual(document.regions, ())
        self.assertEqual(len(document.blocks), 3)
        self.assertEqual(document.rows[0].source_lines, ((0, 0), (1, 0), (2, 0)))

    def test_base_vat_total_in_independent_columns(self):
        words = [word(value, x=10+180*column, y=10+20*row, block=column, line=row)
                 for row, values in enumerate([("BASE", "IVA", "TOTAL"), ("100,00", "21,00", "121,00")])
                 for column, value in enumerate(values)]
        document = layout_service.DocumentLayout((page(words),))
        self.assert_preserved(document)
        self.assertEqual(len(document.rows[0].segments), 3)
        self.assertEqual(len(document.rows[1].segments), 3)

    def test_split_amount_separate_currency_tax_id_and_hyphenated_identifier(self):
        values = ["A-2026/001", "ES", "B-12345678", "1", "234,56", "\u20ac", "07/08/26"]
        words = [word(value, x=10+40*i, index=i) for i, value in enumerate(values)]
        document = layout_service.DocumentLayout((page(words),))
        view = self.assert_preserved(document)
        self.assertIn("A-2026/001", view.text)
        self.assertIn("ES B-12345678", view.text)
        self.assertIn("1 234,56 \u20ac", view.text)
        self.assertIn("07/08/26", view.text)

    def test_repeated_headers_on_separate_pages_keep_distinct_token_ids(self):
        with fitz.open() as document:
            for index in range(2):
                p = document.new_page()
                p.insert_text((72, 72), "FACTURA F-1\nTOTAL 121,00 EUR")
            payload = document.tobytes()
        layout = layout_service.extract_document_layout(payload)
        view = self.assert_preserved(layout)
        self.assertEqual(view.text.count("F-1"), 2)
        self.assertEqual(view.text.count("121,00"), 2)
        self.assertEqual(len(layout.pages), 2)
        self.assertNotEqual(layout.pages[0].tokens[0].token_id, layout.pages[1].tokens[0].token_id)

    def test_rotated_text_is_preserved_with_native_direction(self):
        with fitz.open() as document:
            p = document.new_page()
            p.insert_text((200, 300), "FACTURA F-ROTATED", rotate=90)
            payload = document.tobytes()
        layout = layout_service.extract_document_layout(payload)
        view = self.assert_preserved(layout)
        self.assertIn("F-ROTATED", view.text)
        self.assertTrue(layout.pages[0].writing_direction_verified)
        self.assertTrue(all(t.direction == (0.0, -1.0) for t in layout.tokens))
        self.assertTrue(all("nonhorizontal_orientation" in t.geometry_flags for t in layout.tokens))
        self.assertIn("[ISOLATED nonhorizontal_orientation]", view.text)

    def test_tax_relevant_lexemes_preserved_in_simple_representative_fixture(self):
        values = ["F-2026/01", "B-12345678", "07/08/26", "1.234,56", "21%", "\u20ac"]
        document = layout_service.DocumentLayout((page([
            word(value, x=10+50*i, index=i) for i, value in enumerate(values)
        ]),))
        view = self.assert_preserved(document)
        preservation = benchmark._preservation("\n".join(values), view.text)
        for category in ("words", "identifiers", "dates", "monetary_strings", "percentages", "currency_symbols_codes", "tax_id_like"):
            self.assertEqual(preservation[category]["rate_pct"], 100.0)


class TestGeometricReadingSafety(unittest.TestCase):
    def project(self, result):
        document = layout_service.DocumentLayout((result,))
        view = layout_service.serialize_document_layout(document)
        expected = {t.token_id: t.original_text for t in result.tokens}
        self.assertEqual(len(view.spans), len(expected))
        for span in view.spans:
            self.assertEqual(view.text[span.start:span.end], expected[span.token_id])
            self.assertEqual(view.token_ids_for_range(span.start, span.end), (span.token_id,))
        return document, view

    def test_independent_equal_y_columns_are_not_interleaved(self):
        result = page([word(f"{col}{row}", x=10 + col * 300, y=20 + row * 20, block=col, line=row)
                       for row in range(3) for col in range(2)])
        document, view = self.project(result)
        self.assertEqual([view.text[s.start:s.end] for s in view.spans], ["00", "01", "02", "10", "11", "12"])
        self.assertTrue(all(len(row.segments) == 2 for row in result.rows))
        self.assertEqual(benchmark.geometry_metrics(document)["multicolumn_geometry_pages"], 1)
        self.assertEqual(benchmark.geometry_metrics(document)["horizontal_gap_points"]["min"], 270)

    def test_columns_survive_a_spanning_header_and_footer(self):
        words = [word("HEADER", width=500), word("FOOTER", width=500, y=120, block=9)]
        words += [word(f"{col}-{line}", x=10 + col * 300, y=40 + line * 20, block=col+1, line=line)
                  for line in range(3) for col in range(2)]
        _, view = self.project(page(words))
        self.assertEqual([view.text[s.start:s.end] for s in view.spans],
                         ["HEADER", "0-0", "0-1", "0-2", "1-0", "1-1", "1-2", "FOOTER"])

    def test_staggered_columns_and_same_source_line_stay_separate(self):
        words = [word(f"{col}-{row}", x=10 + col*300, y=20 + row*30 + col*12, index=row*2+col)
                 for row in range(3) for col in range(2)]
        _, view = self.project(page(words))
        self.assertEqual([view.text[s.start:s.end] for s in view.spans], ["0-0", "0-1", "0-2", "1-0", "1-1", "1-2"])

    def test_block_split_visual_row_keeps_provenance(self):
        result = page([word("LABEL", width=40), word("VALUE-1", x=55, block=8)])
        self.project(result)
        self.assertEqual(len(result.rows), 1)
        self.assertEqual(result.rows[0].source_lines, ((0, 0), (8, 0)))
        self.assertEqual(len(result.rows[0].segments), 2)
        self.assertEqual([t.block for t in result.tokens], [0, 8])

    def test_label_above_value_remains_two_rows(self):
        result = page([word("LABEL"), word("VALUE-1", y=30, block=3)])
        _, view = self.project(result)
        self.assertEqual(len(result.rows), 2)
        self.assertIn("LABEL\nVALUE-1", view.text)

    def test_duplicate_geometry_preserves_both_occurrences_in_isolation(self):
        result = page([word("121,00"), word("121,00", x=10.1, y=10.1, block=1), word("EUR", x=60)])
        document, view = self.project(result)
        metrics = benchmark.geometry_metrics(document)
        self.assertEqual(metrics["duplicate_geometry_tokens"], 2)
        self.assertEqual(metrics["duplicate_geometry_pairs"], 1)
        self.assertEqual(metrics["overlapping_tokens"], 2)
        self.assertEqual(view.text.count("121,00"), 2)
        self.assertNotIn("121,00 121,00", view.text)
        self.assertTrue(all(len(r.token_ids) == 1 for r in result.rows))

    def test_nonadjacent_overlapping_boxes_are_detected(self):
        result = page([word("WIDE", width=120), word("MIDDLE", x=30), word("LAST", x=90)])
        document, _ = self.project(result)
        metrics = benchmark.geometry_metrics(document)
        self.assertEqual(metrics["overlap_pairs"], 2)
        self.assertEqual(metrics["overlapping_tokens"], 3)
        self.assertEqual(metrics["duplicate_geometry_tokens"], 0)

    def test_overlapping_font_boxes_on_distinct_rows_do_not_destroy_lines(self):
        result = page([word("ROW1", height=15), word("ROW2", y=21, height=15, line=1)])
        document, view = self.project(result)
        metrics = benchmark.geometry_metrics(document)
        self.assertEqual(metrics["overlapping_tokens"], 2)
        self.assertEqual(metrics["cross_band_overlapping_tokens"], 2)
        self.assertEqual(metrics["same_band_overlapping_tokens"], 0)
        self.assertEqual(metrics["isolated_tokens"], 0)
        self.assertEqual(len(result.rows), 2)
        self.assertIn("ROW1\nROW2", view.text)

    def test_partially_outside_cropbox_is_not_a_horizontal_anchor(self):
        result = page([word("OUTSIDE", x=-1), word("INSIDE", x=70)])
        document, view = self.project(result)
        self.assertEqual(benchmark.geometry_metrics(document)["outside_cropbox_tokens"], 1)
        self.assertEqual(len(result.rows), 2)
        self.assertLess(result.tokens[0].normalized_bbox[0], 0)  # never clipped
        self.assertIn("ISOLATED", view.text)

    def test_cropbox_and_mediabox_are_distinguished(self):
        result = page([word("OUTSIDE", x=-2)], media_bounds=(-100, -50, 700, 850))
        document, _ = self.project(result)
        metrics = benchmark.geometry_metrics(document)
        self.assertEqual(metrics["outside_cropbox_tokens"], 1)
        self.assertEqual(metrics["outside_mediabox_tokens"], 0)

    def test_small_boundary_roundoff_is_explicit_and_not_clamped(self):
        result = page([word("ROUND", x=-0.001)])
        self.project(result)
        self.assertIn("boundary_roundoff", result.tokens[0].geometry_flags)
        self.assertNotIn("outside_cropbox", result.tokens[0].geometry_flags)
        self.assertEqual(result.tokens[0].bbox[0], -0.001)

    def test_rotated_page_uses_unrotated_cropbox_coordinates(self):
        with fitz.open() as document:
            p = document.new_page(width=600, height=800)
            p.insert_text((100, 100), "TEST-1")
            p.set_cropbox(fitz.Rect(50, 40, 550, 740))
            original = p.get_text("words")[0][:4]
            p.set_rotation(90)
            payload = document.tobytes()
        result = layout_service.extract_document_layout(payload).pages[0]
        self.project(result)
        self.assertEqual((result.width, result.height, result.rotation), (500, 700, 90))
        self.assertEqual(result.tokens[0].bbox, original)
        self.assertEqual(result.tokens[0].normalized_bbox, tuple(v / (500 if i%2 == 0 else 700) for i, v in enumerate(original)))
        self.assertEqual(result.tokens[0].geometry_flags, ())
        self.assertIn("rotated_page", result.diagnostics)

    def test_nonzero_mediabox_origin_does_not_flag_inside_text(self):
        with fitz.open() as document:
            p = document.new_page(width=600, height=800)
            p.set_mediabox(fitz.Rect(10, 20, 610, 820))
            p.insert_text((50, 60), "TEST-1")
            payload = document.tobytes()
        result = layout_service.extract_document_layout(payload).pages[0]
        self.project(result)
        self.assertEqual(result.tokens[0].geometry_flags, ())

    def test_unknown_direction_is_explicit_and_isolated(self):
        result = page([word("UNKNOWN"), word("OTHER", x=60, index=1)], line_directions={})
        document, view = self.project(result)
        self.assertFalse(result.writing_direction_verified)
        self.assertEqual(benchmark.geometry_metrics(document)["unknown_orientation_tokens"], 2)
        self.assertEqual(view.text.count("[ISOLATED unknown_orientation]"), 2)

    def test_invalid_direction_is_unknown_not_horizontal(self):
        for direction in (None, (0, 0), (float("nan"), 0), (2, 0)):
            with self.subTest(direction=direction):
                result = page([word("UNKNOWN")], line_directions={(0, 0): (direction, 0)})
                self.assertIn("unknown_orientation", result.tokens[0].geometry_flags)

    def test_vertical_text_does_not_join_horizontal_neighbours(self):
        result = page([word("HORIZONTAL"), word("VERTICAL", x=60, block=1)],
                      line_directions={(0, 0): ((1, 0), 0), (1, 0): ((0, -1), 0)})
        document, view = self.project(result)
        self.assertEqual(benchmark.geometry_metrics(document)["nonhorizontal_tokens"], 1)
        self.assertNotIn("HORIZONTAL\tVERTICAL", view.text)
        self.assertEqual(len(result.rows), 2)

    def test_vertical_writing_mode_and_reversed_direction_are_isolated(self):
        for direction, mode in (((1, 0), 1), ((-1, 0), 0)):
            result = page([word("ROTATED")], line_directions={(0, 0): (direction, mode)})
            self.assertIn("nonhorizontal_orientation", result.tokens[0].geometry_flags)

    def test_fragmented_identifier_and_separate_currency_retain_token_ranges(self):
        values = ["2025IR", "0625", "100,00", "\u20ac"]
        _, view = self.project(page([word(v, x=10+i*50, index=i) for i, v in enumerate(values)]))
        self.assertIn("2025IR 0625", view.text)
        self.assertIn("100,00 \u20ac", view.text)
        self.assertEqual(view.token_ids_for_range(0, view.spans[0].start), ())
        self.assertEqual(view.token_ids_for_range(view.spans[0].end, view.spans[1].start), ())
        self.assertEqual(view.token_ids_for_range(view.spans[0].start, view.spans[1].end),
                         (view.spans[0].token_id, view.spans[1].token_id))

    def test_extraction_uses_one_textpage_for_words_and_direction(self):
        calls = []
        original = fitz.TextPage.extractDICT
        def wrapped(text_page, **kwargs):
            result = original(text_page, **kwargs)
            calls.append(result)
            return result
        with patch.object(fitz.TextPage, "extractDICT", wrapped):
            result = layout_service.extract_document_layout(pdf_bytes())
        self.assertEqual(len(calls), 1)
        self.assertTrue(result.pages[0].writing_direction_verified)
        self.assertTrue(all(b["type"] == 0 for b in calls[0]["blocks"]))

    def test_direction_is_unknown_if_source_line_geometry_cannot_be_matched(self):
        original = fitz.TextPage.extractDICT
        def incompatible(text_page, **kwargs):
            metadata = original(text_page, **kwargs)
            for block in metadata["blocks"]:
                for line in block.get("lines", ()):
                    line["bbox"] = (0, 0, 1, 1)
            return metadata
        with patch.object(fitz.TextPage, "extractDICT", incompatible):
            result = layout_service.extract_document_layout(pdf_bytes())
        self.assertTrue(all("unknown_orientation" in t.geometry_flags for t in result.tokens))
        self.assertFalse(result.pages[0].writing_direction_verified)

    def test_geometry_work_bound_fails_explicitly_without_partial_output(self):
        with patch.object(layout_service, "MAX_GEOMETRY_COMPARISONS", 0):
            with self.assertRaisesRegex(layout_service.LayoutError, "geometry_work_limit_exceeded"):
                page([word("DUP"), word("DUP", block=1)])

    def test_new_geometry_diagnostics_expose_only_counts(self):
        result = benchmark.compare_pdf_bytes(pdf_bytes("PRIVATE123\nSECRET456 12345,67 EUR"))
        self.assertEqual(result["reversible_token_spans"], result["words"])
        self.assertEqual(result["geometry"]["unknown_orientation_tokens"], 0)
        self.assertNotIn("PRIVATE123", json.dumps(result))
        self.assertNotIn("SECRET456", json.dumps(result))
        self.assertNotIn("12345,67", json.dumps(result))


if __name__ == "__main__":
    unittest.main()
