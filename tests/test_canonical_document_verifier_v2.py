import unittest
from copy import deepcopy
from unittest.mock import patch

import app as ledger_app
from services import ai_invoice_service as invoice_service
from services.canonical_document_structure import (
    assemble_structural_document,
    build_anchor_clusters,
    build_evidence_relations,
    build_regions,
    build_semantic_segments,
    extract_field_candidates,
)
from services.canonical_document_verifier_v2 import verify_canonical_document_v2
from services.document_layout import DocumentLayout, build_layout_page


def _layout(*rows):
    words = []
    directions = {}
    for line_index, row in enumerate(rows):
        y = 72.0 + line_index * 22
        default_x = 72.0
        for word_index, item in enumerate(row):
            if isinstance(item, tuple):
                text = item[0]
                x = float(item[1])
                block = int(item[2]) if len(item) > 2 else 0
            else:
                text = item
                x = default_x
                block = 0
            width = max(12.0, len(text) * 6.0)
            words.append(
                (x, y, x + width, y + 10, text, block, line_index, word_index)
            )
            directions[(block, line_index)] = ((1.0, 0.0), 0)
            default_x = x + width + 8
    page = build_layout_page(
        words,
        page=1,
        width=595,
        height=842,
        line_directions=directions,
    )
    return DocumentLayout((page,))


def _normalized(**overrides):
    result = {
        "provider_name": "Proveedor Demo SL",
        "supplier_tax_id": "B12345678",
        "client_name": "Cliente Demo SL",
        "customer_tax_id": "B87654321",
        "invoice_number": "F-100",
        "invoice_date": "2026-08-07",
        "payment_dates": [],
        "currency": "EUR",
        "base_amount": 150.0,
        "vat_amount": 26.0,
        "withholding_amount": 0.0,
        "other_taxes": 0.0,
        "total_amount": 176.0,
        "vat_breakdown": [
            {"rate": 21.0, "base": 100.0, "vat_amount": 21.0},
            {"rate": 10.0, "base": 50.0, "vat_amount": 5.0},
        ],
    }
    result.update(overrides)
    return result


def _structure(layout, registered_company_tax_id="B87654321"):
    segments, token_by_id = build_semantic_segments(layout)
    anchors = build_anchor_clusters(segments, token_by_id)
    candidates = extract_field_candidates(segments, anchors, token_by_id)
    regions, segment_region = build_regions(
        segments,
        anchors,
        candidates,
        registered_company_tax_id=registered_company_tax_id,
    )
    relations = build_evidence_relations(
        segments, anchors, candidates, segment_region
    )
    return assemble_structural_document(
        segments,
        anchors,
        regions,
        candidates,
        relations,
        token_by_id,
        segment_region,
    )


class TestCanonicalStructuralEngine(unittest.TestCase):
    def test_a_invoice_anchor_and_value_in_adjacent_same_row_segments(self):
        layout = _layout(
            (("FACTURA", 72, 0), ("Nº", 126, 1), ("26049906", 190, 2)),
        )
        structure = _structure(layout)
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_number="26049906"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(len(structure.anchors_by_type["invoice_number"]), 1)
        self.assertEqual(result["fields"]["invoice_number"]["status"], "confirmed")
        self.assertEqual(
            result["fields"]["invoice_number"]["valid_associations"], 1
        )

    def test_b_multiline_invoice_label_is_one_anchor_cluster(self):
        layout = _layout(("FACTURA",), ("Nº",), ("26049906",))
        structure = _structure(layout)
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_number="26049906"),
            registered_company_tax_id="B87654321",
        )

        anchors = structure.anchors_by_type["invoice_number"]
        self.assertEqual(len(anchors), 1)
        self.assertEqual(len(anchors[0].token_ids), 2)
        self.assertEqual(result["fields"]["invoice_number"]["status"], "confirmed")

    def test_c_order_reference_does_not_compete_with_invoice_number(self):
        layout = _layout(
            ("FACTURA", "Nº", "26049906"),
            ("PEDIDO", "839382"),
        )
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_number="26049906"),
            registered_company_tax_id="B87654321",
        )

        invoice = result["fields"]["invoice_number"]
        self.assertEqual(invoice["status"], "confirmed")
        self.assertEqual(invoice["competing_associations"], 0)

    def test_order_local_context_dominates_invoice_anchor_above(self):
        layout = _layout(
            ("FACTURA", "Nº"),
            ("PEDIDO", "839382"),
        )
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_number="839382"),
            registered_company_tax_id="B87654321",
        )

        invoice = result["fields"]["invoice_number"]
        self.assertEqual(invoice["status"], "review")
        self.assertEqual(invoice["valid_associations"], 0)
        self.assertEqual(invoice["relation_type"], "negative_context")

    def test_d_equal_y_columns_are_not_cross_associated(self):
        layout = _layout(
            (("FACTURA", 20, 0), ("Nº", 75, 0), ("26049906", 360, 1)),
        )
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_number="26049906"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(result["fields"]["invoice_number"]["status"], "review")
        self.assertEqual(
            result["fields"]["invoice_number"]["valid_associations"], 0
        )

    def test_e_invoice_and_due_dates_keep_their_own_anchors(self):
        layout = _layout(
            ("PROVEEDOR", "NIF", "B12345678"),
            ("CLIENTE", "NIF", "B87654321"),
            ("FECHA", "FACTURA", "07/08/26"),
            ("VENCIMIENTO", "30/08/26"),
        )
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_date="2026-08-07"),
            registered_company_tax_id="B87654321",
        )

        invoice_date = result["fields"]["invoice_date"]
        self.assertEqual(result["supplier_jurisdiction"], "ES")
        self.assertEqual(invoice_date["status"], "confirmed")
        self.assertEqual(invoice_date["valid_associations"], 1)
        self.assertEqual(invoice_date["competing_associations"], 0)

    def test_due_date_local_context_dominates_invoice_date_anchor_above(self):
        layout = _layout(
            ("FECHA", "FACTURA"),
            ("VENCIMIENTO", "30/08/26"),
        )
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_date="2026-08-30"),
            registered_company_tax_id="B87654321",
        )

        invoice_date = result["fields"]["invoice_date"]
        self.assertEqual(invoice_date["status"], "review")
        self.assertEqual(invoice_date["valid_associations"], 0)
        self.assertEqual(invoice_date["relation_type"], "negative_context")

    def test_f_repeated_footer_value_does_not_create_second_association(self):
        layout = _layout(
            ("FACTURA", "Nº", "26049906"),
            ("26049906",),
        )
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_number="26049906"),
            registered_company_tax_id="B87654321",
        )

        invoice = result["fields"]["invoice_number"]
        self.assertEqual(invoice["status"], "confirmed")
        self.assertEqual(invoice["value_occurrences"], 2)
        self.assertEqual(invoice["valid_associations"], 1)

    def test_g_supplier_and_recipient_tax_ids_are_resolved_jointly(self):
        layout = _layout(
            ("PROVEEDOR", "Proveedor", "Demo", "SL"),
            ("NIF", "B12345678"),
            ("CLIENTE", "Cliente", "Demo", "SL"),
            ("NIF", "B87654321"),
        )
        result = verify_canonical_document_v2(
            layout,
            _normalized(),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(result["fields"]["supplier_tax_id"]["status"], "confirmed")
        self.assertEqual(result["fields"]["recipient_tax_id"]["status"], "confirmed")
        self.assertEqual(result["supplier_jurisdiction"], "ES")

    def test_h_multi_rate_vat_rows_do_not_reuse_candidates(self):
        layout = _layout(
            ("TIPO", "BASE", "CUOTA"),
            ("21%", "100,00", "21,00"),
            ("10%", "50,00", "5,00"),
        )
        result = verify_canonical_document_v2(
            layout,
            _normalized(),
            registered_company_tax_id="B87654321",
        )

        breakdown = result["fields"]["vat_breakdown"]
        self.assertEqual(breakdown["status"], "confirmed")
        self.assertEqual(breakdown["valid_associations"], 2)
        self.assertEqual(breakdown["competing_associations"], 0)

    def test_i_total_is_not_confused_with_subtotal(self):
        layout = _layout(
            ("SUBTOTAL", "150,00", "EUR"),
            ("TOTAL", "FACTURA", "176,00", "EUR"),
        )
        result = verify_canonical_document_v2(
            layout,
            _normalized(),
            registered_company_tax_id="B87654321",
        )

        total = result["fields"]["total_amount"]
        self.assertEqual(total["status"], "confirmed")
        self.assertEqual(total["competing_associations"], 0)

    def test_j_value_without_label_is_review_not_contradiction(self):
        layout = _layout(("26049906",))
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_number="26049906"),
            registered_company_tax_id="B87654321",
        )

        invoice = result["fields"]["invoice_number"]
        self.assertEqual(invoice["status"], "review")
        self.assertEqual(invoice["reason_code"], "invoice_number_anchor_missing")

    def test_single_explicit_incompatible_invoice_number_is_contradiction(self):
        layout = _layout(
            ("FACTURA", "Nº", "F-200"),
            ("DETALLE",),
            ("CONCEPTO",),
            ("OBSERVACIONES",),
            ("F-100",),
        )
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_number="F-100"),
            registered_company_tax_id="B87654321",
        )

        invoice = result["fields"]["invoice_number"]
        self.assertEqual(invoice["status"], "contradiction")
        self.assertEqual(invoice["reason_code"], "explicit_invoice_number_differs")

    def test_multiple_explicit_invoice_number_alternatives_are_review(self):
        layout = _layout(
            ("FACTURA", "Nº", "F-200"),
            ("FACTURA", "Nº", "F-300"),
            ("DETALLE",),
            ("CONCEPTO",),
            ("OBSERVACIONES",),
            ("F-100",),
        )
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_number="F-100"),
            registered_company_tax_id="B87654321",
        )

        invoice = result["fields"]["invoice_number"]
        self.assertEqual(invoice["status"], "review")
        self.assertEqual(
            invoice["reason_code"],
            "explicit_invoice_number_alternatives_compete",
        )
        self.assertEqual(invoice["competing_associations"], 2)

    def test_id_de_factura_is_a_strong_invoice_number_label(self):
        result = verify_canonical_document_v2(
            _layout(("ID", "DE", "FACTURA", "7af56838-58ae-4258")),
            _normalized(invoice_number="7af56838-58ae-4258"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(result["fields"]["invoice_number"]["status"], "confirmed")

    def test_contextual_factura_label_only_confirms_same_segment_value(self):
        same_segment = verify_canonical_document_v2(
            _layout(("FACTURA", "FND20579")),
            _normalized(invoice_number="FND20579"),
            registered_company_tax_id="B87654321",
        )
        distant = verify_canonical_document_v2(
            _layout(("FACTURA",), ("DETALLE",), ("FND20579",)),
            _normalized(invoice_number="FND20579"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(same_segment["fields"]["invoice_number"]["status"], "confirmed")
        self.assertEqual(distant["fields"]["invoice_number"]["status"], "review")

    def test_contextual_factura_does_not_create_explicit_alternative(self):
        result = verify_canonical_document_v2(
            _layout(
                ("NÚMERO", "DE", "FACTURA"),
                ("F-100",),
                ("Factura", "Cuota", "Julio", "2026"),
            ),
            _normalized(invoice_number="F-100"),
            registered_company_tax_id="B87654321",
        )

        invoice = result["fields"]["invoice_number"]
        self.assertEqual(invoice["status"], "confirmed")
        self.assertEqual(invoice["competing_associations"], 0)

    def test_same_segment_invoice_number_dominates_remote_below_candidate(self):
        result = verify_canonical_document_v2(
            _layout(
                ("FACTURA", "FND20579"),
                ("FACTURA", "Nº"),
                ("919195691",),
            ),
            _normalized(invoice_number="FND20579"),
            registered_company_tax_id="B87654321",
        )

        invoice = result["fields"]["invoice_number"]
        self.assertEqual(invoice["status"], "confirmed")
        self.assertEqual(
            invoice["reason_code"],
            "dominant_same_segment_invoice_number_association",
        )

    def test_repeated_same_invoice_metadata_is_consistent_not_ambiguous(self):
        result = verify_canonical_document_v2(
            _layout(
                ("Nº", "FACTURA", "INV-100"),
                ("FECHA", "FACTURA", "22/07/2026"),
                ("Nº", "FACTURA", "INV-100"),
                ("FECHA", "FACTURA", "22/07/2026"),
            ),
            _normalized(invoice_number="INV-100", invoice_date="2026-07-22"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(result["fields"]["invoice_number"]["status"], "confirmed")
        self.assertEqual(result["fields"]["invoice_date"]["status"], "confirmed")

    def test_invoice_date_is_not_invalidated_by_same_due_date_elsewhere(self):
        result = verify_canonical_document_v2(
            _layout(
                ("FECHA", "FACTURA", "16/07/2026"),
                ("VENCIMIENTO", "16/07/2026"),
            ),
            _normalized(invoice_date="2026-07-16"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(result["fields"]["invoice_date"]["status"], "confirmed")

    def test_same_date_occurrence_with_invoice_and_due_labels_is_review(self):
        result = verify_canonical_document_v2(
            _layout(("FECHA", "VENCIMIENTO", "16/07/2026")),
            _normalized(invoice_date="2026-07-16"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(result["fields"]["invoice_date"]["status"], "review")

    def test_generic_unrelated_date_is_not_explicit_contradiction(self):
        result = verify_canonical_document_v2(
            _layout(
                ("FACTURA", "F-100"),
                ("MADRID", "01/07/2026"),
                ("Fecha", "05/07/2026"),
            ),
            _normalized(invoice_date="2026-07-01"),
            registered_company_tax_id="B87654321",
        )

        invoice_date = result["fields"]["invoice_date"]
        self.assertEqual(invoice_date["status"], "review")
        self.assertNotEqual(invoice_date["reason_code"], "explicit_invoice_date_differs")

    def test_invoice_header_date_contradicts_proposed_due_date(self):
        result = verify_canonical_document_v2(
            _layout(
                ("PROVEEDOR", "NIF", "B12345678"),
                ("CLIENTE", "NIF", "B87654321"),
                (("FACTURA", 72, 0), ("Nº", 130, 0), ("FECHA", 300, 1)),
                (("F-100", 72, 0), ("09/07/2026", 300, 1)),
                ("VENCIMIENTO", "07/09/2026"),
            ),
            _normalized(invoice_number="F-100", invoice_date="2026-09-07"),
            registered_company_tax_id="B87654321",
        )

        invoice_date = result["fields"]["invoice_date"]
        self.assertEqual(invoice_date["status"], "contradiction")
        self.assertEqual(invoice_date["reason_code"], "explicit_invoice_date_differs")

    def test_invoice_date_field_label_same_row_can_confirm(self):
        layout = _layout(
            ("PROVEEDOR", "NIF", "B12345678"),
            ("CLIENTE", "NIF", "B87654321"),
            ("Invoice", "Date:", "09/07/2026"),
        )
        structure = _structure(layout)
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_date="2026-07-09"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(len(structure.anchors_by_type["invoice_date"]), 1)
        self.assertEqual(result["fields"]["invoice_date"]["status"], "confirmed")

    def test_invoice_date_field_label_above_value_can_confirm(self):
        layout = _layout(
            ("PROVEEDOR", "NIF", "B12345678"),
            ("CLIENTE", "NIF", "B87654321"),
            ("Invoice", "Date"),
            ("09/07/2026",),
        )
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_date="2026-07-09"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(result["fields"]["invoice_date"]["status"], "confirmed")

    def test_embedded_temporal_references_are_not_invoice_date_anchors(self):
        references = (
            ("NET", "60", "FROM", "INVOICE", "DATE"),
            ("PAYABLE", "BY", "45", "DAYS", "FROM", "INVOICE", "DATE"),
            ("60", "DIAS", "FECHA", "FACTURA"),
            ("PAYMENT", "TERMS", "45", "DAYS", "AFTER", "INVOICE", "DATE"),
            ("UNKNOWN", "TEMPORAL", "CONTEXT", "AROUND", "INVOICE", "DATE"),
        )

        for reference in references:
            with self.subTest(reference=reference):
                layout = _layout(reference, ("07/09/2026",))
                structure = _structure(layout)
                result = verify_canonical_document_v2(
                    layout,
                    _normalized(invoice_date="2026-09-07"),
                    registered_company_tax_id="B87654321",
                )

                self.assertEqual(
                    len(structure.anchors_by_type.get("invoice_date", ())), 0
                )
                self.assertNotEqual(
                    result["fields"]["invoice_date"]["status"], "confirmed"
                )

    def test_payment_reference_does_not_create_third_issue_date_anchor(self):
        layout = _layout(
            ("PROVEEDOR", "NIF", "B12345678"),
            ("CLIENTE", "NIF", "B87654321"),
            ("Invoice", "Date:", "09/07/2026"),
            ("Payment", "terms:", "NET", "60", "FROM", "INVOICE", "DATE"),
            ("Due", "Date:", "07/09/2026"),
        )
        structure = _structure(layout)
        baseline_structure = _structure(
            _layout(
                ("PROVEEDOR", "NIF", "B12345678"),
                ("CLIENTE", "NIF", "B87654321"),
                ("Invoice", "Date:", "09/07/2026"),
                ("Due", "Date:", "07/09/2026"),
            )
        )
        issue_result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_date="2026-07-09"),
            registered_company_tax_id="B87654321",
        )
        due_result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_date="2026-09-07"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(
            len(structure.anchors_by_type["invoice_date"]),
            len(baseline_structure.anchors_by_type["invoice_date"]),
        )
        self.assertEqual(issue_result["fields"]["invoice_date"]["status"], "confirmed")
        self.assertNotEqual(due_result["fields"]["invoice_date"]["status"], "confirmed")

    def test_spanish_payment_reference_does_not_create_issue_date_anchor(self):
        layout = _layout(
            ("PROVEEDOR", "NIF", "B12345678"),
            ("CLIENTE", "NIF", "B87654321"),
            ("Fecha", "factura:", "09/07/2026"),
            ("Condiciones", "de", "pago:", "60", "dias", "fecha", "factura"),
            ("Fecha", "vencimiento:", "07/09/2026"),
        )
        structure = _structure(layout)
        baseline_structure = _structure(
            _layout(
                ("PROVEEDOR", "NIF", "B12345678"),
                ("CLIENTE", "NIF", "B87654321"),
                ("Fecha", "factura:", "09/07/2026"),
                ("Fecha", "vencimiento:", "07/09/2026"),
            )
        )
        result = verify_canonical_document_v2(
            layout,
            _normalized(invoice_date="2026-07-09"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(
            len(structure.anchors_by_type["invoice_date"]),
            len(baseline_structure.anchors_by_type["invoice_date"]),
        )
        self.assertEqual(result["fields"]["invoice_date"]["status"], "confirmed")

    def test_temporal_reference_classification_is_parameterized(self):
        for days in (15, 37, 91):
            with self.subTest(days=days):
                layout = _layout(
                    ("CUSTOM", "TERMS", str(days), "AFTER", "INVOICE", "DATE"),
                    ("01/10/2026",),
                )
                structure = _structure(layout)

                self.assertEqual(
                    len(structure.anchors_by_type.get("invoice_date", ())), 0
                )

    def test_fecha_de_la_factura_is_a_strong_date_label(self):
        result = verify_canonical_document_v2(
            _layout(("FECHA", "DE", "LA", "FACTURA", "23/07/2026")),
            _normalized(invoice_date="2026-07-23"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(result["fields"]["invoice_date"]["status"], "confirmed")

    def test_generic_date_label_and_value_in_adjacent_same_row_segments(self):
        result = verify_canonical_document_v2(
            _layout(
                (("Fecha:", 72, 0), ("17/07/2026", 180, 1)),
                (("Vencimiento:", 72, 0), ("17/07/2026", 180, 1)),
            ),
            _normalized(invoice_date="2026-07-17"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(result["fields"]["invoice_date"]["status"], "confirmed")

    def test_month_name_invoice_dates_are_canonical_candidates(self):
        for words, expected in (
            (("FECHA", "29", "jul.", "2026"), "2026-07-29"),
            (("FECHA", "25", "de", "julio", "de", "2026"), "2026-07-25"),
            (("INVOICE", "DATE", "Jul", "01,", "2026"), "2026-07-01"),
        ):
            with self.subTest(words=words):
                result = verify_canonical_document_v2(
                    _layout(("FACTURA", "F-100"), words),
                    _normalized(invoice_date=expected),
                    registered_company_tax_id="B87654321",
                )
                self.assertEqual(result["fields"]["invoice_date"]["status"], "confirmed")

    def test_generic_date_below_invoice_header_is_confirmed(self):
        result = verify_canonical_document_v2(
            _layout(
                ("Proveedor Demo SL", "B12345678"),
                ("FACTURA", "Nº", "TIPO", "FAC", "FECHA"),
                ("26049906", "RI", "09/07/2026"),
            ),
            _normalized(invoice_number="26049906", invoice_date="2026-07-09"),
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(result["fields"]["invoice_date"]["status"], "confirmed")

    def test_tax_number_marker_is_not_part_of_vat_identifier(self):
        structure = _structure(
            _layout(("Ireland", "VAT", "Reg", "No.", "IE9700053D"))
        )

        values = {
            candidate.normalized_value
            for candidate in structure.candidates_by_type.get("tax_id", ())
        }
        self.assertEqual(values, {"IE9700053D"})


class TestCanonicalVerifierV2Integration(unittest.TestCase):
    def _run_independence_pipeline(
        self,
        *,
        legacy_invoice_number="F-100",
        legacy_invoice_date="2026-08-23",
        validate_side_effect=None,
        canonical_mutator=None,
    ):
        layout = _layout(
            ("FACTURA", "Nº", "F-100"),
            ("FECHA", "FACTURA", "23/08/2026"),
        )
        prepared = {
            "eligible": True,
            "reason": "canonical_text_sufficient",
            "input_representation": "canonical_layout_v1",
            "page_count": 1,
            "native_text_chars": 100,
            "sent_text_chars": 100,
            "document_text_complete": True,
            "text": "LEGACY VERIFIER TEXT",
            "verification_text": "LEGACY VERIFIER TEXT",
            "model_input_text": "CANONICAL MODEL TEXT",
            "_canonical_document_layout": layout,
        }
        response = {
            "supplier": {"legal_name": "Model Supplier SL", "tax_id": "B12345678"},
            "customer": {"legal_name": "Model Customer SL", "tax_id": "B87654321"},
            "invoice": {
                "invoice_number": "F-100",
                "issue_date": "2026-08-23",
                "currency": "EUR",
            },
            "due_dates": ["2026-09-01"],
            "taxes": [{"taxable_base": 100, "vat_rate": 21, "vat_amount": 21}],
            "totals": {
                "taxable_base": 100,
                "vat_amount": 21,
                "withholding": 0,
                "other_taxes": 0,
                "total": 121,
            },
            "field_evidence": {},
        }
        captured = {}

        def capture_canonical(_layout, proposal, **_kwargs):
            captured["before"] = deepcopy(proposal)
            if canonical_mutator is not None:
                canonical_mutator(proposal)
            captured["after"] = deepcopy(proposal)
            field = {
                "status": "confirmed",
                "reason_code": "test_confirmation",
                "value_occurrences": 1,
                "anchor_clusters": 1,
                "valid_associations": 1,
                "competing_associations": 0,
                "relation_type": "same_segment",
                "evidence_class": "strong_same_row",
            }
            return {
                "version": "canonical_document_verifier_v2",
                "fields": {
                    "invoice_number": dict(field),
                    "invoice_date": dict(field),
                },
                "consistency": {},
                "structure": {},
                "timings_ms": {},
            }

        def preserve_payment_dates(values, *_args, **_kwargs):
            return list(values), []

        validate = validate_side_effect or (lambda *_args, **_kwargs: [])
        legacy_verification = {
            "invoice_number": {"status": "review", "reason": "legacy_gate"},
            "invoice_date": {"status": "review", "reason": "legacy_gate"},
        }
        with patch.object(invoice_service, "_get_client", return_value=object()), patch.object(
            invoice_service, "_get_invoice_model", return_value="gpt-5.6-sol"
        ), patch.object(
            invoice_service, "_call_invoice_responses", return_value=response
        ), patch.object(
            invoice_service,
            "_reconcile_fast_text_invoice_number",
            return_value=(legacy_invoice_number, [], [], "confirmed", {}),
        ), patch.object(
            invoice_service,
            "_reconcile_fast_text_invoice_date",
            return_value=(legacy_invoice_date, [], [], {}),
        ), patch.object(
            invoice_service,
            "_reconcile_fast_text_payment_dates",
            side_effect=preserve_payment_dates,
        ), patch.object(
            invoice_service,
            "_validate_fast_text_invoice",
            side_effect=validate,
        ), patch.object(
            invoice_service,
            "_assess_fast_text_invoice_safety",
            return_value={
                "accounting_safety_status": "passed",
                "accounting_safety_issues": [],
                "metadata_quality_status": "confirmed",
                "metadata_issues": [],
            },
        ), patch.object(
            invoice_service,
            "_verify_fast_text_document",
            return_value=legacy_verification,
        ), patch.object(
            invoice_service,
            "_decide_fast_text_fast_path",
            return_value=("fallback_v1", ["legacy_gate"]),
        ), patch.object(
            invoice_service,
            "verify_canonical_document_v2",
            side_effect=capture_canonical,
        ):
            result = invoice_service.analyze_invoice_v2_fast_text(
                file_bytes=b"unused",
                filename="factura.pdf",
                mime_type="application/pdf",
                prepared_text=prepared,
            )
        return result, captured

    def test_a_canonical_receives_final_reconciled_invoice_number(self):
        result, captured = self._run_independence_pipeline(
            legacy_invoice_number="LEGACY-B"
        )

        self.assertEqual(captured["before"]["invoice_number"], "LEGACY-B")
        self.assertEqual(result["invoice_number"], "LEGACY-B")

    def test_b_canonical_receives_final_reconciled_invoice_date(self):
        result, captured = self._run_independence_pipeline(
            legacy_invoice_date="2026-09-30"
        )

        self.assertEqual(captured["before"]["invoice_date"], "2026-09-30")
        self.assertEqual(result["invoice_date"], "2026-09-30")

    def test_c_post_reconciliation_validators_cannot_mutate_final_candidate(self):
        def mutate_legacy_nested(_structured, normalized, *_args, **_kwargs):
            normalized["vat_breakdown"][0]["base"] = 999.0
            normalized["payment_dates"].append("2026-10-01")
            return []

        result, captured = self._run_independence_pipeline(
            validate_side_effect=mutate_legacy_nested
        )

        self.assertEqual(captured["before"]["vat_breakdown"][0]["base"], 100.0)
        self.assertEqual(captured["before"]["payment_dates"], ["2026-09-01"])
        self.assertEqual(result["vat_breakdown"][0]["base"], 100.0)
        self.assertEqual(result["payment_dates"], ["2026-09-01"])

    def test_d_canonical_mutation_invalidates_verdict_and_not_persisted_candidate(self):
        def mutate_canonical(proposal):
            proposal["invoice_number"] = "CANONICAL-ONLY"
            proposal["vat_breakdown"][0]["base"] = 777.0

        result, captured = self._run_independence_pipeline(
            canonical_mutator=mutate_canonical
        )

        self.assertEqual(captured["after"]["invoice_number"], "CANONICAL-ONLY")
        self.assertEqual(result["invoice_number"], "F-100")
        self.assertEqual(result["vat_breakdown"][0]["base"], 100.0)
        self.assertEqual(result["fast_path_decision"], "fallback_v1")
        canonical = result["invoice_parser_diagnostics"][
            "canonical_document_verifier_v2"
        ]
        self.assertEqual(canonical["status"], "diagnostic_error")
        self.assertEqual(canonical["fields"], {})

    def test_e_all_verified_values_equal_the_persisted_final_candidate(self):
        def mutate_legacy(_structured, normalized, *_args, **_kwargs):
            normalized["provider_name"] = "Legacy Supplier SL"
            return []

        def mutate_canonical(proposal):
            proposal["supplier_tax_id"] = "CANONICAL123"

        result, captured = self._run_independence_pipeline(
            legacy_invoice_number="FINAL-200",
            legacy_invoice_date="2026-09-30",
            validate_side_effect=mutate_legacy,
            canonical_mutator=mutate_canonical,
        )

        self.assertEqual(captured["before"]["provider_name"], "Model Supplier SL")
        self.assertEqual(captured["before"]["supplier_tax_id"], "B12345678")
        self.assertEqual(captured["after"]["supplier_tax_id"], "CANONICAL123")
        for field in invoice_service._FAST_TEXT_CANONICAL_VERIFIED_FIELDS:
            self.assertEqual(result[field], captured["before"][field], field)
        self.assertEqual(result["provider_name"], "Model Supplier SL")
        self.assertEqual(result["supplier_tax_id"], "B12345678")
        self.assertEqual(result["invoice_number"], "FINAL-200")
        self.assertEqual(result["invoice_date"], "2026-09-30")

    def test_final_issue_date_can_confirm_but_final_due_date_cannot(self):
        layout = _layout(
            ("PROVEEDOR", "NIF", "B12345678"),
            ("CLIENTE", "NIF", "B87654321"),
            (("FACTURA", 72, 0), ("Nº", 130, 0), ("FECHA", 300, 1)),
            (("F-100", 72, 0), ("09/07/2026", 300, 1)),
            ("VENCIMIENTO", "07/09/2026"),
        )
        prepared = {
            "eligible": True,
            "reason": "canonical_text_sufficient",
            "input_representation": "canonical_layout_v1",
            "page_count": 1,
            "native_text_chars": 100,
            "sent_text_chars": 100,
            "document_text_complete": True,
            "text": "LEGACY VERIFIER TEXT",
            "verification_text": "LEGACY VERIFIER TEXT",
            "model_input_text": "CANONICAL MODEL TEXT",
            "_canonical_document_layout": layout,
        }
        response = {
            "supplier": {"legal_name": "Proveedor Demo SL", "tax_id": "B12345678"},
            "customer": {"legal_name": "Cliente Demo SL", "tax_id": "B87654321"},
            "invoice": {
                "invoice_number": "F-100",
                "issue_date": "2026-07-09",
                "currency": "EUR",
            },
            "due_dates": ["2026-09-07"],
            "taxes": [{"taxable_base": 100, "vat_rate": 21, "vat_amount": 21}],
            "totals": {
                "taxable_base": 100,
                "vat_amount": 21,
                "withholding": 0,
                "other_taxes": 0,
                "total": 121,
            },
            "field_evidence": {},
        }
        safety = {
            "accounting_safety_status": "passed",
            "accounting_safety_issues": [],
            "metadata_quality_status": "confirmed",
            "metadata_issues": [],
        }
        legacy_verification = {
            "invoice_number": {"status": "review", "reason": "legacy_gate"},
            "invoice_date": {"status": "review", "reason": "legacy_gate"},
        }

        for final_date, expected_status in (
            ("2026-07-09", "confirmed"),
            ("2026-09-07", "contradiction"),
        ):
            with self.subTest(final_date=final_date), patch.object(
                invoice_service, "_get_client", return_value=object()
            ), patch.object(
                invoice_service, "_get_invoice_model", return_value="gpt-5.6-sol"
            ), patch.object(
                invoice_service, "_call_invoice_responses", return_value=response
            ), patch.object(
                invoice_service,
                "_reconcile_fast_text_invoice_number",
                return_value=("F-100", [], [], "confirmed", {}),
            ), patch.object(
                invoice_service,
                "_reconcile_fast_text_invoice_date",
                return_value=(final_date, [], [], {}),
            ), patch.object(
                invoice_service,
                "_reconcile_fast_text_payment_dates",
                return_value=(["2026-09-07"], []),
            ), patch.object(
                invoice_service, "_validate_fast_text_invoice", return_value=[]
            ), patch.object(
                invoice_service, "_assess_fast_text_invoice_safety", return_value=safety
            ), patch.object(
                invoice_service,
                "_verify_fast_text_document",
                return_value=legacy_verification,
            ), patch.object(
                invoice_service,
                "_decide_fast_text_fast_path",
                return_value=("fallback_v1", ["legacy_gate"]),
            ):
                result = invoice_service.analyze_invoice_v2_fast_text(
                    file_bytes=b"unused",
                    filename="factura.pdf",
                    mime_type="application/pdf",
                    prepared_text=prepared,
                    company_context={
                        "company_name": "Cliente Demo SL",
                        "company_tax_id": "B87654321",
                    },
                )

            canonical = result["invoice_parser_diagnostics"][
                "canonical_document_verifier_v2"
            ]
            self.assertEqual(result["invoice_date"], final_date)
            self.assertEqual(
                canonical["fields"]["invoice_date"]["status"], expected_status
            )

    def test_confirmed_canonical_v2_never_overrides_legacy_fallback(self):
        layout = _layout(
            ("FACTURA", "Nº", "F-100"),
            ("FECHA", "FACTURA", "23/08/2026"),
        )
        prepared = {
            "eligible": True,
            "reason": "canonical_text_sufficient",
            "input_representation": "canonical_layout_v1",
            "page_count": 1,
            "native_text_chars": 100,
            "sent_text_chars": 100,
            "document_text_complete": True,
            "text": "LEGACY VERIFIER TEXT",
            "verification_text": "LEGACY VERIFIER TEXT",
            "model_input_text": "CANONICAL MODEL TEXT",
            "_canonical_document_layout": layout,
        }
        response = {
            "supplier": {"legal_name": "Proveedor Demo SL", "tax_id": "B12345678"},
            "customer": {"legal_name": "Cliente Demo SL", "tax_id": "B87654321"},
            "invoice": {
                "invoice_number": "F-100",
                "issue_date": "2026-08-23",
                "currency": "EUR",
            },
            "due_dates": [],
            "taxes": [{"taxable_base": 100, "vat_rate": 21, "vat_amount": 21}],
            "totals": {
                "taxable_base": 100,
                "vat_amount": 21,
                "withholding": 0,
                "other_taxes": 0,
                "total": 121,
            },
            "field_evidence": {},
        }
        legacy_verification = {
            "invoice_number": {"status": "review", "reason": "legacy_gate"},
            "invoice_date": {"status": "review", "reason": "legacy_gate"},
        }
        with patch.object(invoice_service, "_get_client", return_value=object()), patch.object(
            invoice_service, "_get_invoice_model", return_value="gpt-5.6-sol"
        ), patch.object(
            invoice_service, "_call_invoice_responses", return_value=response
        ), patch.object(
            invoice_service,
            "_reconcile_fast_text_invoice_number",
            return_value=("F-100", [], [], "confirmed", {}),
        ), patch.object(
            invoice_service,
            "_reconcile_fast_text_invoice_date",
            return_value=("2026-08-23", [], [], {}),
        ), patch.object(
            invoice_service,
            "_reconcile_fast_text_payment_dates",
            return_value=([], []),
        ), patch.object(
            invoice_service, "_validate_fast_text_invoice", return_value=[]
        ), patch.object(
            invoice_service,
            "_assess_fast_text_invoice_safety",
            return_value={
                "accounting_safety_status": "passed",
                "accounting_safety_issues": [],
                "metadata_quality_status": "confirmed",
                "metadata_issues": [],
            },
        ), patch.object(
            invoice_service,
            "_verify_fast_text_document",
            return_value=legacy_verification,
        ), patch.object(
            invoice_service,
            "_decide_fast_text_fast_path",
            return_value=("fallback_v1", ["legacy_gate"]),
        ) as decide:
            result = invoice_service.analyze_invoice_v2_fast_text(
                file_bytes=b"unused",
                filename="factura.pdf",
                mime_type="application/pdf",
                prepared_text=prepared,
            )

        decide.assert_called_once_with(legacy_verification)
        self.assertEqual(result["fast_path_decision"], "fallback_v1")
        self.assertEqual(result["fast_path_reasons"], ["legacy_gate"])
        canonical = result["invoice_parser_diagnostics"][
            "canonical_document_verifier_v2"
        ]
        self.assertEqual(canonical["fields"]["invoice_number"]["status"], "confirmed")

    def test_safe_persistence_keeps_only_diagnostic_metadata(self):
        safe = ledger_app._safe_invoice_parser_diagnostics(
            {
                "canonical_document_verifier_v2": {
                    "version": "canonical_document_verifier_v2",
                    "fields": {
                        "invoice_number": {
                            "status": "confirmed",
                            "reason_code": "unique_association",
                            "value_occurrences": 1,
                            "anchor_clusters": 1,
                            "valid_associations": 1,
                            "competing_associations": 0,
                            "relation_type": "same_segment",
                            "evidence_class": "strong_same_row",
                            "text": "FACTURA F-100",
                            "bbox": [1, 2, 3, 4],
                            "candidate_value": "F-100",
                        }
                    },
                    "consistency": {},
                    "structure": {
                        "segments": 2,
                        "anchor_clusters": 1,
                        "regions": 1,
                        "candidates": 1,
                        "relations": 1,
                        "candidate_counts": {"identifier": 1},
                        "region_counts": {"invoice_metadata": 1},
                    },
                    "timings_ms": {
                        "structure_ms": 1.2,
                        "candidate_extraction_ms": 0.8,
                        "evidence_graph_ms": 0.4,
                        "verification_ms": 0.2,
                    },
                }
            }
        )

        persisted = str(safe["canonical_document_verifier_v2"])
        self.assertNotIn("FACTURA F-100", persisted)
        self.assertNotIn("bbox", persisted)
        self.assertNotIn("candidate_value", persisted)
        self.assertEqual(
            safe["canonical_document_verifier_v2"]["fields"]["invoice_number"][
                "value_occurrences"
            ],
            1,
        )


if __name__ == "__main__":
    unittest.main()
