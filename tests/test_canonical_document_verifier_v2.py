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

    def test_a_canonical_receives_invoice_number_before_legacy_reconciliation(self):
        result, captured = self._run_independence_pipeline(
            legacy_invoice_number="LEGACY-B"
        )

        self.assertEqual(captured["before"]["invoice_number"], "F-100")
        self.assertEqual(result["invoice_number"], "LEGACY-B")

    def test_b_canonical_receives_invoice_date_before_legacy_reconciliation(self):
        result, captured = self._run_independence_pipeline(
            legacy_invoice_date="2026-09-30"
        )

        self.assertEqual(captured["before"]["invoice_date"], "2026-08-23")
        self.assertEqual(result["invoice_date"], "2026-09-30")

    def test_c_nested_legacy_mutations_do_not_change_canonical_snapshot(self):
        def mutate_legacy_nested(_structured, normalized, *_args, **_kwargs):
            normalized["vat_breakdown"][0]["base"] = 999.0
            normalized["payment_dates"].append("2026-10-01")
            return []

        result, captured = self._run_independence_pipeline(
            validate_side_effect=mutate_legacy_nested
        )

        self.assertEqual(captured["before"]["vat_breakdown"][0]["base"], 100.0)
        self.assertEqual(captured["before"]["payment_dates"], ["2026-09-01"])
        self.assertEqual(result["vat_breakdown"][0]["base"], 999.0)
        self.assertEqual(
            result["payment_dates"], ["2026-09-01", "2026-10-01"]
        )

    def test_d_canonical_branch_cannot_mutate_legacy_result_or_decision(self):
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

    def test_e_legacy_and_canonical_branches_diverge_without_shared_mutation(self):
        def mutate_legacy(_structured, normalized, *_args, **_kwargs):
            normalized["provider_name"] = "Legacy Supplier SL"
            return []

        def mutate_canonical(proposal):
            proposal["supplier_tax_id"] = "CANONICAL123"

        result, captured = self._run_independence_pipeline(
            validate_side_effect=mutate_legacy,
            canonical_mutator=mutate_canonical,
        )

        self.assertEqual(captured["before"]["provider_name"], "Model Supplier SL")
        self.assertEqual(captured["before"]["supplier_tax_id"], "B12345678")
        self.assertEqual(captured["after"]["supplier_tax_id"], "CANONICAL123")
        self.assertEqual(result["provider_name"], "Legacy Supplier SL")
        self.assertEqual(result["supplier_tax_id"], "B12345678")

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
