import unittest

import app as ledger_app
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


def _proposal(**overrides):
    result = {
        "provider_name": "Vendedor Demo SL",
        "supplier_tax_id": "B12345678",
        "client_name": "Comprador Demo SL",
        "customer_tax_id": "B87654321",
        "invoice_number": "F-100",
        "invoice_date": "2026-08-07",
        "currency": "EUR",
        "base_amount": 100.0,
        "vat_amount": 21.0,
        "withholding_amount": 0.0,
        "other_taxes": 0.0,
        "total_amount": 121.0,
        "vat_breakdown": [
            {"rate": 21.0, "base": 100.0, "vat_amount": 21.0}
        ],
    }
    result.update(overrides)
    return result


def _verify(layout, proposal=None, **kwargs):
    return verify_canonical_document_v2(
        layout,
        proposal or _proposal(),
        registered_company_tax_id=kwargs.get("registered_company_tax_id", "B87654321"),
        registered_company_name=kwargs.get("registered_company_name", "Comprador Demo SL"),
    )


def _structure(layout, **kwargs):
    segments, token_by_id = build_semantic_segments(layout)
    anchors = build_anchor_clusters(segments, token_by_id)
    candidates = extract_field_candidates(segments, anchors, token_by_id)
    regions, segment_region = build_regions(
        segments,
        anchors,
        candidates,
        registered_company_tax_id=kwargs.get(
            "registered_company_tax_id", "B87654321"
        ),
        registered_company_name=kwargs.get(
            "registered_company_name", "Comprador Demo SL"
        ),
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


def _complete_totals(*extra_rows):
    return _layout(
        ("BASE", "IMPONIBLE", "100,00"),
        ("TOTAL", "IVA", "21,00"),
        *extra_rows,
        ("TOTAL", "FACTURA", "121,00"),
    )


class TestPartyIdentityResolver(unittest.TestCase):
    def test_a_supplier_and_recipient_are_separate_clusters(self):
        result = _verify(
            _layout(
                ("PROVEEDOR", "Vendedor", "Demo", "SL"),
                ("NIF", "B12345678"),
                ("CLIENTE", "Comprador", "Demo", "SL"),
                ("NIF", "B87654321"),
            )
        )

        self.assertEqual(result["party_resolution"]["party_cluster_count"], 2)
        self.assertEqual(result["fields"]["provider_name"]["status"], "confirmed")
        self.assertEqual(result["fields"]["supplier_tax_id"]["status"], "confirmed")
        self.assertEqual(result["fields"]["recipient_tax_id"]["status"], "confirmed")

    def test_b_known_recipient_can_appear_before_supplier(self):
        result = _verify(
            _layout(
                ("CLIENTE", "Comprador", "Demo", "SL"),
                ("NIF", "B87654321"),
                ("PROVEEDOR", "Vendedor", "Demo", "SL"),
                ("NIF", "B12345678"),
            )
        )

        self.assertEqual(result["party_resolution"]["recipient_status"], "confirmed")
        self.assertEqual(result["party_resolution"]["supplier_status"], "confirmed")

    def test_c_supplier_and_recipient_in_two_columns(self):
        result = _verify(
            _layout(
                (("CLIENTE", 30, 0), ("Comprador", 100, 0), ("PROVEEDOR", 320, 1), ("Vendedor", 410, 1)),
                (("Demo", 100, 0), ("SL", 135, 0), ("Demo", 410, 1), ("SL", 445, 1)),
                (("NIF", 30, 0), ("B87654321", 100, 0), ("NIF", 320, 1), ("B12345678", 410, 1)),
            )
        )

        self.assertEqual(result["fields"]["supplier_tax_id"]["status"], "confirmed")
        self.assertEqual(result["fields"]["recipient_tax_id"]["status"], "confirmed")

    def test_d_two_tax_ids_remain_bound_to_their_parties(self):
        result = _verify(
            _layout(
                ("EMISOR", "Vendedor", "Demo", "SL"),
                ("VAT", "B12345678"),
                ("RECEPTOR", "Comprador", "Demo", "SL"),
                ("VAT", "B87654321"),
            )
        )

        self.assertEqual(result["fields"]["supplier_tax_id"]["status"], "confirmed")
        self.assertEqual(result["fields"]["recipient_tax_id"]["status"], "confirmed")

    def test_e_name_and_tax_id_can_span_adjacent_lines(self):
        result = _verify(
            _layout(
                ("PROVEEDOR",),
                ("Vendedor", "Demo", "SL"),
                ("NIF", "B12345678"),
                ("CLIENTE",),
                ("Comprador", "Demo", "SL"),
                ("NIF", "B87654321"),
            )
        )

        self.assertEqual(result["fields"]["provider_name"]["status"], "confirmed")
        self.assertEqual(result["fields"]["supplier_tax_id"]["status"], "confirmed")

    def test_f_foreign_supplier_does_not_inherit_spanish_recipient_jurisdiction(self):
        result = _verify(
            _layout(
                ("SUPPLIER", "Foreign", "Vendor", "LLC"),
                ("TAX", "ID", "US123456789"),
                ("CLIENTE", "Comprador", "Demo", "SL"),
                ("NIF", "B87654321"),
                ("FECHA", "FACTURA", "07/08/26"),
            ),
            _proposal(
                provider_name="Foreign Vendor LLC",
                supplier_tax_id="US123456789",
            ),
        )

        self.assertIsNone(result["supplier_jurisdiction"])
        self.assertEqual(result["fields"]["invoice_date"]["status"], "review")

    def test_g_spanish_supplier_enables_spanish_invoice_date_resolution(self):
        result = _verify(
            _layout(
                ("PROVEEDOR", "Vendedor", "Demo", "SL"),
                ("NIF", "B12345678"),
                ("CLIENTE", "Comprador", "Demo", "SL"),
                ("NIF", "B87654321"),
                ("FECHA", "FACTURA", "07/08/26"),
            )
        )

        self.assertEqual(result["supplier_jurisdiction"], "ES")
        self.assertEqual(result["fields"]["invoice_date"]["status"], "confirmed")

    def test_h_multiple_plausible_supplier_clusters_remain_review(self):
        result = _verify(
            _layout(
                (("B12345678", 40, 0), ("B12345678", 360, 1)),
                ("CLIENTE", "Comprador", "Demo", "SL"),
                ("NIF", "B87654321"),
            )
        )

        self.assertEqual(result["party_resolution"]["supplier_status"], "review")
        self.assertEqual(result["fields"]["supplier_tax_id"]["status"], "review")

    def test_product_codes_without_tax_label_are_not_tax_ids(self):
        structure = _structure(
            _layout(
                ("PROVEEDOR", "Vendedor", "Demo", "SL"),
                ("NIF", "B12345678"),
                ("CLIENTE", "Comprador", "Demo", "SL"),
                ("NIF", "B87654321"),
                ("ARTICULO", "OR00038450", "1V401405", "ES8P20093"),
            )
        )

        tax_ids = {
            candidate.normalized_value
            for candidate in structure.candidates_by_type.get("tax_id", ())
        }
        self.assertEqual(tax_ids, {"B12345678", "B87654321"})

    def test_foreign_tax_id_requires_and_accepts_tax_label_context(self):
        structure = _structure(
            _layout(
                ("SUPPLIER", "Foreign", "Vendor", "LLC"),
                ("TAX", "ID", "US123456789"),
                ("REFERENCE", "US987654321"),
            )
        )

        tax_ids = {
            candidate.normalized_value
            for candidate in structure.candidates_by_type.get("tax_id", ())
        }
        self.assertIn("US123456789", tax_ids)
        self.assertNotIn("US987654321", tax_ids)

    def test_mixed_recipient_cluster_is_review_not_false_contradiction(self):
        result = _verify(
            _layout(
                ("CLIENTE", "Comprador", "Demo", "SL"),
                ("NIF", "B87654321", "B12345678"),
            )
        )

        self.assertEqual(result["fields"]["supplier_tax_id"]["status"], "review")


class TestFiscalStructureResolver(unittest.TestCase):
    def test_i_simple_21_percent_fiscal_table(self):
        result = _verify(
            _layout(
                ("TIPO", "BASE", "CUOTA"),
                ("21%", "100,00", "21,00"),
                ("BASE", "IMPONIBLE", "100,00"),
                ("TOTAL", "IVA", "21,00"),
                ("TOTAL", "FACTURA", "121,00"),
            )
        )

        self.assertEqual(result["fiscal_structure"]["fiscal_row_count"], 1)
        self.assertEqual(result["fields"]["vat_breakdown"]["status"], "confirmed")
        self.assertEqual(result["fields"]["base_amount"]["status"], "confirmed")
        self.assertEqual(result["fields"]["vat_amount"]["status"], "confirmed")

    def test_j_multi_rate_rows_use_exclusive_candidates(self):
        result = _verify(
            _layout(
                ("TIPO", "BASE", "CUOTA"),
                ("21%", "100,00", "21,00"),
                ("10%", "50,00", "5,00"),
                ("BASE", "IMPONIBLE", "150,00"),
                ("TOTAL", "IVA", "26,00"),
                ("TOTAL", "FACTURA", "176,00"),
            ),
            _proposal(
                base_amount=150,
                vat_amount=26,
                total_amount=176,
                vat_breakdown=[
                    {"rate": 21, "base": 100, "vat_amount": 21},
                    {"rate": 10, "base": 50, "vat_amount": 5},
                ],
            ),
        )

        self.assertEqual(result["fiscal_structure"]["fiscal_row_count"], 2)
        self.assertEqual(result["fields"]["vat_breakdown"]["status"], "confirmed")
        self.assertEqual(result["fields"]["vat_breakdown"]["valid_associations"], 2)

    def test_k_subtotal_is_not_invoice_total(self):
        result = _verify(
            _layout(
                ("BASE", "IMPONIBLE", "100,00"),
                ("TOTAL", "IVA", "21,00"),
                ("SUBTOTAL", "100,00"),
                ("TOTAL", "FACTURA", "121,00"),
            )
        )

        self.assertEqual(result["fields"]["total_amount"]["status"], "confirmed")
        self.assertEqual(result["fields"]["total_amount"]["competing_associations"], 0)

    def test_l_amount_due_is_distinct_from_invoice_total(self):
        result = _verify(
            _layout(
                ("BASE", "IMPONIBLE", "100,00"),
                ("TOTAL", "IVA", "21,00"),
                ("TOTAL", "FACTURA", "121,00"),
                ("TOTAL", "A", "PAGAR", "100,00"),
            )
        )

        self.assertEqual(result["fields"]["total_amount"]["status"], "confirmed")
        self.assertEqual(result["fields"]["total_amount"]["competing_associations"], 0)

    def test_m_explicit_withholding_is_confirmed(self):
        result = _verify(
            _layout(
                ("BASE", "IMPONIBLE", "100,00"),
                ("TOTAL", "IVA", "21,00"),
                ("RETENCION", "IRPF", "15,00"),
                ("TOTAL", "FACTURA", "106,00"),
            ),
            _proposal(withholding_amount=15, total_amount=106),
        )

        self.assertEqual(result["fields"]["withholding_amount"]["status"], "confirmed")

    def test_n_absent_withholding_is_not_applicable_with_full_coverage(self):
        result = _verify(_complete_totals())

        self.assertEqual(result["fields"]["withholding_amount"]["status"], "not_applicable")

    def test_o_explicit_other_tax_is_confirmed(self):
        result = _verify(
            _layout(
                ("BASE", "IMPONIBLE", "100,00"),
                ("TOTAL", "IVA", "21,00"),
                ("OTROS", "IMPUESTOS", "5,00"),
                ("TOTAL", "FACTURA", "126,00"),
            ),
            _proposal(other_taxes=5, total_amount=126),
        )

        self.assertEqual(result["fields"]["other_taxes"]["status"], "confirmed")

    def test_p_absent_other_tax_is_not_applicable_with_full_coverage(self):
        result = _verify(_complete_totals())

        self.assertEqual(result["fields"]["other_taxes"]["status"], "not_applicable")

    def test_q_borderless_fiscal_table_is_reconstructed(self):
        result = _verify(
            _layout(
                (("TIPO", 40), ("BASE", 180), ("CUOTA", 330)),
                (("21%", 40), ("100,00", 180), ("21,00", 330)),
            )
        )

        self.assertEqual(result["fiscal_structure"]["fiscal_table_status"], "confirmed")
        self.assertEqual(result["fields"]["vat_breakdown"]["status"], "confirmed")

    def test_r_repeated_amounts_in_table_and_totals_use_distinct_candidates(self):
        result = _verify(
            _layout(
                ("TIPO", "BASE", "CUOTA"),
                ("21%", "100,00", "21,00"),
                ("BASE", "IMPONIBLE", "100,00"),
                ("TOTAL", "IVA", "21,00"),
                ("TOTAL", "FACTURA", "121,00"),
            )
        )

        self.assertEqual(result["fields"]["base_amount"]["status"], "confirmed")
        self.assertEqual(result["fields"]["vat_amount"]["status"], "confirmed")
        self.assertEqual(result["fiscal_structure"]["ambiguous_fiscal_rows"], 0)

    def test_s_parallel_fiscal_columns_are_not_cross_assigned(self):
        result = _verify(
            _layout(
                (("TIPO", 30, 0), ("BASE", 90, 0), ("CUOTA", 170, 0), ("TIPO", 320, 1), ("BASE", 380, 1), ("CUOTA", 460, 1)),
                (("21%", 30, 0), ("100,00", 90, 0), ("21,00", 170, 0), ("10%", 320, 1), ("50,00", 380, 1), ("5,00", 460, 1)),
            ),
            _proposal(
                base_amount=150,
                vat_amount=26,
                total_amount=176,
                vat_breakdown=[
                    {"rate": 21, "base": 100, "vat_amount": 21},
                    {"rate": 10, "base": 50, "vat_amount": 5},
                ],
            ),
        )

        self.assertEqual(result["fiscal_structure"]["fiscal_table_status"], "review")
        self.assertEqual(result["fields"]["vat_breakdown"]["status"], "review")

    def test_numeric_rate_under_percent_vat_header_is_reconstructed(self):
        result = _verify(
            _layout(
                (("%", 30), ("IVA", 50), ("BASE", 180), ("IMPONIBLE", 220), ("CUOTA", 350), ("IVA", 395)),
                (("21,00", 30), ("100,00", 180), ("21,00", 350)),
            )
        )

        self.assertEqual(result["fiscal_structure"]["fiscal_table_status"], "confirmed")
        self.assertEqual(result["fields"]["vat_breakdown"]["status"], "confirmed")

    def test_split_fiscal_header_is_reconstructed(self):
        result = _verify(
            _layout(
                (("TIPO", 30), ("BASE", 180)),
                (("IVA", 30), ("IMPONIBLE", 180), ("CUOTA", 350)),
                (("21,00", 30), ("100,00", 180), ("21,00", 350)),
            )
        )

        self.assertEqual(result["fiscal_structure"]["fiscal_table_status"], "confirmed")
        self.assertEqual(result["fields"]["vat_breakdown"]["status"], "confirmed")

    def test_percentage_split_across_two_tokens_is_not_reused_as_money(self):
        result = _verify(
            _layout(
                (("TIPO", 30), ("BASE", 180), ("CUOTA", 350)),
                (("21,00", 30), ("%", 68), ("100,00", 180), ("21,00", 350)),
            )
        )

        self.assertEqual(result["fiscal_structure"]["fiscal_table_status"], "confirmed")
        self.assertEqual(result["fiscal_structure"]["fiscal_row_count"], 1)

    def test_adjacent_complete_amounts_do_not_form_composite_money_candidate(self):
        structure = _structure(_layout(("100,00", "21,00")))

        values = {
            candidate.normalized_value
            for candidate in structure.candidates_by_type.get("money", ())
        }
        self.assertEqual(values, {"100.00", "21.00"})

    def test_provisional_total_mismatch_stays_review(self):
        result = _verify(_layout(("TOTAL", "FACTURA", "999,00")))

        self.assertEqual(result["fields"]["total_amount"]["status"], "review")
        self.assertEqual(
            result["fields"]["total_amount"]["reason_code"],
            "provisional_total_amount_differs",
        )

    def test_confirmed_totals_mismatch_remains_contradiction(self):
        result = _verify(
            _layout(
                ("BASE", "IMPONIBLE", "100,00"),
                ("TOTAL", "IVA", "21,00"),
                ("TOTAL", "FACTURA", "999,00"),
            )
        )

        self.assertEqual(
            result["fields"]["total_amount"]["status"], "contradiction"
        )

    def test_safe_diagnostics_exclude_party_and_fiscal_values(self):
        safe = ledger_app._safe_invoice_parser_diagnostics(
            {
                "canonical_document_verifier_v2": {
                    "version": "canonical_document_verifier_v2",
                    "fields": {},
                    "consistency": {},
                    "structure": {},
                    "timings_ms": {},
                    "party_resolution": {
                        "supplier_status": "confirmed",
                        "recipient_status": "confirmed",
                        "supplier_tax_id_status": "confirmed",
                        "jurisdiction_status": "confirmed",
                        "party_cluster_count": 2,
                        "ambiguous_party_clusters": 0,
                        "supplier_name": "Vendedor Demo SL",
                        "supplier_tax_id": "B12345678",
                    },
                    "fiscal_structure": {
                        "fiscal_table_status": "confirmed",
                        "fiscal_row_count": 1,
                        "totals_block_status": "confirmed",
                        "totals_field_count": 3,
                        "ambiguous_fiscal_rows": 0,
                        "base_amount": "100.00",
                    },
                }
            }
        )["canonical_document_verifier_v2"]

        persisted = str(safe)
        self.assertNotIn("Vendedor Demo SL", persisted)
        self.assertNotIn("B12345678", persisted)
        self.assertNotIn("100.00", persisted)
        self.assertEqual(safe["party_resolution"]["party_cluster_count"], 2)
        self.assertEqual(safe["fiscal_structure"]["fiscal_row_count"], 1)


if __name__ == "__main__":
    unittest.main()
