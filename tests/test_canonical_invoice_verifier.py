import unittest
from unittest.mock import patch

import app as ledger_app
from services import ai_invoice_service as invoice_service
from services.canonical_invoice_verifier import (
    verify_invoice_date,
    verify_invoice_number,
)
from services.document_layout import DocumentLayout, build_layout_page


def _layout(*rows):
    words = []
    for line_index, row in enumerate(rows):
        x = 72.0
        y = 72.0 + (line_index * 20)
        for word_index, value in enumerate(row):
            width = max(12.0, len(value) * 6.0)
            words.append(
                (x, y, x + width, y + 10, value, 0, line_index, word_index)
            )
            x += width + 8
    page = build_layout_page(
        words,
        page=1,
        width=595,
        height=842,
        line_directions={
            (0, line_index): ((1.0, 0.0), 0)
            for line_index in range(len(rows))
        },
    )
    return DocumentLayout((page,))


class TestCanonicalInvoiceNumberVerifier(unittest.TestCase):
    def test_label_and_number_on_same_row(self):
        result = verify_invoice_number(_layout(("Nº", "FACTURA", "F-100")), "F-100")

        self.assertEqual(result["status"], "confirmed")
        self.assertEqual(result["relation_type"], "same_row")

    def test_label_above_number(self):
        result = verify_invoice_number(
            _layout(("Nº", "FACTURA"), ("F-100",)), "F-100"
        )

        self.assertEqual(result["status"], "confirmed")
        self.assertEqual(result["relation_type"], "vertical")

    def test_order_number_is_not_invoice_evidence(self):
        result = verify_invoice_number(_layout(("PEDIDO", "26049906")), "26049906")

        self.assertEqual(result["status"], "review")
        self.assertEqual(result["relation_type"], "secondary")

    def test_delivery_number_is_not_invoice_evidence(self):
        result = verify_invoice_number(
            _layout(("DELIVERY", "NUMBER", "26049906")), "26049906"
        )

        self.assertEqual(result["status"], "review")
        self.assertEqual(result["relation_type"], "secondary")

    def test_two_explicit_candidates_remain_review(self):
        result = verify_invoice_number(
            _layout(
                ("Nº", "FACTURA", "F-100"),
                ("Nº", "FACTURA", "F-101"),
            ),
            "F-100",
        )

        self.assertEqual(result["status"], "review")
        self.assertEqual(result["reason"], "multiple_explicit_invoice_numbers")

    def test_repeated_number_remains_review(self):
        result = verify_invoice_number(
            _layout(
                ("Nº", "FACTURA", "F-100"),
                ("Nº", "FACTURA", "F-100"),
            ),
            "F-100",
        )

        self.assertEqual(result["status"], "review")
        self.assertEqual(result["reason"], "multiple_invoice_number_associations")
        self.assertEqual(result["match_count"], 2)

    def test_incompatible_explicit_invoice_number_is_contradiction(self):
        result = verify_invoice_number(
            _layout(
                ("Nº", "FACTURA", "F-200"),
                ("REFERENCIA", "F-100"),
            ),
            "F-100",
        )

        self.assertEqual(result["status"], "contradiction")
        self.assertEqual(
            result["reason"], "invoice_number_conflicts_with_explicit_label"
        )


class TestCanonicalInvoiceDateVerifier(unittest.TestCase):
    def test_label_and_date_on_same_row(self):
        result = verify_invoice_date(
            _layout(("FECHA", "FACTURA", "23/08/2026")), "2026-08-23"
        )

        self.assertEqual(result["status"], "confirmed")
        self.assertEqual(result["relation_type"], "same_row")

    def test_label_above_date(self):
        result = verify_invoice_date(
            _layout(("FECHA", "FACTURA"), ("23/08/2026",)), "2026-08-23"
        )

        self.assertEqual(result["status"], "confirmed")
        self.assertEqual(result["relation_type"], "vertical")

    def test_due_date_is_excluded(self):
        result = verify_invoice_date(
            _layout(("VENCIMIENTO", "23/08/2026")), "2026-08-23"
        )

        self.assertEqual(result["status"], "review")
        self.assertEqual(result["relation_type"], "excluded")

    def test_order_date_is_excluded(self):
        result = verify_invoice_date(
            _layout(("ORDER", "DATE", "23/08/2026")), "2026-08-23"
        )

        self.assertEqual(result["status"], "review")
        self.assertEqual(result["relation_type"], "excluded")

    def test_two_invoice_dates_remain_review(self):
        result = verify_invoice_date(
            _layout(
                ("FECHA", "FACTURA", "23/08/2026"),
                ("DOCUMENT", "DATE", "24/08/2026"),
            ),
            "2026-08-23",
        )

        self.assertEqual(result["status"], "review")
        self.assertEqual(result["reason"], "multiple_invoice_date_associations")

    def test_spanish_supplier_resolves_ambiguous_day_first_date(self):
        result = verify_invoice_date(
            _layout(
                ("PROVEEDOR", "NIF", "B12345678"),
                ("FECHA", "FACTURA", "07/08/26"),
            ),
            "2026-08-07",
            supplier_tax_id="B12345678",
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(result["status"], "confirmed")

    def test_foreign_supplier_keeps_ambiguous_date_in_review(self):
        result = verify_invoice_date(
            _layout(
                ("SUPPLIER", "TAX", "US123456789"),
                ("INVOICE", "DATE", "07/08/26"),
            ),
            "2026-08-07",
            supplier_tax_id="US123456789",
            registered_company_tax_id="B87654321",
        )

        self.assertEqual(result["status"], "review")
        self.assertEqual(result["reason"], "invoice_date_format_ambiguous")

    def test_incompatible_explicit_invoice_date_is_contradiction(self):
        result = verify_invoice_date(
            _layout(("FECHA", "FACTURA", "24/08/2026")), "2026-08-23"
        )

        self.assertEqual(result["status"], "contradiction")
        self.assertEqual(
            result["reason"], "invoice_date_conflicts_with_explicit_label"
        )


class TestCanonicalDiagnosticIntegration(unittest.TestCase):
    def test_diagnostic_does_not_change_legacy_fast_path_decision(self):
        document_layout = _layout(
            ("Nº", "FACTURA", "F-100"),
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
            "_canonical_document_layout": document_layout,
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
            "invoice_number": {"status": "review", "reason": "legacy_number_review"},
            "invoice_date": {"status": "review", "reason": "legacy_date_review"},
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
        diagnostics = result["invoice_parser_diagnostics"]
        self.assertEqual(
            diagnostics["canonical_verification"]["invoice_number"]["status"],
            "confirmed",
        )
        self.assertEqual(
            diagnostics["canonical_verification_comparison"]["invoice_number"][
                "transition"
            ],
            "legacy_review_to_canonical_confirmed",
        )

    def test_persistence_whitelist_excludes_document_content_and_geometry(self):
        safe = ledger_app._safe_invoice_parser_diagnostics(
            {
                "canonical_verification": {
                    "invoice_number": {
                        "status": "confirmed",
                        "reason": "unique_spatial_invoice_label",
                        "match_count": 1,
                        "relation_type": "same_row",
                        "text": "FACTURA F-100",
                        "coordinates": [1, 2, 3, 4],
                    },
                    "invoice_date": {
                        "status": "review",
                        "reason": "invoice_date_format_ambiguous",
                        "match_count": 1,
                        "relation_type": "same_row",
                        "tokens": ["07/08/26"],
                    },
                },
                "canonical_verification_comparison": {
                    "invoice_number": {
                        "legacy_status": "review",
                        "legacy_reason": "legacy_review",
                        "canonical_status": "confirmed",
                        "canonical_reason": "unique_spatial_invoice_label",
                        "transition": "legacy_review_to_canonical_confirmed",
                        "alternative_value": "F-999",
                    }
                },
            }
        )

        serialized = str(safe)
        self.assertNotIn("FACTURA F-100", serialized)
        self.assertNotIn("coordinates", serialized)
        self.assertNotIn("tokens", serialized)
        self.assertNotIn("alternative_value", serialized)
        self.assertEqual(
            safe["canonical_verification"]["invoice_number"]["status"],
            "confirmed",
        )


if __name__ == "__main__":
    unittest.main()
