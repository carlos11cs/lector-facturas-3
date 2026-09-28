import json
import unittest
from datetime import datetime, timedelta
from unittest.mock import patch

try:
    import fitz
except ModuleNotFoundError:  # pragma: no cover - dependency is required in deployment
    fitz = None
from sqlalchemy import create_engine, select

import app as ledger_app
from services import ai_invoice_service as invoice_service


V2_CRISTINA_RECIPIENT_FIXTURE = """Código
Descripción del Artículo
CRISTINA SIMÓ BESALDUCH
CORREOS, 14
VALENCIA (VALENCIA)
N.I.F. 20903358S
Factura Número: Q000081/2026
KALOS HEALTH AND BEAUTY SL
C/CERVANTES Nº25 BAJO
VALENCIA
N.I.F: B05410667
Fecha: 13/05/26
BASE IMPONIBLE: 677,83 €
TOTAL CUOTAS: 27,11 €
TOTAL FACTURA: 704,94 €
"""


def _digital_pdf_bytes(text_by_page):
    document = fitz.open()
    for text in text_by_page:
        page = document.new_page()
        page.insert_text((72, 72), text)
    data = document.tobytes()
    document.close()
    return data


def _fast_text_invoice_payload(
    *,
    supplier_name,
    supplier_tax_id,
    customer_name,
    customer_tax_id,
):
    return {
        "supplier": {"legal_name": supplier_name, "tax_id": supplier_tax_id},
        "customer": {"legal_name": customer_name, "tax_id": customer_tax_id},
        "invoice": {"invoice_number": "Q000081/2026", "issue_date": "2026-05-13", "currency": "EUR"},
        "due_dates": [],
        "taxes": [{"taxable_base": 677.83, "vat_rate": 4, "vat_amount": 27.11}],
        "totals": {
            "taxable_base": 677.83,
            "vat_amount": 27.11,
            "withholding": 0,
            "other_taxes": 0,
            "total": 704.94,
        },
        "field_evidence": {
            key: {"page": 1, "evidence": "Dato visible", "confidence": 0.9}
            for key in ("supplier", "invoice_number", "issue_date", "taxes", "totals", "withholding", "due_dates")
        },
    }


@unittest.skipIf(fitz is None, "PyMuPDF no disponible")
class TestInvoiceV2FastText(unittest.TestCase):
    def _analyze_fast_text(self, structured, text):
        prepared = {
            "eligible": True,
            "reason": "native_text_sufficient",
            "page_count": 1,
            "native_text_chars": len(text),
            "sent_text_chars": len(text),
            "text": "[PÁGINA 1]\n" + text,
        }
        with patch.object(invoice_service, "_get_client", return_value=object()), patch.object(
            invoice_service, "_get_invoice_model", return_value="gpt-5.6-sol"
        ), patch.object(invoice_service, "_call_invoice_responses", return_value=structured):
            return invoice_service.analyze_invoice_v2_fast_text(
                file_bytes=b"unused",
                filename="factura.pdf",
                mime_type="application/pdf",
                prepared_text=prepared,
            )

    def test_digital_pdf_is_eligible_and_keeps_page_boundaries(self):
        payload = _digital_pdf_bytes(
            [
                "FACTURA\nProveedor Demo SL\nFactura F-1\nFecha 01/09/2026",
                "Base imponible 100,00 EUR\nIVA 21,00 EUR\nTotal a pagar 121,00 EUR\n"
                "Dirección fiscal del emisor y condiciones de vencimiento 01/10/2026",
            ]
        )

        prepared = invoice_service.prepare_invoice_v2_fast_text(
            payload, filename="factura.pdf", mime_type="application/pdf"
        )

        self.assertTrue(prepared["eligible"])
        self.assertEqual(prepared["reason"], "native_text_sufficient")
        self.assertEqual(prepared["page_count"], 2)
        self.assertIn("[PÁGINA 1]", prepared["text"])
        self.assertIn("[PÁGINA 2]", prepared["text"])
        self.assertGreater(prepared["native_text_chars"], 100)
        self.assertGreater(prepared["sent_text_chars"], 0)

    def test_scanned_pdf_is_not_eligible(self):
        payload = _digital_pdf_bytes([""])

        prepared = invoice_service.prepare_invoice_v2_fast_text(
            payload, filename="escaneado.pdf", mime_type="application/pdf"
        )

        self.assertFalse(prepared["eligible"])
        self.assertEqual(prepared["reason"], "scanned_pdf")
        self.assertEqual(prepared["sent_text_chars"], 0)

    def test_v2_request_uses_text_only_strict_schema_and_store_false(self):
        class FakeResponse:
            model = "gpt-5.6-sol"
            status = "completed"
            output = []
            output_text = json.dumps(
                {
                    "supplier": {"legal_name": "Proveedor Demo SL", "tax_id": "B12345678"},
                    "customer": {"legal_name": "Cliente Demo SL", "tax_id": "B87654321"},
                    "invoice": {"invoice_number": "F-1", "issue_date": "2026-09-01", "currency": "EUR"},
                    "due_dates": ["2026-10-01"],
                    "taxes": [{"taxable_base": 100, "vat_rate": 21, "vat_amount": 21}],
                    "totals": {"taxable_base": 100, "vat_amount": 21, "withholding": 0, "other_taxes": 0, "total": 121},
                    "field_evidence": {
                        key: {"page": 1, "evidence": "Dato visible", "confidence": 0.9}
                        for key in ("supplier", "invoice_number", "issue_date", "taxes", "totals", "withholding", "due_dates")
                    },
                }
            )

        class FakeResponses:
            def __init__(self):
                self.kwargs = None

            def create(self, **kwargs):
                self.kwargs = kwargs
                return FakeResponse()

        class FakeClient:
            def __init__(self):
                self.responses = FakeResponses()

            def with_options(self, **_kwargs):
                return self

        prepared = {
            "eligible": True,
            "reason": "native_text_sufficient",
            "page_count": 1,
            "native_text_chars": 500,
            "sent_text_chars": 300,
            "text": "[PÁGINA 1]\nFACTURA F-1",
        }
        client = FakeClient()
        with patch.object(invoice_service, "_get_client", return_value=client), patch.dict(
            "os.environ", {"OPENAI_INVOICE_MODEL": "gpt-5.6-sol"}, clear=True
        ):
            result, _telemetry = invoice_service.analyze_invoice_v2_fast_text(
                file_bytes=b"unused",
                filename="factura.pdf",
                mime_type="application/pdf",
                prepared_text=prepared,
                return_telemetry=True,
            )

        request = client.responses.kwargs
        self.assertFalse(any(item.get("type") == "input_file" for item in request["input"][0]["content"]))
        self.assertEqual(request["store"], False)
        self.assertTrue(request["text"]["format"]["strict"])
        self.assertEqual(request["text"]["format"]["name"], "invoice_fast_text_extraction")
        self.assertEqual(result["validation_status"], "passed")

    def test_v2_validation_rejects_customer_as_supplier_and_math_mismatch(self):
        structured = {
            "supplier": {"legal_name": "KALOS HEALTH AND BEAUTY SL", "tax_id": "B05410667"},
            "customer": {"legal_name": "KALOS HEALTH AND BEAUTY SL", "tax_id": "B05410667"},
            "invoice": {"invoice_number": "F-1", "issue_date": "2026-09-01", "currency": "EUR"},
            "due_dates": [],
            "taxes": [{"taxable_base": 100, "vat_rate": 21, "vat_amount": 21}],
            "totals": {"taxable_base": 100, "vat_amount": 21, "withholding": 0, "other_taxes": 0, "total": 120},
            "field_evidence": {
                key: {"page": 1, "evidence": "Dato", "confidence": 0.9}
                for key in ("supplier", "invoice_number", "issue_date", "taxes", "totals", "withholding", "due_dates")
            },
        }
        normalized = invoice_service._normalize_fast_text_invoice(structured)
        issues = invoice_service._validate_fast_text_invoice(
            structured,
            normalized,
            ["KALOS HEALTH AND BEAUTY SL"],
            registered_company_tax_id="B05410667",
        )

        self.assertIn("supplier_matches_registered_customer", issues)
        self.assertIn("supplier_tax_id_matches_registered_customer", issues)
        self.assertIn("supplier_matches_customer", issues)
        self.assertIn("inconsistent_accounting_equation", issues)

    def test_v2_corrects_invoice_number_labeled_as_invoice_over_order_and_delivery_note(self):
        structured = _fast_text_invoice_payload(
            supplier_name="Henry Schein Medical SL",
            supplier_tax_id="B12345678",
            customer_name="Cliente Demo SL",
            customer_tax_id="B87654321",
        )
        structured["invoice"]["invoice_number"] = "17944113"
        structured["due_dates"] = []
        result = self._analyze_fast_text(
            structured,
            "Nº FACTURA A141949\nPEDIDO 17944113\nALBARÁN 681188\n"
            "RECIBO 15 DIAS FECHA FACTURA",
        )

        self.assertEqual(result["invoice_number"], "A141949")
        self.assertEqual(result["payment_dates"], ["2026-05-28"])
        self.assertEqual(
            result["deterministic_corrections"],
            [
                "invoice_number_corrected_from_explicit_label",
                "due_date_derived_from_payment_terms",
            ],
        )
        self.assertEqual(result["validation_status"], "passed")

    def test_v2_does_not_correct_ambiguous_invoice_number_labels(self):
        structured = _fast_text_invoice_payload(
            supplier_name="Proveedor Demo SL",
            supplier_tax_id="B12345678",
            customer_name="Cliente Demo SL",
            customer_tax_id="B87654321",
        )
        structured["invoice"]["invoice_number"] = "17944113"
        result = self._analyze_fast_text(
            structured,
            "Nº FACTURA A141949\nFACTURA Nº A141950\nPEDIDO 17944113",
        )

        self.assertEqual(result["invoice_number"], "17944113")
        self.assertEqual(result["deterministic_corrections"], [])
        self.assertIn("invoice_number_evidence_ambiguous", result["validation_issues"])
        self.assertEqual(result["validation_status"], "failed")

    def test_v2_rejects_an_order_number_when_no_invoice_label_exists(self):
        structured = _fast_text_invoice_payload(
            supplier_name="Proveedor Demo SL",
            supplier_tax_id="B12345678",
            customer_name="Cliente Demo SL",
            customer_tax_id="B87654321",
        )
        structured["invoice"]["invoice_number"] = "17944113"
        result = self._analyze_fast_text(structured, "PEDIDO 17944113\nALBARÁN 681188")

        self.assertIsNone(result["invoice_number"])
        self.assertIn("invoice_number_is_non_invoice_reference", result["validation_issues"])
        self.assertEqual(result["validation_status"], "failed")

    def test_v2_keeps_explicit_due_date_over_relative_terms(self):
        structured = _fast_text_invoice_payload(
            supplier_name="Proveedor Demo SL",
            supplier_tax_id="B12345678",
            customer_name="Cliente Demo SL",
            customer_tax_id="B87654321",
        )
        result = self._analyze_fast_text(
            structured,
            "Nº FACTURA A141949\nRECIBO 15 DIAS FECHA FACTURA\n"
            "FECHA DE VENCIMIENTO 01/06/2026",
        )

        self.assertEqual(result["payment_dates"], ["2026-06-01"])
        self.assertNotIn("due_date_derived_from_payment_terms", result["deterministic_corrections"])

    def test_v2_does_not_derive_due_date_from_ambiguous_payment_terms(self):
        structured = _fast_text_invoice_payload(
            supplier_name="Proveedor Demo SL",
            supplier_tax_id="B12345678",
            customer_name="Cliente Demo SL",
            customer_tax_id="B87654321",
        )
        structured["due_dates"] = []
        result = self._analyze_fast_text(
            structured,
            "Nº FACTURA A141949\nPAGO 15/30 DIAS FECHA FACTURA",
        )

        self.assertEqual(result["payment_dates"], [])
        self.assertNotIn("due_date_derived_from_payment_terms", result["deterministic_corrections"])

    def test_v2_accepts_unambiguous_relative_terms_without_recibo_prefix(self):
        self.assertEqual(
            invoice_service._extract_fast_text_relative_payment_terms_days(
                "30 DÍAS FECHA FACTURA"
            ),
            30,
        )
        self.assertEqual(
            invoice_service._extract_fast_text_relative_payment_terms_days(
                "60 DIAS FECHA FACTURA"
            ),
            60,
        )

    def test_v2_job_63_regression_matches_v1_after_label_and_terms_reconciliation(self):
        structured = _fast_text_invoice_payload(
            supplier_name="Henry Schein Medical SL",
            supplier_tax_id="B12345678",
            customer_name="Cliente Demo SL",
            customer_tax_id="B87654321",
        )
        structured["invoice"] = {
            "invoice_number": "17944113",
            "issue_date": "2026-07-16",
            "currency": "EUR",
        }
        structured["due_dates"] = []
        structured["taxes"] = [{"taxable_base": 44.23, "vat_rate": 21, "vat_amount": 7.26}]
        structured["totals"] = {
            "taxable_base": 44.23,
            "vat_amount": 7.26,
            "withholding": 0,
            "other_taxes": 0,
            "total": 51.49,
        }
        result = self._analyze_fast_text(
            structured,
            "Nº FACTURA A141949\nPEDIDO 17944113\nALBARÁN 681188\n"
            "RECIBO 15 DIAS FECHA FACTURA",
        )
        v1_result = {
            "provider_name": "Henry Schein Medical SL",
            "invoice_date": "2026-07-16",
            "payment_dates": ["2026-07-31"],
            "base_amount": 44.23,
            "vat_amount": 7.26,
            "withholding_amount": 0.0,
            "total_amount": 51.49,
            "vat_breakdown": [{"base": 44.23, "rate": 21, "vat_amount": 7.26}],
            "structured_extraction": {
                "supplier": {"tax_id": "B12345678"},
                "invoice": {"invoice_number": "A141949"},
            },
        }

        comparison = ledger_app._compare_invoice_v1_and_v2(v1_result, result)

        self.assertEqual(result["invoice_number"], "A141949")
        self.assertEqual(result["payment_dates"], ["2026-07-31"])
        self.assertTrue(comparison["strict_accounting_match"])

    def test_v2_context_keeps_known_recipient_separate_from_person_supplier(self):
        structured = _fast_text_invoice_payload(
            supplier_name="Cristina Simo Besalduch",
            supplier_tax_id="20903358S",
            customer_name="KALOS HEALTH AND BEAUTY SL",
            customer_tax_id="B05410667",
        )
        prepared = {
            "eligible": True,
            "reason": "native_text_sufficient",
            "page_count": 1,
            "native_text_chars": len(V2_CRISTINA_RECIPIENT_FIXTURE),
            "sent_text_chars": len(V2_CRISTINA_RECIPIENT_FIXTURE),
            "text": "[PÁGINA 1]\n" + V2_CRISTINA_RECIPIENT_FIXTURE,
        }
        with patch.object(invoice_service, "_get_client", return_value=object()), patch.object(
            invoice_service, "_get_invoice_model", return_value="gpt-5.6-sol"
        ), patch.object(
            invoice_service, "_call_invoice_responses", return_value=structured
        ) as call_responses:
            result = invoice_service.analyze_invoice_v2_fast_text(
                file_bytes=b"unused",
                filename="cristina.pdf",
                mime_type="application/pdf",
                company_names=["KALOS HEALTH AND BEAUTY SL"],
                company_context={
                    "company_name": "KALOS HEALTH AND BEAUTY SL",
                    "company_tax_id": "B05410667",
                },
                prepared_text=prepared,
            )

        prompt = call_responses.call_args.kwargs["prompt"]
        self.assertIn("KALOS HEALTH AND BEAUTY SL", prompt)
        self.assertIn("B05410667", prompt)
        self.assertIn("cliente/receptor/comprador", prompt)
        self.assertIn("proveedor/emisor", prompt)
        self.assertEqual(result["provider_name"], "Cristina Simo Besalduch")
        self.assertEqual(result["supplier_tax_id"], "20903358S")
        self.assertNotEqual(result["supplier_tax_id"], "B05410667")
        self.assertEqual(result["validation_status"], "passed")

    def test_v2_context_handles_recipient_before_issuer_in_native_text(self):
        structured = _fast_text_invoice_payload(
            supplier_name="Cristina Simo Besalduch",
            supplier_tax_id="20903358S",
            customer_name="KALOS HEALTH AND BEAUTY SL",
            customer_tax_id="B05410667",
        )
        recipient_first_text = (
            "[PÁGINA 1]\nCLIENTE: KALOS HEALTH AND BEAUTY SL\nNIF: B05410667\n"
            "EMISORA: Cristina Simo Besalduch\nNIF: 20903358S\n"
            "Factura Q000081/2026\nBase 677,83\nIVA 27,11\nTotal 704,94"
        )
        prepared = {
            "eligible": True,
            "reason": "native_text_sufficient",
            "page_count": 1,
            "native_text_chars": len(recipient_first_text),
            "sent_text_chars": len(recipient_first_text),
            "text": recipient_first_text,
        }
        with patch.object(invoice_service, "_get_client", return_value=object()), patch.object(
            invoice_service, "_get_invoice_model", return_value="gpt-5.6-sol"
        ), patch.object(
            invoice_service, "_call_invoice_responses", return_value=structured
        ) as call_responses:
            result = invoice_service.analyze_invoice_v2_fast_text(
                file_bytes=b"unused",
                filename="recipient-first.pdf",
                company_context={
                    "company_name": "KALOS HEALTH AND BEAUTY SL",
                    "company_tax_id": "B05410667",
                },
                prepared_text=prepared,
            )

        sent_document = call_responses.call_args.kwargs["response_input"][0]["content"][1]["text"]
        self.assertLess(sent_document.index("KALOS HEALTH"), sent_document.index("Cristina Simo"))
        self.assertEqual(result["provider_name"], "Cristina Simo Besalduch")
        self.assertEqual(result["supplier_tax_id"], "20903358S")
        self.assertEqual(result["validation_status"], "passed")

    def test_v2_validation_accepts_person_supplier_and_company_customer_with_both_tax_ids(self):
        structured = _fast_text_invoice_payload(
            supplier_name="Cristina Simo Besalduch",
            supplier_tax_id="20903358S",
            customer_name="KALOS HEALTH AND BEAUTY SL",
            customer_tax_id="B05410667",
        )
        issues = invoice_service._validate_fast_text_invoice(
            structured,
            invoice_service._normalize_fast_text_invoice(structured),
            ["KALOS HEALTH AND BEAUTY SL"],
            registered_company_tax_id="B05410667",
        )

        self.assertEqual(issues, [])

    def test_v2_validation_accepts_distinct_company_supplier_and_customer(self):
        structured = _fast_text_invoice_payload(
            supplier_name="Mediderma SLU",
            supplier_tax_id="B12345678",
            customer_name="KALOS HEALTH AND BEAUTY SL",
            customer_tax_id="B05410667",
        )
        issues = invoice_service._validate_fast_text_invoice(
            structured,
            invoice_service._normalize_fast_text_invoice(structured),
            ["KALOS HEALTH AND BEAUTY SL"],
            registered_company_tax_id="B05410667",
        )

        self.assertEqual(issues, [])

    def test_v2_context_does_not_require_company_tax_id_to_appear_in_document(self):
        structured = _fast_text_invoice_payload(
            supplier_name="Proveedor Ajeno SL",
            supplier_tax_id="B12345678",
            customer_name="Cliente Invitado SL",
            customer_tax_id="B87654321",
        )
        issues = invoice_service._validate_fast_text_invoice(
            structured,
            invoice_service._normalize_fast_text_invoice(structured),
            ["KALOS HEALTH AND BEAUTY SL"],
            registered_company_tax_id="B05410667",
        )

        self.assertEqual(issues, [])

    def test_v2_validation_rejects_ambiguous_assignment_of_registered_company_as_supplier(self):
        structured = _fast_text_invoice_payload(
            supplier_name="KALOS HEALTH AND BEAUTY SL",
            supplier_tax_id="B05410667",
            customer_name="KALOS HEALTH AND BEAUTY SL",
            customer_tax_id="B05410667",
        )
        issues = invoice_service._validate_fast_text_invoice(
            structured,
            invoice_service._normalize_fast_text_invoice(structured),
            ["KALOS HEALTH AND BEAUTY SL"],
            registered_company_tax_id="B05410667",
        )

        self.assertIn("supplier_matches_registered_customer", issues)
        self.assertIn("supplier_tax_id_matches_registered_customer", issues)
        self.assertIn("supplier_matches_customer", issues)


class TestInvoiceV2ShadowQueue(unittest.TestCase):
    def setUp(self):
        self.engine = create_engine("sqlite:///:memory:", future=True)
        self.engine_patch = patch.object(ledger_app, "engine", self.engine)
        self.engine_patch.start()
        ledger_app.metadata.create_all(self.engine)

    def tearDown(self):
        self.engine_patch.stop()
        self.engine.dispose()

    def _v1_result(self):
        return {
            "analysis_status": "ok",
            "provider_name": "Proveedor Demo, S.L.",
            "invoice_number": "F-001",
            "invoice_date": "2026-09-01",
            "payment_dates": ["2026-10-01"],
            "base_amount": 100.0,
            "vat_amount": 21.0,
            "withholding_amount": 0.0,
            "total_amount": 121.0,
            "vat_breakdown": [{"rate": 21.0, "base": 100.0, "vat_amount": 21.0}],
        }

    def _v1_production_result(self, invoice_number):
        """Match the persisted V1 shape used by production shadow runs."""
        result = self._v1_result()
        result.pop("invoice_number")
        result["structured_extraction"] = {
            "supplier": {"legal_name": "Proveedor Demo, S.L.", "tax_id": "B12345678"},
            "invoice": {"invoice_number": invoice_number, "issue_date": "2026-09-01"},
            "totals": {
                "taxable_base": 100.0,
                "vat_amount": 21.0,
                "withholding": 0.0,
                "total": 121.0,
            },
            "taxes": [{"taxable_base": 100.0, "vat_rate": 21.0, "vat_amount": 21.0}],
            "installments": [{"due_date": "2026-10-01"}],
        }
        return result

    def _v2_result(self, invoice_number="F001"):
        return {
            "analysis_status": "ok",
            "validation_status": "passed",
            "provider_name": "Proveedor Demo SL",
            "supplier_tax_id": "B12345678",
            "invoice_number": invoice_number,
            "invoice_date": "2026-09-01",
            "payment_dates": ["2026-10-01"],
            "currency": "EUR",
            "base_amount": 100.0,
            "vat_amount": 21.0,
            "withholding_amount": 0.0,
            "other_taxes": 0.0,
            "total_amount": 121.0,
            "vat_breakdown": [{"rate": 21.0, "base": 100.0, "vat_amount": 21.0}],
        }

    def _prepared(self, eligible=True):
        return {
            "eligible": eligible,
            "reason": "native_text_sufficient" if eligible else "scanned_pdf",
            "page_count": 1,
            "native_text_chars": 500,
            "sent_text_chars": 300 if eligible else 0,
            "size_reduction_ratio": 0.4,
            "text": "[PÁGINA 1]\nFACTURA",
        }

    def _create_completed_job(self, *, batch_id="batch-shadow", storage_key="private/invoice-analysis/shadow.pdf"):
        now = datetime.utcnow().isoformat()
        with self.engine.begin() as conn:
            result = conn.execute(
                ledger_app.invoice_analysis_jobs_table.insert().values(
                    user_id=7,
                    company_id=11,
                    submitted_by_user_id=7,
                    document_type="expense",
                    original_filename="factura.pdf",
                    mime_type="application/pdf",
                    storage_key=storage_key,
                    status="completed",
                    result_json=json.dumps(self._v1_result()),
                    error_message=None,
                    attempt_count=1,
                    next_attempt_at=None,
                    deferred_retry_count=0,
                    lease_expires_at=None,
                    lease_token=None,
                    lease_renewal_count=0,
                    batch_id=batch_id,
                    batch_position=1,
                    created_at=now,
                    started_at=now,
                    completed_at=now,
                    updated_at=now,
                    expires_at=(datetime.utcnow() + timedelta(days=1)).isoformat(),
                )
            )
        return result.inserted_primary_key[0]

    def _create_processing_job(self):
        job_id = self._create_completed_job()
        now = datetime.utcnow().isoformat()
        with self.engine.begin() as conn:
            conn.execute(
                ledger_app.invoice_analysis_jobs_table.update()
                .where(ledger_app.invoice_analysis_jobs_table.c.id == job_id)
                .values(
                    status="processing",
                    result_json=None,
                    completed_at=None,
                    started_at=now,
                    lease_token="v1-lease-token",
                    lease_expires_at=(datetime.utcnow() + timedelta(minutes=10)).isoformat(),
                )
            )
        return job_id

    def _job(self, job_id):
        with self.engine.connect() as conn:
            return conn.execute(
                select(ledger_app.invoice_analysis_jobs_table).where(
                    ledger_app.invoice_analysis_jobs_table.c.id == job_id
                )
            ).mappings().one()

    def _run(self, job_id):
        with self.engine.connect() as conn:
            return conn.execute(
                select(ledger_app.invoice_analysis_shadow_runs_table).where(
                    ledger_app.invoice_analysis_shadow_runs_table.c.job_id == job_id
                )
            ).mappings().one()

    def _enqueue_shadow(self, job_id, prepared=None):
        job = dict(self._job(job_id))
        self.assertTrue(ledger_app._create_invoice_v2_shadow_run(job, prepared or self._prepared()))

    def test_shadow_disabled_and_deterministic_sampling(self):
        with patch.object(ledger_app, "INVOICE_V2_SHADOW_ENABLED", False):
            self.assertFalse(ledger_app._invoice_v2_shadow_is_selected(7))
        with patch.object(ledger_app, "INVOICE_V2_SHADOW_ENABLED", True), patch.object(
            ledger_app, "INVOICE_V2_SHADOW_SAMPLE_RATE", 1.0
        ):
            self.assertTrue(ledger_app._invoice_v2_shadow_is_selected(7))
        with patch.object(ledger_app, "INVOICE_V2_SHADOW_ENABLED", True), patch.object(
            ledger_app, "INVOICE_V2_SHADOW_SAMPLE_RATE", 0.5
        ):
            self.assertEqual(
                ledger_app._invoice_v2_shadow_is_selected(17),
                ledger_app._invoice_v2_shadow_is_selected(17),
            )

    def test_shadow_runs_allow_one_record_per_version_for_the_same_job(self):
        job_id = self._create_completed_job()
        self._enqueue_shadow(job_id)
        self._enqueue_shadow(job_id)

        with patch.object(ledger_app, "INVOICE_V2_SHADOW_VERSION", "v2-next-text-v1"):
            self._enqueue_shadow(job_id)

        with self.engine.connect() as conn:
            runs = conn.execute(
                select(
                    ledger_app.invoice_analysis_shadow_runs_table.c.shadow_version
                )
                .where(ledger_app.invoice_analysis_shadow_runs_table.c.job_id == job_id)
                .order_by(ledger_app.invoice_analysis_shadow_runs_table.c.shadow_version)
            ).scalars().all()

        self.assertEqual(runs, ["v2-next-text-v1", "v2-sol-text-v3"])

    def test_ineligible_shadow_is_recorded_without_retaining_source(self):
        job_id = self._create_completed_job()
        self._enqueue_shadow(job_id, self._prepared(eligible=False))

        run = self._run(job_id)
        self.assertEqual(run["status"], "skipped")
        self.assertFalse(run["eligible"])
        self.assertEqual(run["validation_status"], "not_applicable")
        self.assertEqual(run["batch_id"], "batch-shadow")

    def test_v1_publishes_before_shadow_and_retains_source_only_for_v2(self):
        job_id = self._create_processing_job()
        job = dict(self._job(job_id))
        with patch.object(ledger_app, "download_private_bytes", return_value=b"%PDF"), patch.object(
            ledger_app, "get_company_names_for_analysis", return_value=[]
        ), patch.object(ledger_app, "fetch_known_suppliers", return_value=[]), patch.object(
            ledger_app, "_async_invoice_analysis_fallback_status", return_value=None
        ), patch.object(
            ledger_app, "_analyze_invoice_with_timeout", return_value=(self._v1_result(), {"openai_ms": 1})
        ), patch.object(
            ledger_app, "_invoice_v2_shadow_candidate", return_value=self._prepared()
        ), patch.object(ledger_app, "delete_private_object") as delete_private:
            self.assertTrue(ledger_app._run_claimed_invoice_analysis_job(job))

        stored = self._job(job_id)
        self.assertEqual(stored["status"], "completed")
        self.assertEqual(json.loads(stored["result_json"]), self._v1_result())
        self.assertEqual(stored["storage_key"], "private/invoice-analysis/shadow.pdf")
        self.assertEqual(self._run(job_id)["status"], "queued")
        delete_private.assert_not_called()

    def test_v2_failure_does_not_modify_official_v1_result_and_releases_source(self):
        job_id = self._create_completed_job()
        self._enqueue_shadow(job_id)
        with patch.object(ledger_app, "download_private_bytes", return_value=b"%PDF"), patch.object(
            ledger_app, "prepare_invoice_v2_fast_text", return_value=self._prepared()
        ), patch.object(ledger_app, "get_company_names_for_analysis", return_value=[]), patch.object(
            ledger_app,
            "analyze_invoice_v2_fast_text",
            return_value=(
                {
                    "analysis_status": "failed",
                    "validation_status": "not_run",
                    "analysis_error": {"status": "rate_limited"},
                },
                {"openai_ms": 1},
            ),
        ), patch.object(ledger_app, "delete_private_object") as delete_private:
            self.assertTrue(ledger_app.run_invoice_v2_shadow_worker_once())

        self.assertEqual(json.loads(self._job(job_id)["result_json"]), self._v1_result())
        self.assertIsNone(self._job(job_id)["storage_key"])
        self.assertEqual(self._run(job_id)["status"], "failed")
        self.assertEqual(self._run(job_id)["error_type"], "rate_limited")
        delete_private.assert_called_once_with("private/invoice-analysis/shadow.pdf")

    def test_shadow_persists_validation_codes_without_document_evidence(self):
        job_id = self._create_completed_job()
        self._enqueue_shadow(job_id)
        rejected = self._v2_result()
        rejected.update(
            {
                "analysis_status": "failed",
                "validation_status": "failed",
                "validation_issues": [
                    "supplier_matches_registered_customer",
                    "supplier_tax_id_matches_registered_customer",
                ],
            }
        )
        with patch.object(ledger_app, "download_private_bytes", return_value=b"%PDF"), patch.object(
            ledger_app, "prepare_invoice_v2_fast_text", return_value=self._prepared()
        ), patch.object(ledger_app, "get_company_names_for_analysis", return_value=[]), patch.object(
            ledger_app, "get_company_context_for_invoice_v2", return_value={}
        ), patch.object(
            ledger_app, "analyze_invoice_v2_fast_text", return_value=(rejected, {"openai_ms": 1})
        ), patch.object(ledger_app, "delete_private_object"):
            self.assertTrue(ledger_app.run_invoice_v2_shadow_worker_once())

        run = self._run(job_id)
        self.assertEqual(run["status"], "completed")
        self.assertEqual(run["validation_status"], "failed")
        self.assertEqual(
            json.loads(run["validation_errors_json"]),
            [
                "supplier_matches_registered_customer",
                "supplier_tax_id_matches_registered_customer",
            ],
        )
        self.assertNotIn("evidence", run["validation_errors_json"])

    def test_shadow_worker_passes_registered_company_identity_only_to_v2(self):
        job_id = self._create_completed_job()
        self._enqueue_shadow(job_id)
        with self.engine.begin() as conn:
            conn.execute(
                ledger_app.companies_table.insert().values(
                    id=11,
                    user_id=7,
                    display_name="Kalos",
                    legal_name="KALOS HEALTH AND BEAUTY SL",
                    tax_id="B05410667",
                    company_type="company",
                    created_at=datetime.utcnow().isoformat(),
                )
            )
        with patch.object(ledger_app, "download_private_bytes", return_value=b"%PDF"), patch.object(
            ledger_app, "prepare_invoice_v2_fast_text", return_value=self._prepared()
        ), patch.object(
            ledger_app, "analyze_invoice_v2_fast_text", return_value=(self._v2_result(), {"openai_ms": 1})
        ) as analyze_v2, patch.object(ledger_app, "delete_private_object"):
            self.assertTrue(ledger_app.run_invoice_v2_shadow_worker_once())

        self.assertEqual(
            analyze_v2.call_args.kwargs["company_context"],
            {
                "company_name": "KALOS HEALTH AND BEAUTY SL",
                "company_tax_id": "B05410667",
            },
        )

    def test_shadow_comparison_uses_accounting_tolerance_and_multiple_vat_lines(self):
        v1 = self._v1_result()
        v1["base_amount"] = 200.0
        v1["vat_amount"] = 31.0
        v1["total_amount"] = 231.0
        v1["vat_breakdown"] = [
            {"rate": 21, "base": 100, "vat_amount": 21},
            {"rate": 10, "base": 100, "vat_amount": 10},
        ]
        v2 = self._v2_result()
        v2.update(
            {
                "base_amount": 200.009,
                "vat_amount": 31.001,
                "total_amount": 231.0,
                "vat_breakdown": list(reversed(v1["vat_breakdown"])),
            }
        )

        comparison = ledger_app._compare_invoice_v1_and_v2(v1, v2)

        self.assertTrue(comparison["provider_match"])
        self.assertTrue(comparison["invoice_number_match"])
        self.assertTrue(comparison["tax_base_match"])
        self.assertTrue(comparison["vat_match"])
        self.assertTrue(comparison["overall_match"])

    def test_shadow_comparison_reads_invoice_number_and_supplier_tax_id_from_real_v1_shape(self):
        v1 = self._v1_production_result("Q000144/2026")
        v2 = self._v2_result("Q000144/2026")

        canonical_v1 = ledger_app._canonical_v1_shadow_result(v1)
        comparison = ledger_app._compare_invoice_v1_and_v2(v1, v2)

        self.assertEqual(canonical_v1["invoice_number"], "Q000144/2026")
        self.assertEqual(canonical_v1["supplier_tax_id"], "B12345678")
        self.assertTrue(comparison["invoice_number_match"])
        self.assertTrue(comparison["overall_match"])
        self.assertTrue(comparison["strict_accounting_match"])
        self.assertTrue(json.loads(comparison["comparison_json"])["supplier_tax_id_match"])

    def test_shadow_supplier_tax_id_matches_equivalent_formatting(self):
        v1 = self._v1_production_result("Q000144/2026")
        v1["structured_extraction"]["supplier"]["tax_id"] = "B-123.45678"
        v2 = self._v2_result("Q000144/2026")
        v2["supplier_tax_id"] = "b 12345678"

        comparison = ledger_app._compare_invoice_v1_and_v2(v1, v2)

        self.assertTrue(json.loads(comparison["comparison_json"])["supplier_tax_id_match"])
        self.assertTrue(comparison["strict_accounting_match"])

    def test_shadow_supplier_tax_id_difference_fails_only_strict_match(self):
        v1 = self._v1_production_result("Q000144/2026")
        v2 = self._v2_result("Q000144/2026")
        v2["supplier_tax_id"] = "B87654321"

        comparison = ledger_app._compare_invoice_v1_and_v2(v1, v2)

        self.assertFalse(json.loads(comparison["comparison_json"])["supplier_tax_id_match"])
        self.assertTrue(comparison["overall_match"])
        self.assertFalse(comparison["strict_accounting_match"])

    def test_shadow_comparison_regression_for_jobs_47_48_and_49(self):
        for invoice_number in ("Q000144/2026", "Q000147/2026", "Q000148/2026"):
            with self.subTest(invoice_number=invoice_number):
                comparison = ledger_app._compare_invoice_v1_and_v2(
                    self._v1_production_result(invoice_number),
                    self._v2_result(invoice_number),
                )

                self.assertTrue(comparison["invoice_number_match"])
                self.assertTrue(comparison["provider_match"])
                self.assertTrue(comparison["invoice_date_match"])
                self.assertTrue(comparison["tax_base_match"])
                self.assertTrue(comparison["vat_match"])
                self.assertTrue(comparison["withholding_match"])
                self.assertTrue(comparison["total_match"])
                self.assertTrue(comparison["due_date_match"])
                self.assertTrue(comparison["overall_match"])
                self.assertTrue(comparison["strict_accounting_match"])

    def test_shadow_comparison_recalculation_is_dry_run_by_default_and_never_calls_openai(self):
        job_id = self._create_completed_job()
        v1 = self._v1_production_result("Q000144/2026")
        v2 = self._v2_result("Q000144/2026")
        with self.engine.begin() as conn:
            conn.execute(
                ledger_app.invoice_analysis_jobs_table.update()
                .where(ledger_app.invoice_analysis_jobs_table.c.id == job_id)
                .values(result_json=json.dumps(v1))
            )
        self._enqueue_shadow(job_id)
        with self.engine.begin() as conn:
            conn.execute(
                ledger_app.invoice_analysis_shadow_runs_table.update()
                .where(ledger_app.invoice_analysis_shadow_runs_table.c.job_id == job_id)
                .values(
                    status="completed",
                    validation_status="passed",
                    result_json=json.dumps(v2),
                    provider_match=True,
                    invoice_number_match=False,
                    invoice_date_match=True,
                    tax_base_match=True,
                    vat_match=True,
                    withholding_match=True,
                    total_match=True,
                    due_date_match=True,
                    overall_match=False,
                    strict_accounting_match=False,
                    comparison_json='{"different_fields":["invoice_number_match"]}',
                )
            )

        dry_run = ledger_app.recalculate_invoice_v2_shadow_comparisons(job_ids=[job_id])
        self.assertEqual(dry_run, {
            "scanned": 1,
            "recalculated": 1,
            "skipped": 0,
            "updated": 0,
            "dry_run": True,
        })
        self.assertFalse(self._run(job_id)["overall_match"])

        applied = ledger_app.recalculate_invoice_v2_shadow_comparisons(
            job_ids=[job_id], apply=True
        )
        self.assertEqual(applied["updated"], 1)
        self.assertTrue(self._run(job_id)["invoice_number_match"])
        self.assertTrue(self._run(job_id)["overall_match"])
        self.assertTrue(self._run(job_id)["strict_accounting_match"])
        self.assertTrue(
            json.loads(self._run(job_id)["comparison_json"])["supplier_tax_id_match"]
        )

    def test_shadow_claim_is_fenced_and_only_one_worker_receives_it(self):
        job_id = self._create_completed_job()
        self._enqueue_shadow(job_id)
        first = ledger_app.claim_next_invoice_v2_shadow_run()
        second = ledger_app.claim_next_invoice_v2_shadow_run()

        self.assertEqual(first["job_id"], job_id)
        self.assertIsNone(second)
        stale = dict(first)
        with self.engine.begin() as conn:
            conn.execute(
                ledger_app.invoice_analysis_shadow_runs_table.update()
                .where(ledger_app.invoice_analysis_shadow_runs_table.c.id == first["id"])
                .values(lease_expires_at="2020-01-01T00:00:00")
            )
        recovered = ledger_app.claim_next_invoice_v2_shadow_run()
        self.assertFalse(
            ledger_app._finish_invoice_v2_shadow_run(
                stale, status="failed", validation_status="not_run", error_type="stale"
            )
        )
        self.assertNotEqual(stale["lease_token"], recovered["lease_token"])

    def test_shadow_result_persists_no_source_text_prompt_or_pdf(self):
        job_id = self._create_completed_job()
        self._enqueue_shadow(job_id)
        v2_result = self._v2_result()
        v2_result["deterministic_corrections"] = [
            "invoice_number_corrected_from_explicit_label",
            "due_date_derived_from_payment_terms",
        ]
        with patch.object(ledger_app, "download_private_bytes", return_value=b"%PDF"), patch.object(
            ledger_app, "prepare_invoice_v2_fast_text", return_value=self._prepared()
        ), patch.object(ledger_app, "get_company_names_for_analysis", return_value=[]), patch.object(
            ledger_app, "analyze_invoice_v2_fast_text", return_value=(v2_result, {"openai_ms": 1})
        ), patch.object(ledger_app, "delete_private_object"):
            self.assertTrue(ledger_app.run_invoice_v2_shadow_worker_once())

        run = self._run(job_id)
        payload = json.loads(run["result_json"])
        self.assertEqual(run["status"], "completed")
        self.assertTrue(run["overall_match"])
        self.assertNotIn("text", payload)
        self.assertNotIn("evidence", payload)
        self.assertNotIn("prompt", payload)
        self.assertEqual(
            payload["deterministic_corrections"],
            [
                "invoice_number_corrected_from_explicit_label",
                "due_date_derived_from_payment_terms",
            ],
        )
        self.assertNotIn("storage_key", ledger_app.invoice_analysis_shadow_runs_table.c)
        self.assertNotIn("original_filename", ledger_app.invoice_analysis_shadow_runs_table.c)
        self.assertIn("validation_errors_json", ledger_app.invoice_analysis_shadow_runs_table.c)
        self.assertIn("strict_accounting_match", ledger_app.invoice_analysis_shadow_runs_table.c)


if __name__ == "__main__":
    unittest.main()
