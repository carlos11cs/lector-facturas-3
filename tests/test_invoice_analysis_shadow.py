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


def _digital_pdf_bytes(text_by_page):
    document = fitz.open()
    for text in text_by_page:
        page = document.new_page()
        page.insert_text((72, 72), text)
    data = document.tobytes()
    document.close()
    return data


@unittest.skipIf(fitz is None, "PyMuPDF no disponible")
class TestInvoiceV2FastText(unittest.TestCase):
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
            structured, normalized, ["KALOS HEALTH AND BEAUTY SL"]
        )

        self.assertIn("supplier_matches_registered_customer", issues)
        self.assertIn("supplier_matches_customer", issues)
        self.assertIn("inconsistent_accounting_equation", issues)


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

    def _v2_result(self):
        return {
            "analysis_status": "ok",
            "validation_status": "passed",
            "provider_name": "Proveedor Demo SL",
            "supplier_tax_id": "B12345678",
            "invoice_number": "F001",
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

        self.assertEqual(runs, ["v2-next-text-v1", "v2-sol-text-v1"])

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
        with patch.object(ledger_app, "download_private_bytes", return_value=b"%PDF"), patch.object(
            ledger_app, "prepare_invoice_v2_fast_text", return_value=self._prepared()
        ), patch.object(ledger_app, "get_company_names_for_analysis", return_value=[]), patch.object(
            ledger_app, "analyze_invoice_v2_fast_text", return_value=(self._v2_result(), {"openai_ms": 1})
        ), patch.object(ledger_app, "delete_private_object"):
            self.assertTrue(ledger_app.run_invoice_v2_shadow_worker_once())

        run = self._run(job_id)
        payload = json.loads(run["result_json"])
        self.assertEqual(run["status"], "completed")
        self.assertTrue(run["overall_match"])
        self.assertNotIn("text", payload)
        self.assertNotIn("evidence", payload)
        self.assertNotIn("prompt", payload)
        self.assertNotIn("storage_key", ledger_app.invoice_analysis_shadow_runs_table.c)
        self.assertNotIn("original_filename", ledger_app.invoice_analysis_shadow_runs_table.c)


if __name__ == "__main__":
    unittest.main()
