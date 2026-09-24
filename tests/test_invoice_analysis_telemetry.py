import unittest
from datetime import datetime, timedelta
from unittest.mock import patch

from sqlalchemy import create_engine, select

import app as ledger_app
from services import ai_invoice_service


class FixedDateTime(datetime):
    current = datetime(2026, 9, 24, 10, 0, 0)

    @classmethod
    def utcnow(cls):
        return cls.current


class TestInvoiceAnalysisTelemetry(unittest.TestCase):
    def setUp(self):
        self.engine = create_engine("sqlite:///:memory:", future=True)
        self.engine_patch = patch.object(ledger_app, "engine", self.engine)
        self.engine_patch.start()
        ledger_app.metadata.create_all(self.engine)

    def tearDown(self):
        self.engine_patch.stop()
        self.engine.dispose()

    def _enqueue_job(self, queued_at="2026-09-24T10:00:00", status="queued"):
        with self.engine.begin() as conn:
            result = conn.execute(
                ledger_app.invoice_analysis_jobs_table.insert().values(
                    user_id=7,
                    company_id=11,
                    submitted_by_user_id=7,
                    document_type="expense",
                    original_filename="factura.pdf",
                    mime_type="application/pdf",
                    storage_key="private/invoice-analysis/test.pdf",
                    status=status,
                    result_json=None,
                    error_message=None,
                    attempt_count=0,
                    lease_expires_at=None,
                    created_at=queued_at,
                    started_at=None,
                    completed_at=None,
                    updated_at=queued_at,
                    expires_at="2026-10-01T10:00:00",
                )
            )
            job_id = result.inserted_primary_key[0]
            ledger_app._create_invoice_analysis_metrics(
                conn,
                job_id=job_id,
                user_id=7,
                company_id=11,
                document_type="expense",
                mime_type="application/pdf",
                file_size_bytes=1234,
                queued_at=queued_at,
            )
        return job_id

    def _metric(self, job_id):
        with self.engine.connect() as conn:
            return conn.execute(
                select(ledger_app.invoice_analysis_metrics_table).where(
                    ledger_app.invoice_analysis_metrics_table.c.job_id == job_id
                )
            ).mappings().one()

    def _job(self, job_id):
        with self.engine.connect() as conn:
            return conn.execute(
                select(ledger_app.invoice_analysis_jobs_table).where(
                    ledger_app.invoice_analysis_jobs_table.c.id == job_id
                )
            ).mappings().one()

    def test_queued_job_records_safe_metadata(self):
        job_id = self._enqueue_job()

        metric = self._metric(job_id)

        self.assertEqual(metric["status"], "queued")
        self.assertEqual(metric["queued_at"], "2026-09-24T10:00:00")
        self.assertEqual(metric["file_size_bytes"], 1234)
        self.assertEqual(metric["mime_type"], "application/pdf")
        self.assertNotIn("original_filename", ledger_app.invoice_analysis_metrics_table.c)
        self.assertNotIn("storage_key", ledger_app.invoice_analysis_metrics_table.c)

    def test_claim_sets_started_at_queue_wait_and_only_one_worker_claims(self):
        job_id = self._enqueue_job()
        FixedDateTime.current = datetime(2026, 9, 24, 10, 0, 2)

        with patch.object(ledger_app, "datetime", FixedDateTime), patch.dict(
            "os.environ", {"RENDER_INSTANCE_ID": "srv-instance-a"}, clear=False
        ):
            first_claim = ledger_app.claim_next_invoice_analysis_job()
            second_claim = ledger_app.claim_next_invoice_analysis_job()

        metric = self._metric(job_id)
        self.assertIsNotNone(first_claim)
        self.assertIsNone(second_claim)
        self.assertEqual(metric["status"], "processing")
        self.assertEqual(metric["started_at"], "2026-09-24T10:00:02")
        self.assertEqual(metric["queue_wait_ms"], 2000)
        self.assertEqual(metric["worker_instance_id"], "srv-instance-a")

    def test_claim_stays_processing_if_telemetry_write_fails(self):
        job_id = self._enqueue_job()
        FixedDateTime.current = datetime(2026, 9, 24, 10, 0, 2)

        with patch.object(ledger_app, "datetime", FixedDateTime), patch.object(
            ledger_app,
            "_mark_invoice_analysis_metrics_processing",
            side_effect=RuntimeError("metrics unavailable"),
        ):
            claim = ledger_app.claim_next_invoice_analysis_job()

        self.assertIsNotNone(claim)
        self.assertEqual(self._job(job_id)["status"], "processing")

    def test_completion_persists_durations_usage_ocr_and_audit(self):
        job_id = self._enqueue_job()
        with self.engine.begin() as conn:
            ledger_app._mark_invoice_analysis_metrics_processing(
                conn, job_id, "2026-09-24T10:00:02"
            )
            ledger_app._complete_invoice_analysis_metrics(
                conn,
                job_id=job_id,
                status="completed",
                completed_at="2026-09-24T10:00:12",
                telemetry={
                    "preprocessing_ms": 47,
                    "ocr_ms": 321,
                    "openai_ms": 7345,
                    "openai_model": "gpt-5.6-sol",
                    "input_tokens": 4210,
                    "output_tokens": 1200,
                    "reasoning_tokens": 320,
                    "total_tokens": 5410,
                    "ocr_used": True,
                    "audit_used": True,
                    "second_review_used": True,
                    "processing_type": "pdf_ocr_fallback",
                },
            )

        metric = self._metric(job_id)
        self.assertEqual(metric["status"], "completed")
        self.assertEqual(metric["completed_at"], "2026-09-24T10:00:12")
        self.assertEqual(metric["queue_wait_ms"], 2000)
        self.assertEqual(metric["processing_ms"], 10000)
        self.assertEqual(metric["openai_ms"], 7345)
        self.assertEqual(metric["openai_model"], "gpt-5.6-sol")
        self.assertEqual(metric["total_tokens"], 5410)
        self.assertTrue(metric["ocr_used"])
        self.assertTrue(metric["audit_used"])
        self.assertTrue(metric["second_review_used"])

    def test_worker_completion_keeps_source_deletion_and_persists_telemetry(self):
        job_id = self._enqueue_job(status="processing")
        job = dict(self._job(job_id))
        job["started_at"] = "2026-09-24T10:00:02"
        with self.engine.begin() as conn:
            conn.execute(
                ledger_app.invoice_analysis_jobs_table.update()
                .where(ledger_app.invoice_analysis_jobs_table.c.id == job_id)
                .values(started_at=job["started_at"])
            )
            ledger_app._mark_invoice_analysis_metrics_processing(
                conn, job_id, job["started_at"]
            )

        telemetry = {
            "preprocessing_ms": 20,
            "ocr_ms": None,
            "openai_ms": 400,
            "openai_model": "gpt-5.6-sol",
            "input_tokens": 100,
            "output_tokens": 50,
            "reasoning_tokens": 10,
            "total_tokens": 150,
            "ocr_used": False,
            "audit_used": False,
            "second_review_used": False,
            "processing_type": "pdf_embedded_text",
        }
        with patch.object(ledger_app, "cleanup_expired_invoice_analysis_jobs", return_value=0), patch.object(
            ledger_app, "claim_next_invoice_analysis_job", return_value=job
        ), patch.object(ledger_app, "download_private_bytes", return_value=b"%PDF-test"), patch.object(
            ledger_app, "get_company_names_for_analysis", return_value=[]
        ), patch.object(ledger_app, "fetch_known_suppliers", return_value=[]), patch.object(
            ledger_app, "_async_invoice_analysis_fallback_status", return_value=None
        ), patch.object(
            ledger_app,
            "_analyze_invoice_with_timeout",
            return_value=({"provider_name": "Proveedor"}, telemetry),
        ), patch.object(ledger_app, "delete_private_object") as delete_private:
            self.assertTrue(ledger_app.run_invoice_analysis_worker_once())

        self.assertEqual(self._job(job_id)["status"], "completed")
        self.assertEqual(self._metric(job_id)["total_tokens"], 150)
        delete_private.assert_called_once_with("private/invoice-analysis/test.pdf")

    def test_worker_failure_records_safe_error_type_and_keeps_source_deletion(self):
        job_id = self._enqueue_job(status="processing")
        job = dict(self._job(job_id))
        job["started_at"] = "2026-09-24T10:00:02"
        with self.engine.begin() as conn:
            conn.execute(
                ledger_app.invoice_analysis_jobs_table.update()
                .where(ledger_app.invoice_analysis_jobs_table.c.id == job_id)
                .values(started_at=job["started_at"])
            )
            ledger_app._mark_invoice_analysis_metrics_processing(
                conn, job_id, job["started_at"]
            )

        with patch.object(ledger_app, "cleanup_expired_invoice_analysis_jobs", return_value=0), patch.object(
            ledger_app, "claim_next_invoice_analysis_job", return_value=job
        ), patch.object(
            ledger_app, "download_private_bytes", side_effect=RuntimeError("storage unavailable")
        ), patch.object(ledger_app, "delete_private_object") as delete_private:
            self.assertTrue(ledger_app.run_invoice_analysis_worker_once())

        self.assertEqual(self._job(job_id)["status"], "failed")
        metric = self._metric(job_id)
        self.assertEqual(metric["status"], "failed")
        self.assertEqual(metric["error_type"], "RuntimeError")
        delete_private.assert_called_once_with("private/invoice-analysis/test.pdf")


class TestInvoiceResponseTelemetry(unittest.TestCase):
    def test_response_usage_is_aggregated_without_another_api_call(self):
        telemetry = {"openai_ms": 100, "input_tokens": 3}
        response = {
            "model": "gpt-5.6-sol",
            "usage": {
                "input_tokens": 7,
                "output_tokens": 5,
                "total_tokens": 12,
                "output_tokens_details": {"reasoning_tokens": 2},
            },
        }

        ai_invoice_service._record_invoice_response_telemetry(
            telemetry,
            response=response,
            model="fallback-model",
            elapsed_ms=250,
            audit=True,
        )

        self.assertEqual(telemetry["openai_ms"], 350)
        self.assertEqual(telemetry["openai_model"], "gpt-5.6-sol")
        self.assertEqual(telemetry["input_tokens"], 10)
        self.assertEqual(telemetry["output_tokens"], 5)
        self.assertEqual(telemetry["reasoning_tokens"], 2)
        self.assertEqual(telemetry["total_tokens"], 12)
        self.assertTrue(telemetry["audit_used"])
        self.assertTrue(telemetry["second_review_used"])

    def test_missing_usage_remains_unknown(self):
        telemetry = {"openai_ms": 0}

        ai_invoice_service._record_invoice_response_telemetry(
            telemetry,
            response={"model": "gpt-5.6-sol", "usage": {}},
            model="fallback-model",
            elapsed_ms=25,
            audit=False,
        )

        self.assertEqual(telemetry["openai_ms"], 25)
        self.assertNotIn("input_tokens", telemetry)
        self.assertNotIn("output_tokens", telemetry)
        self.assertNotIn("reasoning_tokens", telemetry)
        self.assertNotIn("total_tokens", telemetry)


if __name__ == "__main__":
    unittest.main()
