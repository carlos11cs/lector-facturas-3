import json
import unittest
from datetime import datetime, timedelta
from unittest.mock import patch

from sqlalchemy import create_engine, select

import app as ledger_app


class TestInvoiceAnalysisDismissal(unittest.TestCase):
    def setUp(self):
        self.engine = create_engine("sqlite:///:memory:", future=True)
        self.engine_patch = patch.object(ledger_app, "engine", self.engine)
        self.engine_patch.start()
        ledger_app.metadata.create_all(self.engine)

    def tearDown(self):
        self.engine_patch.stop()
        self.engine.dispose()

    def _enqueue(self, *, status="completed", company_id=11, batch_id="dismiss-batch"):
        now = datetime.utcnow().isoformat()
        with self.engine.begin() as conn:
            result = conn.execute(
                ledger_app.invoice_analysis_jobs_table.insert().values(
                    user_id=7,
                    company_id=company_id,
                    submitted_by_user_id=7,
                    document_type="expense",
                    original_filename="factura.pdf",
                    mime_type="application/pdf",
                    storage_key="private/invoice-analysis/dismiss-test.pdf",
                    status=status,
                    result_json=json.dumps({"provider_name": "Proveedor demo"})
                    if status == "completed"
                    else None,
                    error_message="No se pudo analizar" if status == "failed" else None,
                    attempt_count=1,
                    next_attempt_at=None,
                    deferred_retry_count=0,
                    lease_expires_at=None,
                    lease_token=None,
                    lease_renewal_count=0,
                    batch_id=batch_id,
                    batch_position=1,
                    created_at=now,
                    started_at=now if status != "queued" else None,
                    completed_at=now if status in {"completed", "failed"} else None,
                    dismissed_at=None,
                    updated_at=now,
                    expires_at=(datetime.utcnow() + timedelta(days=1)).isoformat(),
                )
            )
            job_id = result.inserted_primary_key[0]
            ledger_app._create_invoice_analysis_metrics(
                conn,
                job_id=job_id,
                user_id=7,
                company_id=company_id,
                document_type="expense",
                mime_type="application/pdf",
                file_size_bytes=100,
                queued_at=now,
                batch_id=batch_id,
                batch_position=1,
            )
        return job_id

    def _job(self, job_id):
        with self.engine.connect() as conn:
            return conn.execute(
                select(ledger_app.invoice_analysis_jobs_table).where(
                    ledger_app.invoice_analysis_jobs_table.c.id == job_id
                )
            ).mappings().one()

    def _dismiss(self, job_id, *, company_id=11):
        with patch.object(ledger_app, "get_data_owner_id", return_value=7), patch.object(
            ledger_app, "get_company_id", return_value=company_id
        ), ledger_app.app.test_request_context(
            f"/api/invoice-analysis-jobs/{job_id}", method="DELETE"
        ):
            return ledger_app.app.make_response(
                ledger_app.dismiss_invoice_analysis_job(job_id)
            )

    def _list(self, *, company_id=11):
        with patch.object(ledger_app, "get_data_owner_id", return_value=7), patch.object(
            ledger_app, "get_company_id", return_value=company_id
        ), ledger_app.app.test_request_context("/api/invoice-analysis-jobs"):
            return ledger_app.list_invoice_analysis_jobs()

    def _add_shadow_run(self, job_id):
        now = datetime.utcnow().isoformat()
        with self.engine.begin() as conn:
            conn.execute(
                ledger_app.invoice_analysis_shadow_runs_table.insert().values(
                    job_id=job_id,
                    user_id=7,
                    company_id=11,
                    batch_id="dismiss-batch",
                    batch_position=1,
                    shadow_version="v2-sol-text-v1",
                    route="v2_fast_text_native",
                    model="gpt-5.6-sol",
                    reasoning_effort="low",
                    eligible=True,
                    eligibility_reason="native_text_sufficient",
                    status="completed",
                    validation_status="passed",
                    attempt_count=1,
                    created_at=now,
                    started_at=now,
                    completed_at=now,
                    updated_at=now,
                )
            )

    def test_dismiss_completed_hides_reload_and_preserves_technical_history(self):
        job_id = self._enqueue(status="completed")
        self._add_shadow_run(job_id)

        response = self._dismiss(job_id)

        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.get_json()["dismissed"])
        stored = self._job(job_id)
        self.assertEqual(stored["status"], "completed")
        self.assertEqual(json.loads(stored["result_json"])["provider_name"], "Proveedor demo")
        self.assertIsNotNone(stored["dismissed_at"])
        self.assertEqual(self._list().get_json()["jobs"], [])
        with self.engine.connect() as conn:
            metric = conn.execute(
                select(ledger_app.invoice_analysis_metrics_table.c.id).where(
                    ledger_app.invoice_analysis_metrics_table.c.job_id == job_id
                )
            ).scalar_one_or_none()
            shadow = conn.execute(
                select(ledger_app.invoice_analysis_shadow_runs_table.c.id).where(
                    ledger_app.invoice_analysis_shadow_runs_table.c.job_id == job_id
                )
            ).scalar_one_or_none()
        self.assertIsNotNone(metric)
        self.assertIsNotNone(shadow)

    def test_dismiss_failed_is_idempotent_and_keeps_failed_status(self):
        job_id = self._enqueue(status="failed")

        first = self._dismiss(job_id)
        second = self._dismiss(job_id)

        self.assertTrue(first.get_json()["dismissed"])
        self.assertTrue(second.get_json()["dismissed"])
        self.assertTrue(second.get_json()["idempotent"])
        self.assertEqual(self._job(job_id)["status"], "failed")
        self.assertEqual(self._list().get_json()["jobs"], [])

    def test_reload_returns_only_jobs_not_dismissed_and_batch_hides_dismissed(self):
        dismissed_job_id = self._enqueue(status="completed")
        visible_job_id = self._enqueue(status="failed")
        self._dismiss(dismissed_job_id)

        listed_ids = [job["id"] for job in self._list().get_json()["jobs"]]
        self.assertEqual(listed_ids, [visible_job_id])
        with patch.object(ledger_app, "get_data_owner_id", return_value=7), patch.object(
            ledger_app, "get_company_id", return_value=11
        ), ledger_app.app.test_request_context(
            "/api/invoice-analysis-batches/dismiss-batch?page=1&page_size=50"
        ):
            response = ledger_app.get_invoice_analysis_batch("dismiss-batch")
        self.assertEqual([job["id"] for job in response.get_json()["jobs"]], [visible_job_id])

    def test_dismiss_is_scoped_to_company(self):
        job_id = self._enqueue(status="completed", company_id=22)

        response = self._dismiss(job_id, company_id=11)

        self.assertEqual(response.status_code, 404)
        self.assertIsNone(self._job(job_id)["dismissed_at"])

    def test_dismiss_during_processing_keeps_worker_and_source_cleanup_intact(self):
        job_id = self._enqueue(status="queued")
        claimed = ledger_app.claim_next_invoice_analysis_job()
        self.assertEqual(claimed["id"], job_id)
        self.assertTrue(self._dismiss(job_id).get_json()["dismissed"])

        extracted = {"analysis_status": "ok", "provider_name": "Proveedor demo"}
        with patch.object(ledger_app, "download_private_bytes", return_value=b"%PDF-test"), patch.object(
            ledger_app, "get_company_names_for_analysis", return_value=[]
        ), patch.object(ledger_app, "fetch_known_suppliers", return_value=[]), patch.object(
            ledger_app, "_analyze_invoice_with_timeout", return_value=(extracted, {"openai_ms": 1})
        ), patch.object(ledger_app, "delete_private_object") as delete_private:
            ledger_app._run_claimed_invoice_analysis_job(claimed)

        stored = self._job(job_id)
        self.assertEqual(stored["status"], "completed")
        self.assertIsNotNone(stored["dismissed_at"])
        self.assertIsNone(stored["storage_key"])
        self.assertEqual(self._list().get_json()["jobs"], [])
        delete_private.assert_called_once_with("private/invoice-analysis/dismiss-test.pdf")


if __name__ == "__main__":
    unittest.main()
