import json
import unittest
from datetime import datetime, timedelta
from unittest.mock import patch

from sqlalchemy import create_engine, select

import app as ledger_app


class TestInvoiceAnalysisRateLimits(unittest.TestCase):
    def setUp(self):
        self.engine = create_engine("sqlite:///:memory:", future=True)
        self.engine_patch = patch.object(ledger_app, "engine", self.engine)
        self.engine_patch.start()
        ledger_app.metadata.create_all(self.engine)

    def tearDown(self):
        self.engine_patch.stop()
        self.engine.dispose()

    def _enqueue(self, *, company_id=11, next_attempt_at=None, deferred_retry_count=0):
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
                    storage_key="private/invoice-analysis/test.pdf",
                    status="queued",
                    result_json=None,
                    error_message=None,
                    attempt_count=0,
                    next_attempt_at=next_attempt_at,
                    deferred_retry_count=deferred_retry_count,
                    lease_expires_at=None,
                    lease_token=None,
                    lease_renewal_count=0,
                    batch_id=None,
                    batch_position=None,
                    created_at=now,
                    started_at=None,
                    completed_at=None,
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
            )
        return job_id

    def _job(self, job_id):
        with self.engine.connect() as conn:
            return conn.execute(
                select(ledger_app.invoice_analysis_jobs_table).where(
                    ledger_app.invoice_analysis_jobs_table.c.id == job_id
                )
            ).mappings().one()

    def _metric(self, job_id):
        with self.engine.connect() as conn:
            return conn.execute(
                select(ledger_app.invoice_analysis_metrics_table).where(
                    ledger_app.invoice_analysis_metrics_table.c.job_id == job_id
                )
            ).mappings().one()

    def _processing_job(self, *, deferred_retry_count=0):
        job_id = self._enqueue(deferred_retry_count=deferred_retry_count)
        with self.engine.begin() as conn:
            conn.execute(
                ledger_app.invoice_analysis_jobs_table.update()
                .where(ledger_app.invoice_analysis_jobs_table.c.id == job_id)
                .values(status="processing", started_at=datetime.utcnow().isoformat())
            )
            ledger_app._mark_invoice_analysis_metrics_processing(
                conn, job_id, datetime.utcnow().isoformat()
            )
        return dict(self._job(job_id))

    @staticmethod
    def _rate_limited_failure(*, retryable=True, kind="rpm", retry_after=None):
        metadata = {
            "http_status": 429,
            "error_class": "RateLimitError",
            "error_code": "rate_limit_exceeded",
            "error_type": "requests",
            "request_id": "req_429_test",
            "rate_limit_kind": kind,
            "retryable": retryable,
        }
        if retry_after is not None:
            metadata["retry_after_seconds"] = retry_after
        return {
            "analysis_status": "failed",
            "analysis_error": {
                "status": "rate_limited",
                "detail": "rate_limit_exceeded",
                "metadata": metadata,
            },
        }

    def _run_worker(self, job, extracted, delete_private):
        with patch.object(ledger_app, "download_private_bytes", return_value=b"%PDF-test"), patch.object(
            ledger_app, "get_company_names_for_analysis", return_value=[]
        ), patch.object(ledger_app, "fetch_known_suppliers", return_value=[]), patch.object(
            ledger_app,
            "_analyze_invoice_with_timeout",
            return_value=(extracted, {"openai_ms": 10}),
        ):
            ledger_app._run_claimed_invoice_analysis_job(job)

    def test_rate_limit_is_requeued_and_preserves_private_source(self):
        job = self._processing_job()
        with patch.object(ledger_app.random, "uniform", return_value=1.0), patch.object(
            ledger_app, "delete_private_object"
        ) as delete_private:
            self._run_worker(job, self._rate_limited_failure(), delete_private)

        stored = self._job(job["id"])
        metric = self._metric(job["id"])
        self.assertEqual(stored["status"], "queued")
        self.assertEqual(stored["deferred_retry_count"], 1)
        self.assertIsNotNone(stored["next_attempt_at"])
        self.assertEqual(stored["storage_key"], "private/invoice-analysis/test.pdf")
        self.assertEqual(metric["status"], "retrying")
        self.assertEqual(metric["error_type"], "rate_limited")
        self.assertEqual(metric["deferred_retry_count"], 1)
        metadata = json.loads(metric["rate_limit_metadata_json"])
        self.assertEqual(metadata["http_status"], 429)
        self.assertEqual(metadata["request_id"], "req_429_test")
        self.assertNotIn("original_filename", metadata)
        self.assertNotIn("storage_key", metadata)
        delete_private.assert_not_called()

    def test_retry_after_has_priority_over_jittered_backoff(self):
        job_id = self._enqueue()
        job = ledger_app.claim_next_invoice_analysis_job()
        now = datetime(2026, 9, 28, 10, 0, 0)
        self.assertEqual(job["id"], job_id)

        with patch.object(ledger_app.random, "uniform", side_effect=AssertionError("no jitter")):
            self.assertTrue(
                ledger_app._defer_rate_limited_invoice_analysis_job(
                    job,
                    metadata=self._rate_limited_failure(retry_after=17)["analysis_error"]["metadata"],
                    now=now,
                )
            )

        self.assertEqual(
            self._job(job_id)["next_attempt_at"],
            (now + timedelta(seconds=17)).isoformat(),
        )

    def test_backoff_without_retry_after_uses_the_configured_sequence(self):
        with patch.object(ledger_app.random, "uniform", return_value=1.0):
            delays = [
                ledger_app._rate_limit_retry_delay_seconds(retry_count, {})
                for retry_count in range(1, 5)
            ]
        self.assertEqual(delays, [30, 60, 120, 300])

    def test_future_retry_is_not_claimed_before_next_attempt(self):
        delayed_job_id = self._enqueue(
            next_attempt_at=(datetime.utcnow() + timedelta(minutes=10)).isoformat(),
            deferred_retry_count=1,
        )
        ready_job_id = self._enqueue(company_id=22)

        claim = ledger_app.claim_next_invoice_analysis_job()

        self.assertEqual(claim["id"], ready_job_id)
        self.assertEqual(self._job(delayed_job_id)["status"], "queued")

    def test_rate_limit_is_failed_after_the_deferred_retry_budget_and_source_is_deleted(self):
        job = self._processing_job(
            deferred_retry_count=ledger_app.MAX_RATE_LIMIT_DEFERRED_RETRIES
        )
        with patch.object(ledger_app, "delete_private_object") as delete_private:
            self._run_worker(job, self._rate_limited_failure(), delete_private)

        stored = self._job(job["id"])
        metric = self._metric(job["id"])
        self.assertEqual(stored["status"], "failed")
        self.assertIsNone(stored["storage_key"])
        self.assertEqual(metric["status"], "failed")
        self.assertEqual(metric["error_type"], "rate_limited")
        delete_private.assert_called_once_with("private/invoice-analysis/test.pdf")

    def test_quota_limit_is_failed_without_retry_and_source_is_deleted(self):
        job = self._processing_job()
        with patch.object(ledger_app, "delete_private_object") as delete_private:
            self._run_worker(
                job,
                self._rate_limited_failure(
                    retryable=False, kind="insufficient_quota"
                ),
                delete_private,
            )

        stored = self._job(job["id"])
        metric = self._metric(job["id"])
        self.assertEqual(stored["status"], "failed")
        self.assertEqual(metric["status"], "failed")
        self.assertEqual(metric["error_type"], "insufficient_quota")
        self.assertEqual(stored["deferred_retry_count"], 0)
        self.assertEqual(
            json.loads(metric["rate_limit_metadata_json"])["rate_limit_kind"],
            "insufficient_quota",
        )
        delete_private.assert_called_once_with("private/invoice-analysis/test.pdf")

    def test_generic_failed_analysis_never_becomes_completed(self):
        job = self._processing_job()
        extracted = {
            "analysis_status": "failed",
            "analysis_error": {"status": "api_error", "detail": "APIError"},
        }
        with patch.object(ledger_app, "delete_private_object") as delete_private:
            self._run_worker(job, extracted, delete_private)

        self.assertEqual(self._job(job["id"])["status"], "failed")
        metric = self._metric(job["id"])
        self.assertEqual(metric["status"], "failed")
        self.assertEqual(metric["error_type"], "api_error")
        delete_private.assert_called_once_with("private/invoice-analysis/test.pdf")

    def test_stale_worker_cannot_defer_a_recovered_lease(self):
        job_id = self._enqueue()
        old_claim = ledger_app.claim_next_invoice_analysis_job()
        with self.engine.begin() as conn:
            conn.execute(
                ledger_app.invoice_analysis_jobs_table.update()
                .where(ledger_app.invoice_analysis_jobs_table.c.id == job_id)
                .values(lease_expires_at="2020-01-01T00:00:00")
            )
        recovered = ledger_app.claim_next_invoice_analysis_job()

        accepted = ledger_app._defer_rate_limited_invoice_analysis_job(
            old_claim,
            metadata=self._rate_limited_failure()["analysis_error"]["metadata"],
        )

        self.assertFalse(accepted)
        stored = self._job(job_id)
        self.assertEqual(stored["status"], "processing")
        self.assertEqual(stored["lease_token"], recovered["lease_token"])


if __name__ == "__main__":
    unittest.main()
