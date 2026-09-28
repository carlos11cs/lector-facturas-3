import os
import tempfile
import threading
import time
import unittest
from concurrent.futures import ThreadPoolExecutor
from unittest.mock import patch

from sqlalchemy import create_engine, select
from sqlalchemy.dialects import postgresql

import app as ledger_app


class TestInvoiceAnalysisConcurrency(unittest.TestCase):
    def setUp(self):
        self.engine = create_engine("sqlite:///:memory:", future=True)
        self.engine_patch = patch.object(ledger_app, "engine", self.engine)
        self.engine_patch.start()
        ledger_app.metadata.create_all(self.engine)

    def tearDown(self):
        self.engine_patch.stop()
        self.engine.dispose()

    def _enqueue(
        self,
        *,
        company_id=11,
        created_at="2026-09-25T09:00:00",
        batch_id=None,
        batch_position=None,
    ):
        with self.engine.begin() as conn:
            result = conn.execute(
                ledger_app.invoice_analysis_jobs_table.insert().values(
                    user_id=7,
                    company_id=company_id,
                    submitted_by_user_id=7,
                    document_type="expense",
                    original_filename=f"factura-{company_id}-{created_at}.pdf",
                    mime_type="application/pdf",
                    storage_key=f"private/invoice-analysis/{company_id}-{created_at}.pdf",
                    status="queued",
                    result_json=None,
                    error_message=None,
                    attempt_count=0,
                    lease_expires_at=None,
                    lease_token=None,
                    lease_renewal_count=0,
                    batch_id=batch_id,
                    batch_position=batch_position,
                    created_at=created_at,
                    started_at=None,
                    completed_at=None,
                    updated_at=created_at,
                    expires_at="2026-10-01T10:00:00",
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
                queued_at=created_at,
                batch_id=batch_id,
                batch_position=batch_position,
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

    def test_claim_uses_postgresql_skip_locked(self):
        statement = ledger_app._invoice_analysis_claim_statement(
            "2026-09-25T10:00:00"
        )
        compiled = str(
            statement.compile(dialect=postgresql.dialect(), compile_kwargs={"literal_binds": True})
        )

        self.assertIn("FOR UPDATE OF invoice_analysis_jobs SKIP LOCKED", compiled)

    def test_two_claim_attempts_cannot_take_the_same_job(self):
        job_id = self._enqueue()

        first = ledger_app.claim_next_invoice_analysis_job()
        second = ledger_app.claim_next_invoice_analysis_job()

        self.assertEqual(first["id"], job_id)
        self.assertIsNotNone(first["lease_token"])
        self.assertIsNone(second)
        self.assertEqual(self._job(job_id)["status"], "processing")

    def test_parallel_sqlite_claimers_only_receive_one_copy(self):
        # SQLite has no SKIP LOCKED, but its conditional update still ensures
        # the local development fallback never returns the same job twice.
        self.engine_patch.stop()
        self.engine.dispose()
        handle = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
        handle.close()
        engine = create_engine(
            f"sqlite:///{handle.name}",
            future=True,
            connect_args={"check_same_thread": False, "timeout": 5},
        )
        self.engine = engine
        self.engine_patch = patch.object(ledger_app, "engine", engine)
        self.engine_patch.start()
        ledger_app.metadata.create_all(engine)
        job_id = self._enqueue()
        barrier = threading.Barrier(2)

        def claim():
            barrier.wait()
            try:
                return ledger_app.claim_next_invoice_analysis_job()
            except Exception:
                # SQLite may report a transient write lock under a genuine
                # race. PostgreSQL uses SKIP LOCKED and does not have this path.
                return None

        try:
            with ThreadPoolExecutor(max_workers=2) as executor:
                results = list(executor.map(lambda _unused: claim(), range(2)))
            claimed_ids = [result["id"] for result in results if result]
            self.assertEqual(claimed_ids, [job_id])
        finally:
            engine.dispose()
            os.unlink(handle.name)

    def test_several_jobs_can_be_claimed_without_duplication(self):
        job_ids = [
            self._enqueue(company_id=11, created_at=f"2026-09-25T09:00:0{index}")
            for index in range(3)
        ]
        with patch.object(ledger_app, "COMPANY_CONCURRENCY", 3):
            claims = [ledger_app.claim_next_invoice_analysis_job() for _ in job_ids]

        self.assertEqual({claim["id"] for claim in claims if claim}, set(job_ids))
        self.assertEqual(len({claim["lease_token"] for claim in claims if claim}), 3)

    def test_expired_lease_is_reclaimed_with_a_new_token(self):
        job_id = self._enqueue()
        old_claim = ledger_app.claim_next_invoice_analysis_job()
        with self.engine.begin() as conn:
            conn.execute(
                ledger_app.invoice_analysis_jobs_table.update()
                .where(ledger_app.invoice_analysis_jobs_table.c.id == job_id)
                .values(lease_expires_at="2020-01-01T00:00:00")
            )

        recovered = ledger_app.claim_next_invoice_analysis_job()

        self.assertEqual(recovered["id"], job_id)
        self.assertNotEqual(old_claim["lease_token"], recovered["lease_token"])
        self.assertEqual(self._job(job_id)["attempt_count"], 2)

    def test_stale_worker_cannot_finish_a_recovered_job(self):
        job_id = self._enqueue()
        old_claim = ledger_app.claim_next_invoice_analysis_job()
        with self.engine.begin() as conn:
            conn.execute(
                ledger_app.invoice_analysis_jobs_table.update()
                .where(ledger_app.invoice_analysis_jobs_table.c.id == job_id)
                .values(lease_expires_at="2020-01-01T00:00:00")
            )
        recovered = ledger_app.claim_next_invoice_analysis_job()

        accepted = ledger_app._finish_invoice_analysis_job(
            old_claim,
            status="completed",
            completed_at_iso="2026-09-25T10:00:00",
            extracted={"provider_name": "stale"},
        )

        job = self._job(job_id)
        self.assertFalse(accepted)
        self.assertEqual(job["status"], "processing")
        self.assertEqual(job["lease_token"], recovered["lease_token"])
        self.assertIsNone(job["result_json"])

    def test_renewal_requires_the_current_token_and_is_measured(self):
        job_id = self._enqueue()
        claim = ledger_app.claim_next_invoice_analysis_job()

        renewal_count = ledger_app.renew_invoice_analysis_lease(
            job_id, claim["lease_token"]
        )
        rejected = ledger_app.renew_invoice_analysis_lease(job_id, "stale-token")

        self.assertEqual(renewal_count, 1)
        self.assertIsNone(rejected)
        self.assertEqual(self._job(job_id)["lease_renewal_count"], 1)
        self.assertEqual(self._metric(job_id)["lease_renewal_count"], 1)

    def test_lease_renewer_stops_after_completion_path_requests_stop(self):
        renewer = ledger_app._InvoiceAnalysisLeaseRenewer(8, "lease-token")
        with patch.object(
            ledger_app, "ASYNC_INVOICE_ANALYSIS_LEASE_RENEWAL_SECONDS", 0.01
        ), patch.object(
            ledger_app, "renew_invoice_analysis_lease", return_value=1
        ) as renew:
            renewer.start()
            time.sleep(0.04)
            renewer.stop()
            calls_at_stop = renew.call_count
            time.sleep(0.03)

        self.assertGreaterEqual(calls_at_stop, 1)
        self.assertEqual(renew.call_count, calls_at_stop)

    def test_company_fairness_reserves_capacity_for_another_company(self):
        first_company_job = self._enqueue(company_id=11, created_at="2026-09-25T09:00:00")
        self._enqueue(company_id=11, created_at="2026-09-25T09:00:01")
        other_company_job = self._enqueue(company_id=22, created_at="2026-09-25T09:00:02")

        with patch.object(ledger_app, "COMPANY_CONCURRENCY", 1):
            first = ledger_app.claim_next_invoice_analysis_job()
            second = ledger_app.claim_next_invoice_analysis_job()

        self.assertEqual(first["id"], first_company_job)
        self.assertEqual(second["id"], other_company_job)

    def test_batch_serialization_and_paginated_endpoint_are_scoped_to_company(self):
        batch_id = "batch20260925"
        second_position_job_id = self._enqueue(batch_id=batch_id, batch_position=2)
        self._enqueue(batch_id=batch_id, batch_position=1)
        self._enqueue(company_id=22, batch_id=batch_id, batch_position=1)

        with patch.object(ledger_app, "get_data_owner_id", return_value=7), patch.object(
            ledger_app, "get_company_id", return_value=11
        ), ledger_app.app.test_request_context(
            f"/api/invoice-analysis-batches/{batch_id}?page=1&page_size=1"
        ):
            response = ledger_app.get_invoice_analysis_batch(batch_id)

        data = response.get_json()
        self.assertTrue(data["ok"])
        self.assertEqual(data["total"], 2)
        self.assertTrue(data["hasMore"])
        self.assertEqual(data["jobs"][0]["batchPosition"], 1)
        self.assertEqual(data["statusCounts"], {"queued": 2})
        self.assertEqual(self._metric(second_position_job_id)["batch_id"], batch_id)
        self.assertEqual(self._metric(second_position_job_id)["batch_position"], 2)

    def test_safe_defaults_preserve_single_worker_production_rollout(self):
        self.assertEqual(ledger_app.WORKER_CONCURRENCY, 1)
        self.assertEqual(ledger_app.FULL_DOCUMENT_CONCURRENCY, 1)
        self.assertEqual(ledger_app.OCR_CONCURRENCY, 1)
        self.assertEqual(ledger_app.COMPANY_CONCURRENCY, 2)

    def test_concurrency_limit_reason_is_preserved_in_historical_metrics(self):
        job_id = self._enqueue()
        with self.engine.begin() as conn:
            ledger_app._mark_invoice_analysis_metrics_processing(
                conn, job_id, "2026-09-25T09:00:01"
            )
        ledger_app._set_invoice_analysis_concurrency_limit(
            job_id, "full_document_concurrency"
        )
        with self.engine.begin() as conn:
            ledger_app._complete_invoice_analysis_metrics(
                conn,
                job_id=job_id,
                status="completed",
                completed_at="2026-09-25T09:00:02",
            )

        self.assertEqual(
            self._metric(job_id)["concurrency_limit_reason"],
            "full_document_concurrency",
        )


if __name__ == "__main__":
    unittest.main()
