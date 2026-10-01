import json
import unittest
from datetime import datetime, timedelta

from sqlalchemy import create_engine

import app as ledger_app
from scripts import invoice_v2_benchmark as benchmark


class TestInvoiceV2Benchmark(unittest.TestCase):
    def setUp(self):
        self.engine = create_engine("sqlite:///:memory:", future=True)
        ledger_app.metadata.create_all(self.engine)
        self.created_at = datetime(2026, 9, 29, 10, 0, 0)

    def tearDown(self):
        self.engine.dispose()

    @staticmethod
    def _v1_result(vat_amount=21.0, vat_breakdown=None, marker="PRIVATE_V1"):
        return {
            "provider_name": marker,
            "vat_amount": vat_amount,
            "vat_breakdown": vat_breakdown
            if vat_breakdown is not None
            else [{"rate": 21.0, "base": 100.0, "vat_amount": vat_amount}],
        }

    @staticmethod
    def _v2_result(vat_amount=21.0, vat_breakdown=None, marker="PRIVATE_V2"):
        return {
            "provider_name": marker,
            "vat_amount": vat_amount,
            "vat_breakdown": vat_breakdown
            if vat_breakdown is not None
            else [{"rate": 21.0, "base": 100.0, "vat_amount": vat_amount}],
        }

    def _add_run(
        self,
        *,
        batch_id="batch-a",
        created_offset=0,
        eligible=True,
        status="completed",
        validation_status="passed",
        strict_match=True,
        shadow_version="v2-sol-text-v4",
        accounting_safety_status=None,
        accounting_safety_issues=None,
        metadata_quality_status=None,
        metadata_issues=None,
        invoice_number_evidence_status=None,
        input_diagnostics=None,
        document_text_complete=None,
        document_verification=None,
        fast_path_decision=None,
        fast_path_reasons=None,
        full_document_match=None,
        full_document_comparison=None,
        comparison=None,
        validation_errors=None,
        error_type=None,
        v1_processing_ms=200,
        v2_total_ms=100,
        v1_openai_ms=160,
        v2_openai_ms=80,
        v1_tokens=None,
        v2_tokens=None,
        v1_result=None,
        v2_result=None,
    ):
        created_at = (self.created_at + timedelta(minutes=created_offset)).isoformat()
        expires_at = (self.created_at + timedelta(days=1)).isoformat()
        v1_tokens = v1_tokens or {
            "input_tokens": 1000,
            "output_tokens": 200,
            "reasoning_tokens": 100,
            "total_tokens": 1200,
        }
        v2_tokens = v2_tokens or {
            "input_tokens": 400,
            "output_tokens": 100,
            "reasoning_tokens": 50,
            "total_tokens": 550,
        }
        with self.engine.begin() as conn:
            job = conn.execute(
                ledger_app.invoice_analysis_jobs_table.insert().values(
                    user_id=7,
                    company_id=11,
                    submitted_by_user_id=7,
                    document_type="expense",
                    original_filename="never-rendered.pdf",
                    mime_type="application/pdf",
                    storage_key="private/never-rendered.pdf",
                    status="completed",
                    result_json=json.dumps(v1_result or self._v1_result()),
                    error_message=None,
                    attempt_count=1,
                    next_attempt_at=None,
                    deferred_retry_count=0,
                    lease_expires_at=None,
                    lease_token=None,
                    lease_renewal_count=0,
                    batch_id=batch_id,
                    batch_position=1,
                    created_at=created_at,
                    started_at=created_at,
                    completed_at=created_at,
                    dismissed_at=None,
                    updated_at=created_at,
                    expires_at=expires_at,
                )
            )
            job_id = job.inserted_primary_key[0]
            conn.execute(
                ledger_app.invoice_analysis_metrics_table.insert().values(
                    job_id=job_id,
                    user_id=7,
                    company_id=11,
                    document_type="expense",
                    mime_type="application/pdf",
                    file_size_bytes=100,
                    batch_id=batch_id,
                    batch_position=1,
                    queued_at=created_at,
                    started_at=created_at,
                    completed_at=created_at,
                    queue_wait_ms=0,
                    processing_ms=v1_processing_ms,
                    total_elapsed_ms=v1_processing_ms,
                    preprocessing_ms=1,
                    ocr_ms=0,
                    openai_ms=v1_openai_ms,
                    openai_model="gpt-5.6-sol",
                    ocr_used=False,
                    audit_used=False,
                    second_review_used=False,
                    processing_type="pdf",
                    analysis_route="current_full_document",
                    lease_renewal_count=0,
                    concurrency_limit_reason=None,
                    next_attempt_at=None,
                    deferred_retry_count=0,
                    rate_limit_metadata_json=None,
                    status="completed",
                    error_type=None,
                    worker_instance_id="test-worker",
                    updated_at=created_at,
                    **v1_tokens,
                )
            )
            conn.execute(
                ledger_app.invoice_analysis_shadow_runs_table.insert().values(
                    job_id=job_id,
                    user_id=7,
                    company_id=11,
                    batch_id=batch_id,
                    batch_position=1,
                    shadow_version=shadow_version,
                    route="v2_fast_text_native",
                    model="gpt-5.6-sol",
                    reasoning_effort="low",
                    eligible=eligible,
                    eligibility_reason="native_text_sufficient" if eligible else "scanned_pdf",
                    status=status,
                    validation_status=validation_status,
                    preprocessing_ms=1,
                    openai_ms=v2_openai_ms,
                    parsing_ms=1,
                    validation_ms=1,
                    total_ms=v2_total_ms,
                    mime_type="application/pdf",
                    page_count=1,
                    native_text_chars=1000,
                    sent_text_chars=300,
                    text_reduction_ratio=0.7,
                    result_json=json.dumps(v2_result or self._v2_result()),
                    provider_match=None,
                    invoice_number_match=None,
                    invoice_date_match=None,
                    tax_base_match=None,
                    vat_match=None,
                    withholding_match=None,
                    total_match=None,
                    due_date_match=None,
                    overall_match=None,
                    strict_accounting_match=strict_match,
                    comparison_json=json.dumps(comparison) if comparison is not None else None,
                    validation_errors_json=json.dumps(validation_errors)
                    if validation_errors is not None
                    else None,
                    accounting_safety_status=accounting_safety_status,
                    accounting_safety_issues_json=json.dumps(accounting_safety_issues)
                    if accounting_safety_issues is not None
                    else None,
                    metadata_quality_status=metadata_quality_status,
                    metadata_issues_json=json.dumps(metadata_issues)
                    if metadata_issues is not None
                    else None,
                    invoice_number_evidence_status=invoice_number_evidence_status,
                    invoice_parser_diagnostics_json=(
                        json.dumps({"input": input_diagnostics})
                        if input_diagnostics is not None
                        else None
                    ),
                    document_text_complete=document_text_complete,
                    document_text_chars_original=1000 if document_text_complete is not None else None,
                    document_text_chars_used=1000 if document_text_complete is not None else None,
                    document_verification_json=json.dumps(document_verification)
                    if document_verification is not None
                    else None,
                    fast_path_decision=fast_path_decision,
                    fast_path_reasons_json=json.dumps(fast_path_reasons)
                    if fast_path_reasons is not None
                    else None,
                    full_document_match=full_document_match,
                    full_document_comparison_json=json.dumps(full_document_comparison)
                    if full_document_comparison is not None
                    else None,
                    error_type=error_type,
                    attempt_count=1,
                    lease_token=None,
                    lease_expires_at=None,
                    created_at=created_at,
                    started_at=created_at,
                    completed_at=created_at,
                    updated_at=created_at,
                    worker_instance_id="test-worker",
                    **v2_tokens,
                )
            )
        return job_id

    def _rows(self, **kwargs):
        return benchmark.load_benchmark_rows(
            self.engine, shadow_version="v2-sol-text-v4", **kwargs
        )

    def test_batch_and_latest_filters_select_only_requested_shadow_runs(self):
        first = self._add_run(batch_id="batch-a", created_offset=1)
        latest = self._add_run(batch_id="batch-a", created_offset=3)
        self._add_run(batch_id="batch-b", created_offset=4)

        batch_rows = self._rows(batch_id="batch-a")
        self.assertEqual([row["job_id"] for row in batch_rows], [first, latest])
        latest_rows = self._rows(batch_id="batch-a", latest=1)
        self.assertEqual([row["job_id"] for row in latest_rows], [latest])
        ranged_rows = self._rows(job_min=latest, job_max=latest)
        self.assertEqual([row["job_id"] for row in ranged_rows], [latest])

    def test_canonical_input_diagnostics_are_reported_without_source_content(self):
        self._add_run(
            shadow_version="v2-sol-canonical-text-v1",
            input_diagnostics={
                "input_representation": "canonical_layout_v1",
                "canonical_full_chars": 1200,
                "canonical_model_chars": 1000,
                "canonical_truncated": True,
                "canonical_pages": 2,
                "canonical_token_count": 90,
                "canonical_row_count": 30,
                "canonical_segment_count": 35,
                "canonical_layout_ms": 12,
                "canonical_serialize_ms": 2,
            },
        )

        rows = benchmark.load_benchmark_rows(
            self.engine, shadow_version="v2-sol-canonical-text-v1"
        )
        report = benchmark.build_benchmark_report(
            rows, {"shadow_version": "v2-sol-canonical-text-v1"}
        )
        rendered = benchmark.render_benchmark_report(report)

        self.assertEqual(
            report["input_representation"]["counts"],
            {"canonical_layout_v1": 1},
        )
        self.assertEqual(report["input_representation"]["canonical_truncated"], 1)
        self.assertEqual(
            report["input_representation"]["canonical_layout_ms"]["mean"], 12
        )
        self.assertIn("v2-sol-canonical-text-v1", rendered)
        self.assertIn("canonical_layout_v1=1", rendered)
        self.assertNotIn("never-rendered.pdf", rendered)

    def test_report_counts_rates_validation_and_safe_candidate_without_guessing(self):
        safe_job = self._add_run()
        ineligible_job = self._add_run(
            eligible=False,
            status="skipped",
            validation_status="not_applicable",
            strict_match=None,
        )
        failed_validation_job = self._add_run(
            validation_status="failed",
            strict_match=None,
            validation_errors=["supplier_matches_registered_customer"],
        )
        mismatch_job = self._add_run(
            strict_match=False,
            comparison={
                "different_fields": [
                    "provider_match",
                    "invoice_number_match",
                    "invoice_date_match",
                    "tax_base_match",
                    "withholding_match",
                    "total_match",
                    "due_date_match",
                ],
                "supplier_tax_id_match": False,
            },
        )
        report = benchmark.build_benchmark_report(
            self._rows(), {"shadow_version": "v2-sol-text-v4"}
        )

        self.assertEqual(report["volume"]["total"], 4)
        self.assertEqual(report["volume"]["eligible"], 3)
        self.assertEqual(report["volume"]["ineligible"], 1)
        self.assertEqual(report["volume"]["validation_passed"], 2)
        self.assertEqual(report["volume"]["validation_failed"], 1)
        self.assertEqual(report["volume"]["strict_match"], 1)
        self.assertEqual(report["volume"]["strict_mismatch"], 1)
        self.assertEqual(report["safe_fast_path_candidate_count"], 1)
        self.assertEqual(report["rates"]["safe_fast_path_candidate"], 25.0)
        self.assertEqual(report["rates"]["requires_v1_fallback"], 75.0)
        self.assertEqual(
            report["validation_errors"]["supplier_matches_registered_customer"], 1
        )
        strict_fields = report["discrepancies"]["validation_passed_strict_mismatch"]["fields"]
        self.assertEqual(strict_fields["provider"]["job_ids"], [mismatch_job])
        self.assertEqual(strict_fields["supplier_tax_id"]["job_ids"], [mismatch_job])
        self.assertEqual(strict_fields["invoice_number"]["job_ids"], [mismatch_job])
        self.assertEqual(strict_fields["invoice_date"]["job_ids"], [mismatch_job])
        self.assertEqual(strict_fields["base_amount"]["job_ids"], [mismatch_job])
        self.assertEqual(strict_fields["withholding_amount"]["job_ids"], [mismatch_job])
        self.assertEqual(strict_fields["total_amount"]["job_ids"], [mismatch_job])
        self.assertEqual(strict_fields["payment_dates"]["job_ids"], [mismatch_job])
        self.assertEqual(
            report["discrepancies"]["validation_failed"]["job_ids"], [failed_validation_job]
        )
        self.assertNotIn(ineligible_job, [row["job_id"] for row in report["review_rows"]])
        self.assertIn(safe_job, [row["job_id"] for row in self._rows()])

    def test_vat_discrepancies_are_split_using_persisted_normalized_results(self):
        vat_amount_job = self._add_run(
            strict_match=False,
            comparison={"different_fields": ["vat_match"], "supplier_tax_id_match": True},
            v2_result=self._v2_result(
                vat_amount=20.0,
                vat_breakdown=[{"rate": 21.0, "base": 100.0, "vat_amount": 21.0}],
            ),
        )
        vat_breakdown_job = self._add_run(
            strict_match=False,
            comparison={"different_fields": ["vat_match"], "supplier_tax_id_match": True},
            v2_result=self._v2_result(
                vat_breakdown=[{"rate": 10.0, "base": 100.0, "vat_amount": 10.0}]
            ),
        )
        report = benchmark.build_benchmark_report(
            self._rows(), {"shadow_version": "v2-sol-text-v4"}
        )
        fields = report["discrepancies"]["validation_passed_strict_mismatch"]["fields"]

        self.assertEqual(fields["vat_amount"]["job_ids"], [vat_amount_job])
        self.assertEqual(fields["vat_breakdown"]["job_ids"], [vat_breakdown_job])

    def test_latency_percentiles_tokens_and_missing_v1_values_are_reported_without_estimates(self):
        self._add_run(v1_processing_ms=100, v2_total_ms=50, v1_openai_ms=80, v2_openai_ms=40)
        self._add_run(v1_processing_ms=200, v2_total_ms=100, v1_openai_ms=160, v2_openai_ms=80)
        self._add_run(v1_processing_ms=300, v2_total_ms=150, v1_openai_ms=240, v2_openai_ms=120)
        rows = self._rows()
        rows[0]["v1_processing_ms"] = None
        rows[0]["v1_input_tokens"] = None
        report = benchmark.build_benchmark_report(rows, {"shadow_version": "v2-sol-text-v4"})

        self.assertEqual(report["latency"]["v1_processing_ms"]["count"], 2)
        self.assertEqual(report["latency"]["v2_total_ms"]["p50"], 100.0)
        self.assertEqual(report["latency"]["v2_total_ms"]["p95"], 145.0)
        self.assertEqual(report["latency"]["paired"]["count"], 2)
        self.assertEqual(report["latency"]["paired"]["absolute_reduction"], 125.0)
        self.assertEqual(report["latency"]["paired"]["reduction_percent"], 50.0)
        self.assertEqual(report["tokens"]["input_tokens"]["v1"]["count"], 2)
        self.assertEqual(report["tokens"]["input_tokens"]["paired"]["count"], 2)

    def test_rendered_report_does_not_expose_document_payloads_or_filenames(self):
        self._add_run(
            strict_match=False,
            comparison={"different_fields": ["provider_match"], "supplier_tax_id_match": True},
            v1_result=self._v1_result(marker="PRIVATE_V1_DO_NOT_PRINT"),
            v2_result=self._v2_result(marker="PRIVATE_V2_DO_NOT_PRINT"),
        )
        report = benchmark.build_benchmark_report(
            self._rows(), {"shadow_version": "v2-sol-text-v4"}
        )
        output = benchmark.render_benchmark_report(report)

        self.assertIn("Job 1", output)
        self.assertNotIn("PRIVATE_V1_DO_NOT_PRINT", output)
        self.assertNotIn("PRIVATE_V2_DO_NOT_PRINT", output)
        self.assertNotIn("never-rendered.pdf", output)
        self.assertNotIn("private/never-rendered.pdf", output)

    def test_v5_report_separates_accounting_safety_and_metadata_quality(self):
        fully_confirmed = self._add_run(
            shadow_version="v2-sol-text-v5",
            accounting_safety_status="passed",
            accounting_safety_issues=[],
            metadata_quality_status="confirmed",
            metadata_issues=[],
            invoice_number_evidence_status="confirmed",
        )
        metadata_review = self._add_run(
            shadow_version="v2-sol-text-v5",
            validation_status="failed",
            strict_match=False,
            accounting_safety_status="passed",
            accounting_safety_issues=[],
            metadata_quality_status="review_required",
            metadata_issues=["invoice_number_evidence_missing"],
            invoice_number_evidence_status="missing",
        )
        accounting_failed = self._add_run(
            shadow_version="v2-sol-text-v5",
            validation_status="failed",
            strict_match=None,
            accounting_safety_status="failed",
            accounting_safety_issues=["tax_base_mismatch"],
            metadata_quality_status="confirmed",
            metadata_issues=[],
            invoice_number_evidence_status="confirmed",
        )

        rows = benchmark.load_benchmark_rows(
            self.engine, shadow_version="v2-sol-text-v5"
        )
        report = benchmark.build_benchmark_report(
            rows, {"shadow_version": "v2-sol-text-v5"}
        )

        self.assertEqual(report["volume"]["accounting_safety_passed"], 2)
        self.assertEqual(report["volume"]["accounting_safety_failed"], 1)
        self.assertEqual(report["volume"]["metadata_confirmed"], 2)
        self.assertEqual(report["volume"]["metadata_review_required"], 1)
        self.assertEqual(report["volume"]["fully_confirmed_fast_path_candidate"], 1)
        self.assertEqual(report["volume"]["accounting_safe_metadata_review"], 1)
        self.assertEqual(report["fully_confirmed_fast_path_candidate_count"], 1)
        self.assertAlmostEqual(
            report["rates"]["fully_confirmed_fast_path_candidate"], 100 / 3
        )
        self.assertNotIn(fully_confirmed, [row["job_id"] for row in report["review_rows"]])
        self.assertIn(metadata_review, [row["job_id"] for row in report["review_rows"]])
        self.assertIn(accounting_failed, [row["job_id"] for row in report["review_rows"]])

    def test_v7_rates_use_comparable_denominators_and_never_exceed_one_hundred_percent(self):
        self._add_run(
            shadow_version="v2-sol-text-v10",
            validation_status="passed",
            strict_match=True,
            accounting_safety_status="passed",
            accounting_safety_issues=[],
            metadata_quality_status="confirmed",
            metadata_issues=[],
            invoice_number_evidence_status="confirmed",
        )
        self._add_run(
            shadow_version="v2-sol-text-v10",
            validation_status="failed",
            strict_match=True,
            accounting_safety_status="passed",
            accounting_safety_issues=[],
            metadata_quality_status="review_required",
            metadata_issues=["invoice_number_evidence_missing"],
            invoice_number_evidence_status="missing",
        )
        self._add_run(
            shadow_version="v2-sol-text-v10",
            validation_status="passed",
            strict_match=False,
            accounting_safety_status="passed",
            accounting_safety_issues=[],
            metadata_quality_status="confirmed",
            metadata_issues=[],
            invoice_number_evidence_status="confirmed",
        )
        self._add_run(
            shadow_version="v2-sol-text-v10",
            validation_status="failed",
            strict_match=None,
            accounting_safety_status="failed",
            accounting_safety_issues=["invalid_invoice_date"],
            metadata_quality_status="confirmed",
            metadata_issues=[],
            invoice_number_evidence_status="confirmed",
        )

        report = benchmark.build_benchmark_report(
            benchmark.load_benchmark_rows(self.engine, shadow_version="v2-sol-text-v10"),
            {"shadow_version": "v2-sol-text-v10"},
        )
        rates = report["rates"]

        self.assertAlmostEqual(rates["strict_match_of_validation_passed"], 50.0)
        self.assertAlmostEqual(rates["strict_match_of_accounting_safe"], 200 / 3)
        self.assertAlmostEqual(rates["accounting_safety_passed_of_completed"], 75.0)
        self.assertAlmostEqual(rates["metadata_confirmed_of_evaluated"], 75.0)
        self.assertAlmostEqual(rates["fully_confirmed_fast_path_candidate"], 25.0)
        self.assertEqual(report["volume"]["fallback_accounting"], 1)
        self.assertEqual(report["volume"]["fallback_metadata"], 1)
        self.assertEqual(report["volume"]["fallback_unassessed_or_mismatch"], 1)
        self.assertTrue(all(0 <= value <= 100 for value in rates.values() if value is not None))

    def test_v11_report_separates_shadow_decision_verification_and_full_comparison(self):
        accepted = self._add_run(
            shadow_version="v2-sol-text-v11",
            fast_path_decision="accept_v2",
            document_text_complete=True,
            document_verification={
                "base_amount": {"status": "confirmed"},
                "currency": {"status": "confirmed"},
            },
            full_document_match=True,
            full_document_comparison={"different_fields": []},
        )
        fallback = self._add_run(
            shadow_version="v2-sol-text-v11",
            fast_path_decision="fallback_v1",
            fast_path_reasons=[
                "document_text:document_text_truncated",
                "currency:currency_missing_or_unsupported",
            ],
            document_text_complete=False,
            document_verification={
                "document_text": {"status": "review"},
                "currency": {"status": "review"},
            },
            full_document_match=False,
            full_document_comparison={
                "different_fields": ["other_taxes_match", "currency_match"]
            },
        )

        report = benchmark.build_benchmark_report(
            benchmark.load_benchmark_rows(self.engine, shadow_version="v2-sol-text-v11"),
            {"shadow_version": "v2-sol-text-v11"},
        )

        self.assertEqual(report["volume"]["accept_v2_shadow"], 1)
        self.assertEqual(report["volume"]["fallback_v1_shadow"], 1)
        self.assertEqual(report["volume"]["document_text_truncated"], 1)
        self.assertEqual(report["rates"]["accept_v2_shadow_of_eligible_completed"], 50.0)
        self.assertEqual(report["rates"]["full_document_match_of_accept_v2_shadow"], 100.0)
        self.assertEqual(report["fast_path"]["reasons"]["currency:currency_missing_or_unsupported"], 1)
        self.assertEqual(
            report["fast_path"]["verification_failures"]["currency"]["job_ids"], [fallback]
        )
        self.assertEqual(
            report["fast_path"]["full_document_discrepancies"]["currency"]["job_ids"],
            [fallback],
        )
        self.assertIn(accepted, [row["job_id"] for row in benchmark.load_benchmark_rows(
            self.engine, shadow_version="v2-sol-text-v11"
        )])


if __name__ == "__main__":
    unittest.main()
