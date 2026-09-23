import json
import unittest
from unittest.mock import patch

try:
    import app as ledger_app

    APP_IMPORT_ERROR = None
except Exception as exc:  # pragma: no cover - import guard for limited local envs
    ledger_app = None
    APP_IMPORT_ERROR = exc


@unittest.skipIf(APP_IMPORT_ERROR is not None, f"app import failed: {APP_IMPORT_ERROR}")
class TestInvoiceAnalysisJobs(unittest.TestCase):
    def test_completed_job_exposes_only_temporary_result(self):
        job = {
            "id": 7,
            "company_id": 3,
            "document_type": "expense",
            "original_filename": "factura.pdf",
            "storage_key": "private/invoice-analysis/secret.pdf",
            "status": "completed",
            "result_json": json.dumps({"provider_name": "Proveedor demo"}),
            "error_message": None,
            "created_at": "2026-09-22T10:00:00",
            "started_at": "2026-09-22T10:01:00",
            "completed_at": "2026-09-22T10:02:00",
            "expires_at": "2026-09-23T10:00:00",
        }

        serialized = ledger_app.serialize_invoice_analysis_job(job)

        self.assertEqual(serialized["result"], {"provider_name": "Proveedor demo"})
        self.assertNotIn("storage_key", serialized)
        self.assertNotIn("error_message", serialized)

    def test_persistent_queue_requires_opt_in_and_private_storage(self):
        with patch.object(ledger_app, "ASYNC_INVOICE_ANALYSIS_ENABLED", False), patch.object(
            ledger_app, "has_private_object_storage", return_value=True
        ):
            self.assertFalse(ledger_app.async_invoice_analysis_is_available())

        with patch.object(ledger_app, "ASYNC_INVOICE_ANALYSIS_ENABLED", True), patch.object(
            ledger_app, "has_private_object_storage", return_value=False
        ):
            self.assertFalse(ledger_app.async_invoice_analysis_is_available())

        with patch.object(ledger_app, "ASYNC_INVOICE_ANALYSIS_ENABLED", True), patch.object(
            ledger_app, "has_private_object_storage", return_value=True
        ):
            self.assertTrue(ledger_app.async_invoice_analysis_is_available())


if __name__ == "__main__":
    unittest.main()
