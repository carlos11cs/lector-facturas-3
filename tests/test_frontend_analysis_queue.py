import unittest
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class TestFrontendAnalysisQueue(unittest.TestCase):
    def test_browser_timeout_outlives_server_analysis_and_starts_when_processing(self):
        script = (PROJECT_ROOT / "static/js/app.js").read_text(encoding="utf-8")

        self.assertIn("const ANALYSIS_PENDING_TIMEOUT_MS = 610 * 1000;", script)
        processing_marker = "task.item.analysisQueued = false;\n    scheduleAnalysisTimeout(task.item, task.render);"
        self.assertIn(processing_marker, script)
        enqueue_block = script.split("function enqueueAnalysisTask", 1)[1].split("function showLowQualityModal", 1)[0]
        self.assertNotIn("scheduleAnalysisTimeout(item, render);", enqueue_block)

    def test_multiple_uploads_use_configurable_concurrency_and_batch_polling(self):
        script = (PROJECT_ROOT / "static/js/app.js").read_text(encoding="utf-8")

        self.assertIn("window.LEDGED_ANALYSIS_MAX_CONCURRENCY || 2", script)
        self.assertIn("const analysisBatchPolls = new Map();", script)
        self.assertIn("function createAnalysisBatchId()", script)
        self.assertIn('formData.append("batch_id", item.analysisBatchId);', script)
        self.assertIn("/api/invoice-analysis-batches/${encodeURIComponent(batchId)}", script)

    def test_failed_or_deferred_persistent_jobs_are_not_applied_as_invoices(self):
        script = (PROJECT_ROOT / "static/js/app.js").read_text(encoding="utf-8")

        self.assertIn('if (item.analysisStatus === "failed")', script)
        self.assertIn("item.analysisRetryAt = job.nextAttemptAt || null;", script)
        self.assertIn("const RETRYING_ANALYSIS_MESSAGE", script)
        self.assertIn("getAnalysisQueueMessage(", script)
        self.assertIn("retryAtMs - Date.now() + ANALYSIS_PENDING_TIMEOUT_MS", script)

    def test_remove_uses_soft_dismiss_for_persistent_jobs(self):
        script = (PROJECT_ROOT / "static/js/app.js").read_text(encoding="utf-8")

        self.assertIn("function dismissPersistentAnalysisItem(item)", script)
        self.assertIn('method: "DELETE"', script)
        self.assertIn('removeBtn.textContent = "Quitar";', script)
        self.assertNotIn("function abortPendingAnalysis(item)", script)


if __name__ == "__main__":
    unittest.main()
