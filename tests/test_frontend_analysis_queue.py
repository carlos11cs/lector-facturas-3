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


if __name__ == "__main__":
    unittest.main()
