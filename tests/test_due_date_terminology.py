import unittest
from pathlib import Path

try:
    from app import parse_payment_dates, replace_payment_date, resolve_payment_status, serialize_payment_dates
    APP_IMPORT_ERROR = None
except Exception as exc:  # pragma: no cover - import guard
    APP_IMPORT_ERROR = exc


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class TestDueDateTerminology(unittest.TestCase):
    def test_visible_due_date_labels_use_vencimiento(self):
        script = (PROJECT_ROOT / "static/js/app.js").read_text(encoding="utf-8")
        template = (PROJECT_ROOT / "templates/index.html").read_text(encoding="utf-8")

        self.assertEqual(script.count('paymentLabel.textContent = "Fechas de vencimiento";'), 2)
        self.assertNotIn('paymentLabel.textContent = "Fechas de pago";', script)
        self.assertIn("fecha de vencimiento", template)
        self.assertNotIn("fecha de pago", template)


@unittest.skipIf(APP_IMPORT_ERROR is not None, f"app import failed: {APP_IMPORT_ERROR}")
class TestLegacyDueDateCompatibility(unittest.TestCase):
    def test_legacy_payment_dates_remain_editable_due_dates(self):
        legacy_dates = '["2026-10-10", "2026-11-09"]'
        updated_dates = replace_payment_date(
            parse_payment_dates(legacy_dates), "2026-11-09", "2026-11-12"
        )

        self.assertEqual(updated_dates, ["2026-10-10", "2026-11-12"])
        self.assertEqual(
            parse_payment_dates(serialize_payment_dates(updated_dates)), updated_dates
        )

    def test_completed_date_marks_a_due_date_paid_without_changing_it(self):
        due_date = "2026-10-10"

        self.assertEqual(resolve_payment_status("expense", due_date, [], "2026-10-11"), "overdue")
        self.assertEqual(
            resolve_payment_status("expense", due_date, [due_date], "2026-10-11"),
            "paid",
        )


if __name__ == "__main__":
    unittest.main()
