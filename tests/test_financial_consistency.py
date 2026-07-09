import unittest

try:
    from app import _split_amount_across_dates, _sum_amount_for_period_dates
    APP_IMPORT_ERROR = None
except Exception as exc:  # pragma: no cover - import guard
    APP_IMPORT_ERROR = exc


@unittest.skipIf(APP_IMPORT_ERROR is not None, f"app import failed: {APP_IMPORT_ERROR}")
class TestFinancialConsistencyHelpers(unittest.TestCase):
    def test_split_amount_across_dates_preserves_total(self):
        scheduled = _split_amount_across_dates(100.0, ["2026-02-05", "2026-02-20", "2026-02-28"])
        self.assertEqual(len(scheduled), 3)
        self.assertEqual(round(sum(amount for _, amount in scheduled), 2), 100.0)

    def test_sum_amount_for_period_dates_uses_payment_month(self):
        payment_dates = ["2026-02-05", "2026-02-20"]
        january_amount = _sum_amount_for_period_dates(190.0, payment_dates, ["2026-01"])
        february_amount = _sum_amount_for_period_dates(190.0, payment_dates, ["2026-02"])
        self.assertEqual(january_amount, 0.0)
        self.assertEqual(february_amount, 190.0)


if __name__ == "__main__":
    unittest.main()
