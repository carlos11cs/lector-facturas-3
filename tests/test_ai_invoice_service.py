import unittest

from services import ai_invoice_service as svc


FIXTURE_TEXT = """Industrial Farmacéutica Cantabria, S.A.
N.I.F.: A-39000914
Barrio Solía 30, La Concha de Villaescusa
(39690), Cantabria, España
IMPUESTOS
TOTAL BRUTO
BASE IMPONIBLE
%
I.V.A.
%
REC.EQUIV
TOTAL
1.171,34
(1)
21,00
245,98
1.171,34
1.417,32 EUR
FACTURA Nº
TIPO FAC
FECHA
CLIENTE Nº
PR. ZONA
Nº ALBARAN
N.I.F.
26009701
RI
03/02/2026
235044
VENCIMIENTO:
60 dias fecha factura
04/04/2026
"""


class TestAiInvoiceService(unittest.TestCase):
    def test_parse_eu_amount(self):
        cases = {
            "1.171,34": 1171.34,
            "1 171,34": 1171.34,
            "1417,32 EUR": 1417.32,
            "245,98": 245.98,
            "0,00": 0.0,
        }
        for raw, expected in cases.items():
            parsed = svc.parse_eu_amount(raw)
            self.assertIsNotNone(parsed)
            self.assertAlmostEqual(parsed, expected, places=2)

    def test_extract_tax_summary_from_text(self):
        summary = svc._extract_tax_summary_from_text(FIXTURE_TEXT)
        self.assertTrue(summary.get("found"))
        self.assertAlmostEqual(summary.get("base_amount"), 1171.34, places=2)
        self.assertAlmostEqual(summary.get("vat_rate"), 21.0, places=2)
        self.assertAlmostEqual(summary.get("vat_amount"), 245.98, places=2)
        self.assertAlmostEqual(summary.get("total_amount"), 1417.32, places=2)

    def test_tax_summary_overrides_llm(self):
        summary = svc._extract_tax_summary_from_text(FIXTURE_TEXT)
        base, vat, total, rate, source = svc._apply_tax_summary_override(
            FIXTURE_TEXT,
            971.34,
            245.98,
            1217.32,
            21.0,
            summary,
        )
        self.assertEqual(source, "regex_tax_summary")
        self.assertAlmostEqual(base, 1171.34, places=2)
        self.assertAlmostEqual(vat, 245.98, places=2)
        self.assertAlmostEqual(total, 1417.32, places=2)
        self.assertAlmostEqual(rate, 21.0, places=2)

        corrected = svc._reconcile_vat_breakdown(
            [{"rate": 21.0, "base": 971.36, "vat_amount": 245.98, "total": 1217.32}],
            base,
            vat,
            total,
            rate,
            source,
        )
        self.assertEqual(len(corrected), 1)
        self.assertAlmostEqual(corrected[0]["base"], 1171.34, places=2)
        self.assertAlmostEqual(corrected[0]["vat_amount"], 245.98, places=2)
        self.assertAlmostEqual(corrected[0]["total"], 1417.32, places=2)

    def test_invoice_and_payment_dates(self):
        invoice_date = svc._extract_invoice_date_from_text(FIXTURE_TEXT)
        self.assertEqual(invoice_date, "2026-02-03")
        payment_dates = svc._find_payment_dates_by_keywords(FIXTURE_TEXT, invoice_date)
        self.assertIn("2026-04-04", payment_dates)

    def test_confidence_score_mapping(self):
        self.assertAlmostEqual(svc._confidence_score_for_source("regex_tax_summary"), 0.98, places=2)
        self.assertAlmostEqual(svc._confidence_score_for_source("llm"), 0.85, places=2)
        self.assertAlmostEqual(svc._confidence_score_for_source("fallback"), 0.60, places=2)

    def test_llm_amounts_trustworthy(self):
        self.assertTrue(svc._is_llm_amounts_trustworthy(362.58, 21.0, 76.14, 438.72))
        self.assertFalse(svc._is_llm_amounts_trustworthy(21.0, 21.0, 21.0, 61.98))

    def test_normalize_multivat_rates_by_math(self):
        extracted = {
            "vat_breakdown": [
                {"base": 18.86, "vat_amount": 3.96, "rate": 4},
                {"base": 8.40, "vat_amount": 0.84, "rate": 21},
            ],
            "totals": {"base": None, "vat": None, "total": None},
        }
        normalized = svc.normalize_and_validate_amounts(extracted)
        breakdown = normalized["vat_breakdown"]
        self.assertEqual(len(breakdown), 2)
        rates = sorted([line["rate"] for line in breakdown])
        self.assertEqual(rates, [10.0, 21.0])
        self.assertIsNone(normalized["vat_rate"])
        self.assertAlmostEqual(normalized["base_amount"], 27.26, places=2)
        self.assertAlmostEqual(normalized["vat_amount"], 4.8, places=2)
        self.assertAlmostEqual(normalized["total_amount"], 32.06, places=2)

    def test_totals_reconciled_from_breakdown(self):
        extracted = {
            "totals": {"base": 100.0, "vat": 10.0, "total": 110.0},
            "vat_breakdown": [
                {"base": 50.0, "vat_amount": 10.5},
                {"base": 60.0, "vat_amount": 12.6},
            ],
        }
        normalized = svc.normalize_and_validate_amounts(extracted)
        self.assertAlmostEqual(normalized["base_amount"], 110.0, places=2)
        self.assertAlmostEqual(normalized["vat_amount"], 23.1, places=2)
        self.assertAlmostEqual(normalized["total_amount"], 133.1, places=2)

    def test_partial_when_totals_incoherent(self):
        extracted = {"totals": {"base": 100.0, "vat": 10.0, "total": 50.0}}
        normalized = svc.normalize_and_validate_amounts(extracted)
        self.assertEqual(normalized["analysis_status"], "partial")
        self.assertIsNone(normalized["base_amount"])
        self.assertIsNone(normalized["vat_amount"])
        self.assertIsNone(normalized["total_amount"])

    def test_payment_terms_15_days(self):
        text = "RECIBO 15 DIAS FECHA FACTURA"
        terms = svc.extract_payment_terms_days(text)
        self.assertEqual(terms, 15)
        invoice_date = "2020-02-26"
        payment_date = (svc.date.fromisoformat(invoice_date) + svc.timedelta(days=terms)).isoformat()
        self.assertEqual(payment_date, "2020-03-12")


if __name__ == "__main__":
    unittest.main()
