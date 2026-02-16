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

    def test_invoice_and_payment_dates(self):
        invoice_date = svc._extract_invoice_date_from_text(FIXTURE_TEXT)
        self.assertEqual(invoice_date, "2026-02-03")
        payment_dates = svc._find_payment_dates_by_keywords(FIXTURE_TEXT, invoice_date)
        self.assertIn("2026-04-04", payment_dates)

    def test_confidence_score_mapping(self):
        self.assertAlmostEqual(svc._confidence_score_for_source("regex_tax_summary"), 0.98, places=2)
        self.assertAlmostEqual(svc._confidence_score_for_source("llm"), 0.85, places=2)
        self.assertAlmostEqual(svc._confidence_score_for_source("fallback"), 0.60, places=2)


if __name__ == "__main__":
    unittest.main()
