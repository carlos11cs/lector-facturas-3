import unittest
import io
import zipfile

try:
    from app import (
        detect_document_type,
        enrich_payroll_data,
        extract_text_from_document_bytes,
        extract_confidence_and_fields,
        extract_tax_model_data,
        normalize_tax_period_key,
        normalize_period_label,
    )
    APP_IMPORT_ERROR = None
except Exception as exc:  # pragma: no cover - test guard for limited local envs
    APP_IMPORT_ERROR = exc


@unittest.skipIf(APP_IMPORT_ERROR is not None, f"app import failed: {APP_IMPORT_ERROR}")
class TestDocumentCenterHelpers(unittest.TestCase):
    def test_detect_document_type_for_payroll(self):
        text = "Nómina mayo 2026\nDevengos\nDeducciones\nLíquido a percibir"
        self.assertEqual(detect_document_type(text, "nomina-mayo.pdf"), "payroll")

    def test_detect_document_type_for_sales_invoice_when_company_present(self):
        text = (
            "Factura emitida\nCliente: Clínica Norte\n"
            "Ledged Consulting SL\nBase imponible 1.000,00"
        )
        self.assertEqual(
            detect_document_type(text, "factura-emitida.pdf", ["Ledged Consulting SL"]),
            "sales_invoice",
        )

    def test_extract_confidence_marks_complete_invoice_as_ready(self):
        score, status, issue_type, issue_description = extract_confidence_and_fields(
            "purchase_invoice",
            {
                "provider_name": "Proveedor Demo SL",
                "tax_id": "B12345678",
                "invoice_date": "2026-07-01",
                "invoice_number": "F-2026-001",
                "base_amount": 100.0,
                "vat_amount": 21.0,
                "total_amount": 121.0,
            },
            text_value="Factura completa con base imponible e IVA",
        )
        self.assertGreaterEqual(score, 80)
        self.assertEqual(status, "ready_to_register")
        self.assertIsNone(issue_type)
        self.assertIsNone(issue_description)

    def test_extract_confidence_marks_sparse_document_for_review(self):
        score, status, issue_type, issue_description = extract_confidence_and_fields(
            "unknown",
            {"concept": "Documento sin datos"},
            text_value="texto mínimo",
        )
        self.assertLess(score, 50)
        self.assertEqual(status, "needs_review")
        self.assertEqual(issue_type, "low_confidence")
        self.assertTrue(issue_description)

    def test_normalize_period_label_uses_trimmed_value(self):
        self.assertEqual(normalize_period_label("  Julio 2026 "), "Julio 2026")

    def test_enrich_payroll_data_derives_amounts_and_period(self):
        text = (
            "Nómina mayo 2026\n"
            "Total devengado 2.000,00\n"
            "Total deducciones 450,00\n"
            "Líquido a percibir 1.550,00\n"
            "Coste empresa 2.650,00"
        )
        enriched = enrich_payroll_data({}, text, "05.2026 Nomina Ana.pdf")
        self.assertEqual(enriched["payroll_period"], "Mayo 2026")
        self.assertEqual(enriched["base_amount"], 2000.0)
        self.assertEqual(enriched["payroll_net_amount"], 1550.0)
        self.assertEqual(enriched["payroll_total_deductions_amount"], 450.0)
        self.assertEqual(enriched["payroll_employer_cost_amount"], 2650.0)

    def test_extract_tax_model_data_reads_model_period_and_amount(self):
        text = (
            "Modelo 303\n"
            "Ejercicio 2026\n"
            "2T 2026\n"
            "Resultado de la liquidación 845,22"
        )
        extracted = extract_tax_model_data(text)
        self.assertEqual(extracted["model_name"], "303")
        self.assertTrue(extracted["tax_period"] in {"2026", "T2 2026"})
        self.assertEqual(extracted["amount"], 845.22)
        self.assertEqual(extracted["filing_status"], "a_ingresar")

    def test_extract_tax_model_data_detects_offset_status(self):
        text = (
            "Modelo 303\n"
            "3 trimestre 2026\n"
            "Importe a compensar 120,50"
        )
        extracted = extract_tax_model_data(text)
        self.assertEqual(extracted["filing_status"], "a_compensar")
        self.assertEqual(extracted["offset_amount"], 120.50)
        self.assertEqual(extracted["amount"], 120.50)

    def test_extract_tax_model_data_detects_no_activity(self):
        text = "Modelo 303\nEjercicio 2026\nSin actividad"
        extracted = extract_tax_model_data(text)
        self.assertEqual(extracted["filing_status"], "sin_actividad")
        self.assertEqual(extracted["amount"], 0.0)

    def test_normalize_tax_period_key_unifies_quarter_formats(self):
        self.assertEqual(normalize_tax_period_key("T2 2026"), "T2 2026")
        self.assertEqual(normalize_tax_period_key("2 trimestre 2026"), "T2 2026")

    def test_extract_text_from_document_bytes_falls_back_for_xlsx_archive(self):
        workbook_bytes = io.BytesIO()
        with zipfile.ZipFile(workbook_bytes, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr(
                "xl/sharedStrings.xml",
                """<?xml version="1.0" encoding="UTF-8"?>
                <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <si><t>Fecha</t></si>
                  <si><t>Importe</t></si>
                  <si><t>30/06/2026</t></si>
                  <si><t>120,55</t></si>
                </sst>""",
            )
            archive.writestr(
                "xl/worksheets/sheet1.xml",
                """<?xml version="1.0" encoding="UTF-8"?>
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                  <sheetData>
                    <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>
                    <row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2" t="s"><v>3</v></c></row>
                  </sheetData>
                </worksheet>""",
            )
        extracted = extract_text_from_document_bytes(workbook_bytes.getvalue(), "extracto.xlsx")
        self.assertIn("Fecha | Importe", extracted)
        self.assertIn("30/06/2026 | 120,55", extracted)


if __name__ == "__main__":
    unittest.main()
