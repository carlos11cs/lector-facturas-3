import unittest
from datetime import date
from io import BytesIO

try:
    from app import (
        parse_accounting_import_file,
        serialize_accounting_import_preview,
    )

    APP_IMPORT_ERROR = None
except Exception as exc:  # pragma: no cover - import guard for limited local envs
    APP_IMPORT_ERROR = exc


@unittest.skipIf(APP_IMPORT_ERROR is not None, f"app import failed: {APP_IMPORT_ERROR}")
class TestAccountingImportHelpers(unittest.TestCase):
    def test_purchase_csv_uses_absolute_withholding_and_net_payable_total(self):
        csv_content = "\n".join(
            [
                "Fecha factura;Proveedor;Concepto;Número factura;Base imponible;Tipo IVA;Cuota IVA;Retención IRPF;Total;Vencimiento",
                "31/08/2026;Profesional sanitario;Actividad médica;F-2026-1;2264,15;21;475,47;-339,62;2400,00;01/10/2026",
            ]
        ).encode("utf-8")

        parsed = parse_accounting_import_file(
            csv_content, "compras.csv", "purchases", "a3"
        )

        self.assertEqual(parsed["columns"]["date"], 0)
        self.assertEqual(parsed["columns"]["counterparty"], 1)
        record = parsed["records"][0]
        self.assertEqual(record["invoice_date"], "2026-08-31")
        self.assertEqual(record["withholding_amount"], 339.62)
        self.assertEqual(record["vat_amount"], 475.47)
        self.assertEqual(record["total_amount"], 2400.00)
        self.assertEqual(record["payment_dates"], ["2026-10-01"])
        self.assertEqual(record["errors"], [])

    def test_import_derives_allowed_vat_rate_from_vat_amount(self):
        csv_content = "\n".join(
            [
                "Fecha;Proveedor;Base;Importe IVA;Total",
                "2026-09-10;Proveedor demo;1000,00;210,00;1210,00",
            ]
        ).encode("utf-8")

        record = parse_accounting_import_file(
            csv_content, "contasol.csv", "purchases", "contasol"
        )["records"][0]

        self.assertEqual(record["vat_rate"], 21)
        self.assertEqual(record["vat_amount"], 210.00)
        self.assertEqual(record["errors"], [])

    def test_ledged_purchase_export_headers_round_trip(self):
        csv_content = "\n".join(
            [
                "fecha;documento_tipo;origen_tipo;origen_id;contraparte;concepto;base;iva;retencion;total",
                "2026-09-10;Factura proveedor;invoice;12;Proveedor demo;Servicio;1000,00;210,00;0,00;1210,00",
            ]
        ).encode("utf-8")

        record = parse_accounting_import_file(
            csv_content, "export_ledged.csv", "purchases", "generic"
        )["records"][0]

        self.assertEqual(record["counterparty"], "Proveedor demo")
        self.assertEqual(record["vat_rate"], 21)
        self.assertEqual(record["errors"], [])

    def test_import_rejects_total_that_does_not_match_tax_calculation(self):
        csv_content = "\n".join(
            [
                "Fecha;Proveedor;Base imponible;Tipo IVA;Total",
                "2026-09-10;Proveedor demo;1000,00;21;1250,00",
            ]
        ).encode("utf-8")

        record = parse_accounting_import_file(
            csv_content, "a3.csv", "purchases", "a3"
        )["records"][0]

        self.assertIn("El total no cuadra con base, IVA y retención.", record["errors"])

    def test_import_marks_negative_amounts_as_rectificative(self):
        csv_content = "\n".join(
            [
                "Fecha;Proveedor;Base imponible;Tipo IVA;Cuota IVA;Total",
                "2026-09-10;Proveedor demo;-100,00;21;-21,00;-121,00",
            ]
        ).encode("utf-8")

        record = parse_accounting_import_file(
            csv_content, "abono.csv", "purchases", "generic"
        )["records"][0]

        self.assertTrue(record["is_rectificativa"])
        self.assertEqual(record["errors"], [])

    def test_sales_withholding_is_preserved_and_reduces_receivable_total(self):
        csv_content = "\n".join(
            [
                "Fecha;Cliente;Base imponible;Tipo IVA;Cuota IVA;Retención IRPF;Total",
                "2026-09-10;Cliente demo;100,00;21;21,00;15,00;106,00",
            ]
        ).encode("utf-8")

        record = parse_accounting_import_file(
            csv_content, "ventas.csv", "sales", "generic"
        )["records"][0]

        self.assertEqual(record["withholding_amount"], 15.00)
        self.assertEqual(record["total_amount"], 106.00)
        self.assertEqual(record["errors"], [])

    def test_import_preserves_multiple_vat_rates_in_breakdown(self):
        csv_content = "\n".join(
            [
                "Fecha;Proveedor;Base IVA 4%;Cuota IVA 4%;Base IVA 21%;Cuota IVA 21%;Retención IRPF;Total",
                "2026-09-10;Proveedor demo;100,00;4,00;200,00;42,00;15,00;331,00",
            ]
        ).encode("utf-8")

        record = parse_accounting_import_file(
            csv_content, "iva_mixto.csv", "purchases", "generic"
        )["records"][0]

        self.assertEqual(record["base_amount"], 300.00)
        self.assertEqual(record["vat_amount"], 46.00)
        self.assertEqual(record["vat_rate"], -1)
        self.assertEqual(record["total_amount"], 331.00)
        self.assertEqual(
            record["vat_breakdown"],
            [
                {"rate": 4, "base": 100.00, "vat_amount": 4.00, "total": 104.00},
                {"rate": 21, "base": 200.00, "vat_amount": 42.00, "total": 242.00},
            ],
        )
        self.assertEqual(record["errors"], [])

    def test_xlsx_finds_header_after_metadata_rows(self):
        from openpyxl import Workbook

        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Exportación de compras"])
        sheet.append(["Ejercicio", 2026])
        sheet.append([])
        sheet.append(["Fecha contable", "Acreedor", "Base", "Tipo IVA", "IVA", "Total"])
        sheet.append([date(2026, 9, 10), "Proveedor demo", 100.00, 21, 21.00, 121.00])
        buffer = BytesIO()
        workbook.save(buffer)

        record = parse_accounting_import_file(
            buffer.getvalue(), "odoo.xlsx", "purchases", "odoo"
        )["records"][0]

        self.assertEqual(record["invoice_date"], "2026-09-10")
        self.assertEqual(record["counterparty"], "Proveedor demo")
        self.assertEqual(record["total_amount"], 121.00)
        self.assertEqual(record["errors"], [])

    def test_preview_reports_invalid_and_duplicate_records(self):
        records = [
            {
                "row_number": 2,
                "invoice_date": "2026-09-10",
                "counterparty": "Proveedor demo",
                "concept": "",
                "base_amount": 100.0,
                "vat_rate": 21,
                "vat_amount": 21.0,
                "withholding_amount": 0.0,
                "total_amount": 121.0,
                "duplicate": False,
                "errors": [],
            },
            {
                "row_number": 3,
                "invoice_date": None,
                "counterparty": "",
                "concept": "",
                "base_amount": None,
                "vat_rate": None,
                "vat_amount": None,
                "withholding_amount": 0.0,
                "total_amount": None,
                "duplicate": False,
                "errors": ["Fecha inválida o ausente."],
            },
            {
                "row_number": 4,
                "invoice_date": "2026-09-10",
                "counterparty": "Proveedor demo",
                "concept": "",
                "base_amount": 100.0,
                "vat_rate": 21,
                "vat_amount": 21.0,
                "withholding_amount": 0.0,
                "total_amount": 121.0,
                "duplicate": True,
                "errors": [],
            },
        ]

        preview = serialize_accounting_import_preview(records, {"date": 0})

        self.assertEqual(preview["summary"], {"rows": 3, "ready": 1, "invalid": 1, "duplicates": 1})
        self.assertEqual(preview["issues"][0]["row"], 3)


if __name__ == "__main__":
    unittest.main()
