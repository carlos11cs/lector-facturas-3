import csv
import io
import json
import unittest
import zipfile
from decimal import Decimal
from pathlib import Path

from sqlalchemy import create_engine, select
from sqlalchemy.pool import StaticPool

import app as app_module
from services.accounting_export import (
    build_accounting_export_data,
    export_profile_registry,
)
from services.accounting_export.profiles.generic import (
    PURCHASE_COLUMNS,
    SALES_COLUMNS,
)


def _read_csv(payload):
    return list(
        csv.DictReader(
            io.StringIO(payload.decode("utf-8-sig")), delimiter=";"
        )
    )


class TestAccountingExportData(unittest.TestCase):
    def test_preserves_multi_vat_and_optional_accounting_metadata(self):
        export_data = build_accounting_export_data(
            {
                "purchase_invoices": [
                    {
                        "id": 7,
                        "invoice_date": "2026-08-31",
                        "supplier": "Proveedor demo",
                        "invoice_number": "F-2026-7",
                        "invoice_series": "F",
                        "counterparty_tax_id": "B12345678",
                        "counterparty_country": "ES",
                        "currency": "EUR",
                        "base_amount": 150,
                        "vat_amount": 26,
                        "withholding_amount": 0,
                        "total_amount": 176,
                        "vat_deductible": True,
                        "vat_breakdown": json.dumps(
                            [
                                {"rate": 21, "base": 100, "vat_amount": 21, "total": 121},
                                {"rate": 10, "base": 50, "vat_amount": 5, "total": 55},
                            ]
                        ),
                        "payment_schedule": json.dumps(
                            [{"due_date": "2026-09-30", "amount": 176, "actual_payment_date": None}]
                        ),
                    }
                ]
            }
        )

        document = export_data.documents[0]
        self.assertEqual(document.invoice_number, "F-2026-7")
        self.assertEqual(document.series, "F")
        self.assertEqual(document.counterparty.tax_id, "B12345678")
        self.assertEqual(document.currency, "EUR")
        self.assertEqual(len(document.fiscal_breakdown), 2)
        self.assertEqual(document.fiscal_breakdown[0].base, Decimal("100.00"))
        self.assertEqual(document.fiscal_breakdown[1].rate, Decimal("10.00"))
        self.assertEqual(document.due_dates[0].amount, Decimal("176.00"))

    def test_historical_row_does_not_invent_missing_values(self):
        export_data = build_accounting_export_data(
            {
                "income_invoices": [
                    {
                        "id": 8,
                        "invoice_date": "2026-07-01",
                        "client": "Cliente histórico",
                        "base_amount": 100,
                        "vat_amount": 21,
                        "vat_rate": 21,
                        "total_amount": 121,
                    }
                ]
            }
        )

        document = export_data.documents[0]
        self.assertIsNone(document.invoice_number)
        self.assertIsNone(document.series)
        self.assertIsNone(document.counterparty.tax_id)
        self.assertIsNone(document.counterparty.country)
        self.assertIsNone(document.currency)
        self.assertIsNone(document.provenance.source_system)

    def test_generic_profile_is_the_only_registered_profile_and_dto_is_vendor_neutral(self):
        profile = export_profile_registry.get("generic_ledged_v1")
        self.assertEqual(profile.profile_id, "generic_ledged_v1")
        with self.assertRaises(KeyError):
            export_profile_registry.get("contasol")

        model_source = Path("services/accounting_export/models.py").read_text(
            encoding="utf-8"
        ).lower()
        for vendor in ("contasol", "a3", "odoo", "cegid"):
            self.assertNotIn(vendor, model_source)

    def test_individual_and_zip_exports_share_rows_columns_and_sales_withholding(self):
        export_data = build_accounting_export_data(
            {
                "purchase_invoices": [
                    {
                        "id": 1,
                        "invoice_date": "2026-07-01",
                        "supplier": "Proveedor",
                        "original_filename": "compra.pdf",
                        "base_amount": 100,
                        "vat_amount": 21,
                        "withholding_amount": 0,
                        "total_amount": 121,
                        "vat_deductible": True,
                    }
                ],
                "income_invoices": [
                    {
                        "id": 2,
                        "invoice_date": "2026-07-02",
                        "client": "Cliente",
                        "original_filename": "venta.pdf",
                        "base_amount": 100,
                        "vat_amount": 21,
                        "withholding_amount": 15,
                        "total_amount": 106,
                        "vat_rate": 21,
                    }
                ],
            }
        )
        profile = export_profile_registry.get("generic_ledged_v1")
        purchase_table = profile.table(export_data, "purchases")
        sales_table = profile.table(export_data, "sales")
        package = profile.package(export_data, {})

        with zipfile.ZipFile(package) as archive:
            zip_purchases = archive.read("compras.csv")
            zip_sales = archive.read("ventas.csv")

        self.assertEqual(
            zip_purchases,
            profile.csv_bytes(purchase_table.rows, purchase_table.columns),
        )
        self.assertEqual(
            zip_sales,
            profile.csv_bytes(sales_table.rows, sales_table.columns),
        )
        self.assertEqual(tuple(purchase_table.columns), PURCHASE_COLUMNS)
        self.assertEqual(tuple(sales_table.columns), SALES_COLUMNS)
        self.assertEqual(_read_csv(zip_sales)[0]["retencion"], "15,00")


class TestAccountingPersistenceFlows(unittest.TestCase):
    def setUp(self):
        self.original_engine = app_module.engine
        self.engine = create_engine(
            "sqlite://",
            future=True,
            connect_args={"check_same_thread": False},
            poolclass=StaticPool,
        )
        app_module.metadata.create_all(self.engine)
        app_module.engine = self.engine
        with self.engine.begin() as conn:
            conn.execute(
                app_module.users_table.insert().values(
                    id=1,
                    email="owner@test.local",
                    password_hash="test",
                    role="owner",
                    plan="internal",
                    created_at="2026-01-01T00:00:00",
                    is_active=True,
                )
            )
            conn.execute(
                app_module.companies_table.insert().values(
                    id=1,
                    user_id=1,
                    display_name="Empresa prueba",
                    legal_name="Empresa Prueba SL",
                    tax_id="B00000000",
                    company_type="company",
                    created_at="2026-01-01T00:00:00",
                )
            )
        self.client = app_module.app.test_client()
        with self.client.session_transaction() as session:
            session["user_id"] = 1

    def tearDown(self):
        app_module.engine = self.original_engine
        self.engine.dispose()

    def test_ai_save_retrieve_and_export_data_preserve_metadata(self):
        response = self.client.post(
            "/api/upload",
            json={
                "companyId": 1,
                "entries": [
                    {
                        "originalFilename": "factura.pdf",
                        "date": "2026-08-31",
                        "paymentDates": ["2026-09-30"],
                        "supplier": "Proveedor IA",
                        "base": 150,
                        "vat": 21,
                        "vatAmount": 26,
                        "total": 178,
                        "vatBreakdown": [
                            {"rate": 21, "base": 100, "vat_amount": 21, "total": 121},
                            {"rate": 10, "base": 50, "vat_amount": 5, "total": 55},
                        ],
                        "invoiceNumber": "F-2026-99",
                        "invoiceSeries": "F",
                        "counterpartyTaxId": "B12345678",
                        "counterpartyCountry": "ES",
                        "counterpartyAddress": "Calle prueba 1",
                        "currency": "EUR",
                        "concept": "Servicios",
                        "orderReference": "P-7",
                        "sourceSystem": "invoice_ai",
                        "externalDocumentId": "ext-99",
                        "otherTaxes": [
                            {"tax_type": "surcharge", "tax_amount": 2}
                        ],
                        "paymentSchedule": [
                            {"due_date": "2026-09-30", "amount": 178}
                        ],
                        "extractionSource": "llm",
                        "confidenceScore": 0.85,
                        "analysisStatus": "ok",
                    }
                ],
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())

        listed = self.client.get(
            "/api/invoices?company_id=1&month=8&year=2026"
        ).get_json()["invoices"][0]
        self.assertEqual(listed["invoice_number"], "F-2026-99")
        self.assertEqual(listed["counterparty_tax_id"], "B12345678")
        self.assertEqual(listed["currency"], "EUR")
        self.assertEqual(listed["payment_schedule"][0]["amount"], 178.0)
        self.assertEqual(listed["extraction_source"], "llm")
        self.assertEqual(listed["confidence_score"], 0.85)

        with self.engine.connect() as conn:
            row = conn.execute(
                select(app_module.invoices_table)
            ).mappings().one()
        export_data = build_accounting_export_data(
            {"purchase_invoices": [row]}
        )
        document = export_data.documents[0]
        self.assertEqual(document.invoice_number, "F-2026-99")
        self.assertEqual(len(document.fiscal_breakdown), 2)
        self.assertEqual(document.totals.other_taxes_amount, Decimal("2.00"))
        self.assertEqual(document.due_dates[0].amount, Decimal("178.00"))

    def test_csv_import_persists_purchase_and_sales_document_numbers(self):
        for record_type, party_header, party, number in (
            ("purchases", "Proveedor", "Proveedor CSV", "C-100"),
            ("sales", "Cliente", "Cliente CSV", "V-200"),
        ):
            content = (
                f"Fecha;{party_header};Concepto;Número factura;NIF;Moneda;Base;Tipo IVA;IVA;Total;Rectificativa\n"
                f"2026-09-10;{party};Servicio;{number};B12345678;EUR;100,00;21;21,00;121,00;Sí\n"
            ).encode("utf-8")
            response = self.client.post(
                "/api/accounting-integrations/import",
                data={
                    "record_type": record_type,
                    "source": "a3",
                    "file": (io.BytesIO(content), f"{record_type}.csv"),
                },
                content_type="multipart/form-data",
            )
            self.assertEqual(response.status_code, 200, response.get_json())

        with self.engine.connect() as conn:
            purchase = conn.execute(
                select(app_module.invoices_table)
            ).mappings().one()
            sale = conn.execute(
                select(app_module.income_invoices_table)
            ).mappings().one()
        self.assertEqual(purchase["invoice_number"], "C-100")
        self.assertEqual(sale["invoice_number"], "V-200")
        self.assertEqual(purchase["concept"], "Servicio")
        self.assertEqual(sale["concept"], "Servicio")
        self.assertEqual(purchase["counterparty_tax_id"], "B12345678")
        self.assertEqual(sale["counterparty_tax_id"], "B12345678")
        self.assertEqual(purchase["currency"], "EUR")
        self.assertEqual(sale["currency"], "EUR")
        self.assertEqual(purchase["source_system"], "a3")
        self.assertEqual(sale["source_system"], "a3")
        self.assertTrue(purchase["is_rectificative"])
        self.assertTrue(sale["is_rectificative"])

    def test_income_save_preserves_metadata_and_absent_fields_remain_null(self):
        response = self.client.post(
            "/api/income-invoices",
            json={
                "companyId": 1,
                "entries": [
                    {
                        "originalFilename": "emitida.pdf",
                        "date": "2026-08-15",
                        "client": "Cliente IA",
                        "base": 100,
                        "vat": 21,
                        "vatAmount": 21,
                        "total": 121,
                        "invoiceNumber": "V-2026-5",
                        "counterpartyTaxId": "B87654321",
                        "currency": "EUR",
                        "extractionSource": "llm",
                        "confidenceScore": 0.9,
                        "analysisStatus": "ok",
                    }
                ],
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())

        with self.engine.connect() as conn:
            row = conn.execute(
                select(app_module.income_invoices_table)
            ).mappings().one()
        self.assertEqual(row["invoice_number"], "V-2026-5")
        self.assertEqual(row["counterparty_tax_id"], "B87654321")
        self.assertEqual(row["currency"], "EUR")
        self.assertEqual(row["extraction_source"], "llm")
        self.assertEqual(row["confidence_score"], 0.9)
        self.assertIsNone(row["invoice_series"])
        self.assertIsNone(row["counterparty_country"])
        self.assertIsNone(row["external_document_id"])

    def test_new_columns_are_nullable_and_no_duplicate_document_table_exists(self):
        self.assertNotIn("accounting_documents", app_module.metadata.tables)
        for table in (app_module.invoices_table, app_module.income_invoices_table):
            for column in (
                "invoice_number",
                "counterparty_tax_id",
                "currency",
                "other_taxes",
                "payment_schedule",
            ):
                self.assertTrue(table.c[column].nullable)


if __name__ == "__main__":
    unittest.main()
