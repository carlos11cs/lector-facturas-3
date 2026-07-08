import unittest

try:
    from app import (
        build_journal_export_rows,
        build_purchase_export_rows,
        build_sales_export_rows,
        parse_iso_date,
        suggest_expense_account,
    )
    APP_IMPORT_ERROR = None
except Exception as exc:  # pragma: no cover - test guard for limited local envs
    APP_IMPORT_ERROR = exc


@unittest.skipIf(APP_IMPORT_ERROR is not None, f"app import failed: {APP_IMPORT_ERROR}")
class TestAccountingExportHelpers(unittest.TestCase):
    def test_parse_iso_date_accepts_datetime_iso(self):
        parsed = parse_iso_date("2026-07-08T12:30:00")
        self.assertEqual(parsed.isoformat(), "2026-07-08")

    def test_suggest_expense_account_maps_payroll(self):
        self.assertEqual(
            suggest_expense_account(expense_type="nomina"),
            ("640", "Sueldos y salarios"),
        )

    def test_build_purchase_export_rows_includes_no_invoice_expense(self):
        rows = build_purchase_export_rows(
            {
                "purchase_invoices": [],
                "no_invoice_expenses": [
                    {
                        "id": 7,
                        "expense_date": "2026-07-01",
                        "expense_type": "alquiler_local",
                        "concept": "Alquiler julio",
                        "amount": 1210.0,
                        "base_amount": 1000.0,
                        "vat_amount": 210.0,
                        "withholding_amount": 190.0,
                        "vat_deductible": True,
                        "expense_family": "rent",
                        "expense_subtype": "local_rent",
                        "pnl_bucket": "operating_expense",
                        "tax_model_targets": '["115","180","303"]',
                        "payroll_employee_name": None,
                    }
                ],
            }
        )
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["cuenta_sugerida"], "621 Arrendamientos y cánones")
        self.assertEqual(rows[0]["retencion"], 190.0)

    def test_build_sales_export_rows_combines_manual_and_invoice(self):
        rows = build_sales_export_rows(
            {
                "income_invoices": [
                    {
                        "id": 1,
                        "invoice_date": "2026-07-02",
                        "client": "Cliente Demo",
                        "original_filename": "F-1.pdf",
                        "base_amount": 100.0,
                        "vat_amount": 21.0,
                        "total_amount": 121.0,
                        "vat_rate": 21,
                        "payment_date": None,
                    }
                ],
                "manual_sales": [
                    {
                        "id": 2,
                        "anio": 2026,
                        "mes": 7,
                        "invoice_date": None,
                        "concept": "Caja diaria",
                        "base_facturada": 50.0,
                        "iva_repercutido": 10.5,
                        "total_amount": 60.5,
                        "tipo_iva": 21,
                    }
                ],
            }
        )
        self.assertEqual(len(rows), 2)
        self.assertEqual({row["documento_tipo"] for row in rows}, {"Factura emitida", "Registro manual"})

    def test_build_journal_export_rows_for_payroll_balances_entry(self):
        rows = build_journal_export_rows(
            {
                "purchase_invoices": [],
                "income_invoices": [],
                "manual_sales": [],
                "loan_installments": [],
                "no_invoice_expenses": [
                    {
                        "id": 9,
                        "expense_date": "2026-07-05",
                        "expense_type": "nomina",
                        "concept": "Nómina Ana",
                        "amount": 2000.0,
                        "base_amount": 2000.0,
                        "payroll_net_amount": 1550.0,
                        "payroll_total_deductions_amount": 450.0,
                        "payroll_employer_cost_amount": 2650.0,
                        "expense_family": "personnel",
                        "expense_subtype": "payroll",
                        "pnl_bucket": "personnel_expense",
                        "payroll_employee_name": "Ana",
                    }
                ],
            }
        )
        debit_total = round(sum(row["debe"] for row in rows), 2)
        credit_total = round(sum(row["haber"] for row in rows), 2)
        self.assertEqual(debit_total, credit_total)
        self.assertTrue(any(row["cuenta"] == "642" for row in rows))


if __name__ == "__main__":
    unittest.main()
