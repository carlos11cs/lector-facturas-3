from __future__ import annotations

import json
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Any, Callable, Dict, Iterable, Mapping, Optional

from .models import (
    AccountingAttachment,
    AccountingCompany,
    AccountingDocument,
    AccountingEntry,
    AccountingEntryLine,
    AccountingExportData,
    AccountingPeriod,
    AccountingProvenance,
    AccountingReferences,
    AccountingTotals,
    Counterparty,
    FiscalBreakdownLine,
    PaymentDue,
    TaxComponent,
)

CENT = Decimal("0.01")


def _decimal(value: Any, default: Optional[Decimal] = Decimal("0.00")) -> Optional[Decimal]:
    if value in (None, ""):
        return default
    try:
        return Decimal(str(value)).quantize(CENT, rounding=ROUND_HALF_UP)
    except (InvalidOperation, TypeError, ValueError):
        return default


def _json_list(value: Any) -> list:
    if not value:
        return []
    if isinstance(value, list):
        return value
    if isinstance(value, tuple):
        return list(value)
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
        except (TypeError, ValueError):
            return []
        return parsed if isinstance(parsed, list) else []
    return []


def _text(value: Any) -> Optional[str]:
    normalized = str(value).strip() if value is not None else ""
    return normalized or None


def _first_not_none(*values: Any) -> Any:
    for value in values:
        if value is not None:
            return value
    return None


def suggest_expense_account(
    *, expense_type=None, expense_family=None, expense_subtype=None, pnl_bucket=None
):
    normalized_type = (expense_type or "").strip().lower()
    normalized_family = (expense_family or "").strip().lower()
    normalized_subtype = (expense_subtype or "").strip().lower()
    normalized_bucket = (pnl_bucket or "").strip().lower()
    mapping = {
        "alquiler_local": ("621", "Arrendamientos y cánones"),
        "alquiler_cabina": ("621", "Arrendamientos y cánones"),
        "nomina": ("640", "Sueldos y salarios"),
        "seguridad_social": ("642", "Seguridad Social a cargo de la empresa"),
        "amortizacion": ("681", "Amortización del inmovilizado material"),
        "kilometraje": ("629", "Otros servicios"),
        "prestamo": ("662", "Intereses de deudas"),
    }
    if normalized_type in mapping:
        return mapping[normalized_type]
    if normalized_family == "rent":
        return ("621", "Arrendamientos y cánones")
    if normalized_family == "personnel":
        return ("640", "Gastos de personal")
    if normalized_family == "financing" or normalized_bucket == "financial_expense":
        return ("662", "Intereses de deudas")
    if normalized_subtype == "amortization" or normalized_bucket == "amortization_expense":
        return ("681", "Amortización del inmovilizado material")
    return ("629", "Otros servicios")


def _fiscal_breakdown(row: Mapping[str, Any]) -> tuple[FiscalBreakdownLine, ...]:
    raw_lines = _json_list(row.get("vat_breakdown"))
    if raw_lines:
        return tuple(
            FiscalBreakdownLine(
                tax_type="vat",
                rate=_decimal(line.get("rate"), None),
                base=_decimal(
                    _first_not_none(line.get("base"), line.get("base_amount")), None
                ),
                tax_amount=_decimal(
                    _first_not_none(
                        line.get("vat_amount"), line.get("tax_amount")
                    ),
                    None,
                ),
                total=_decimal(
                    _first_not_none(line.get("total"), line.get("total_amount")),
                    None,
                ),
            )
            for line in raw_lines
            if isinstance(line, dict)
        )
    base = _decimal(row.get("base_amount"), None)
    vat = _decimal(row.get("vat_amount"), None)
    if base is None and vat is None:
        return ()
    return (
        FiscalBreakdownLine(
            tax_type="vat",
            rate=_decimal(row.get("vat_rate"), None),
            base=base,
            tax_amount=vat,
            total=(base + vat).quantize(CENT) if base is not None and vat is not None else None,
        ),
    )


def _tax_components(value: Any, fallback_amount: Any = None) -> tuple[TaxComponent, ...]:
    components = []
    for item in _json_list(value):
        if not isinstance(item, dict):
            continue
        amount = _decimal(
            _first_not_none(item.get("tax_amount"), item.get("amount")), None
        )
        if amount is None:
            continue
        components.append(
            TaxComponent(
                tax_type=_text(item.get("tax_type") or item.get("type")),
                rate=_decimal(item.get("rate"), None),
                base=_decimal(item.get("base"), None),
                tax_amount=abs(amount),
            )
        )
    if components:
        return tuple(components)
    fallback = _decimal(fallback_amount, None)
    if fallback is None or fallback == 0:
        return ()
    return (TaxComponent(None, None, None, abs(fallback)),)


def _payment_schedule(row: Mapping[str, Any]) -> tuple[PaymentDue, ...]:
    completed_dates = [
        item
        for item in (_text(value) for value in _json_list(row.get("payment_completed_dates")))
        if item
    ]
    schedule = []
    for index, item in enumerate(_json_list(row.get("payment_schedule"))):
        if not isinstance(item, dict) or not _text(item.get("due_date")):
            continue
        schedule.append(
            PaymentDue(
                due_date=_text(item.get("due_date")) or "",
                amount=_decimal(item.get("amount"), None),
                actual_payment_date=(
                    _text(item.get("actual_payment_date"))
                    or (completed_dates[index] if index < len(completed_dates) else None)
                ),
            )
        )
    if schedule:
        return tuple(schedule)
    dates = [_text(item) for item in _json_list(row.get("payment_dates"))]
    dates = [item for item in dates if item]
    if not dates and _text(row.get("payment_date")):
        dates = [_text(row.get("payment_date"))]
    return tuple(
        PaymentDue(
            due_date=item,
            actual_payment_date=(
                completed_dates[index] if index < len(completed_dates) else None
            ),
        )
        for index, item in enumerate(dates)
    )


def _invoice_document(row: Mapping[str, Any], *, direction: str) -> AccountingDocument:
    is_purchase = direction == "purchase"
    counterparty_name = row.get("supplier") if is_purchase else row.get("client")
    withholding = abs(_decimal(row.get("withholding_amount")) or Decimal("0.00"))
    other_taxes = _tax_components(row.get("other_taxes"))
    other_total = sum((item.tax_amount for item in other_taxes), Decimal("0.00"))
    metadata = {
        "original_filename": row.get("original_filename"),
        "vat_rate": row.get("vat_rate"),
        "vat_deductible": row.get("vat_deductible"),
        "expense_category": row.get("expense_category"),
        "expense_family": row.get("expense_family"),
        "expense_subtype": row.get("expense_subtype"),
        "pnl_bucket": row.get("pnl_bucket"),
        "tax_model_targets": row.get("tax_model_targets"),
        "suggested_account": suggest_expense_account(
            expense_family=row.get("expense_family"),
            expense_subtype=row.get("expense_subtype"),
            pnl_bucket=row.get("pnl_bucket"),
        ),
    }
    return AccountingDocument(
        internal_id=int(row.get("id")),
        source_type="purchase_invoice" if is_purchase else "income_invoice",
        document_type="purchase_invoice" if is_purchase else "sales_invoice",
        direction=direction,
        issue_date=str(row.get("invoice_date") or ""),
        counterparty=Counterparty(
            name=_text(counterparty_name),
            tax_id=_text(row.get("counterparty_tax_id")),
            country=_text(row.get("counterparty_country")),
            address=_text(row.get("counterparty_address")),
        ),
        totals=AccountingTotals(
            taxable_base=_decimal(row.get("base_amount")) or Decimal("0.00"),
            vat_amount=_decimal(row.get("vat_amount")) or Decimal("0.00"),
            withholding_amount=withholding,
            other_taxes_amount=other_total,
            total_amount=_decimal(row.get("total_amount")) or Decimal("0.00"),
        ),
        fiscal_breakdown=_fiscal_breakdown(row),
        withholdings=_tax_components(row.get("withholding_details"), withholding),
        other_taxes=other_taxes,
        due_dates=_payment_schedule(row),
        invoice_number=_text(row.get("invoice_number")),
        series=_text(row.get("invoice_series")),
        is_rectificative=bool(row.get("is_rectificative")),
        currency=_text(row.get("currency")),
        concept=_text(row.get("concept")),
        references=AccountingReferences(
            document_reference=_text(row.get("document_reference")),
            order_reference=_text(row.get("order_reference")),
            delivery_note_reference=_text(row.get("delivery_note_reference")),
            rectified_invoice_reference=_text(row.get("rectified_invoice_reference")),
        ),
        provenance=AccountingProvenance(
            source_system=_text(row.get("source_system")),
            external_document_id=_text(row.get("external_document_id")),
            extraction_source=_text(row.get("extraction_source")),
            confidence_score=_decimal(row.get("confidence_score"), None),
        ),
        accounting_metadata=metadata,
    )


def _no_invoice_document(row: Mapping[str, Any]) -> AccountingDocument:
    base = _decimal(row.get("base_amount") or row.get("amount")) or Decimal("0.00")
    vat = _decimal(row.get("vat_amount")) or Decimal("0.00")
    withholding = abs(_decimal(row.get("withholding_amount")) or Decimal("0.00"))
    metadata = dict(row)
    metadata["suggested_account"] = suggest_expense_account(
        expense_type=row.get("expense_type"),
        expense_family=row.get("expense_family"),
        expense_subtype=row.get("expense_subtype"),
        pnl_bucket=row.get("pnl_bucket"),
    )
    return AccountingDocument(
        internal_id=int(row.get("id")),
        source_type="no_invoice_expense",
        document_type=_text(row.get("expense_type")) or "other_expense",
        direction="purchase",
        issue_date=str(row.get("expense_date") or ""),
        counterparty=Counterparty(
            name=_text(row.get("payroll_employee_name") or row.get("concept"))
        ),
        totals=AccountingTotals(base, vat, withholding, Decimal("0.00"), _decimal(row.get("amount")) or Decimal("0.00")),
        fiscal_breakdown=_fiscal_breakdown(row),
        due_dates=_payment_schedule(row),
        concept=_text(row.get("concept")),
        accounting_metadata=metadata,
    )


def _manual_sale_document(row: Mapping[str, Any]) -> AccountingDocument:
    issue_date = _text(row.get("invoice_date"))
    if not issue_date:
        issue_date = f"{int(row.get('anio')):04d}-{int(row.get('mes')):02d}-01"
    base = _decimal(row.get("base_facturada")) or Decimal("0.00")
    vat = _decimal(row.get("iva_repercutido")) or Decimal("0.00")
    total = _decimal(row.get("total_amount"), None) or (base + vat).quantize(CENT)
    return AccountingDocument(
        internal_id=int(row.get("id")),
        source_type="manual_billing",
        document_type="manual_sale",
        direction="sale",
        issue_date=issue_date,
        counterparty=Counterparty(name=None),
        totals=AccountingTotals(base, vat, Decimal("0.00"), Decimal("0.00"), total),
        concept=_text(row.get("concept")) or "Facturación manual",
        accounting_metadata=dict(row),
    )


def _entry(
    entry_id: str,
    entry_date: str,
    reference: str,
    source_type: str,
    source_id: int,
    counterparty: Optional[str],
    lines: Iterable[Mapping[str, Any]],
) -> AccountingEntry:
    normalized_lines = []
    for line in lines:
        debit = _decimal(line.get("debit") or line.get("debe")) or Decimal("0.00")
        credit = _decimal(line.get("credit") or line.get("haber")) or Decimal("0.00")
        if abs(debit) < Decimal("0.005") and abs(credit) < Decimal("0.005"):
            continue
        normalized_lines.append(
            AccountingEntryLine(
                line_number=len(normalized_lines) + 1,
                account=_text(line.get("account") or line.get("cuenta")),
                account_label=_text(line.get("account_label") or line.get("descripcion_cuenta")),
                debit=debit,
                credit=credit,
                counterparty=counterparty,
            )
        )
    return AccountingEntry(
        entry_id=entry_id,
        date=entry_date,
        reference=reference,
        source_document_id=f"{source_type}:{source_id}",
        source_type=source_type,
        source_id=source_id,
        journal="GENERAL",
        counterparty=counterparty,
        lines=tuple(normalized_lines),
    )


def _entries_for_documents(documents: Iterable[AccountingDocument]) -> list[AccountingEntry]:
    entries = []
    for document in documents:
        metadata = document.accounting_metadata
        base = document.totals.taxable_base
        vat = document.totals.vat_amount
        total = document.totals.total_amount
        withholding = document.totals.withholding_amount
        if document.source_type == "purchase_invoice":
            account, label = suggest_expense_account(
                expense_family=metadata.get("expense_family"),
                expense_subtype=metadata.get("expense_subtype"),
                pnl_bucket=metadata.get("pnl_bucket"),
            )
            lines = []
            if metadata.get("vat_deductible") and vat > 0:
                lines.extend((
                    {"account": account, "account_label": label, "debit": base},
                    {"account": "472", "account_label": "Hacienda Pública, IVA soportado", "debit": vat},
                ))
            else:
                lines.append({"account": account, "account_label": label, "debit": base + vat})
            if withholding > 0:
                lines.append({"account": "4751", "account_label": "Hacienda Pública acreedora por retenciones practicadas", "credit": withholding})
            if document.totals.other_taxes_amount > 0:
                lines.append({
                    "account": None,
                    "account_label": "Otros impuestos documentados",
                    "debit": document.totals.other_taxes_amount,
                })
            if total > 0:
                lines.append({"account": "410", "account_label": "Acreedores por prestaciones de servicios", "credit": total})
            entries.append(_entry(f"PUR-{document.internal_id}", document.issue_date, f"Factura proveedor {document.counterparty.name or metadata.get('original_filename')}", document.source_type, document.internal_id, document.counterparty.name, lines))
        elif document.source_type == "income_invoice":
            lines = [{"account": "430", "account_label": "Clientes", "debit": total}]
            if withholding:
                lines.append({"account": "473", "account_label": "Hacienda Pública, retenciones y pagos a cuenta", "debit": withholding})
            lines.extend((
                {"account": "700", "account_label": "Ventas de mercaderías / servicios", "credit": base},
                {"account": "477", "account_label": "Hacienda Pública, IVA repercutido", "credit": vat},
            ))
            if document.totals.other_taxes_amount > 0:
                lines.append({
                    "account": None,
                    "account_label": "Otros impuestos documentados",
                    "credit": document.totals.other_taxes_amount,
                })
            entries.append(_entry(f"SAL-{document.internal_id}", document.issue_date, f"Factura emitida {document.counterparty.name or metadata.get('original_filename')}", document.source_type, document.internal_id, document.counterparty.name, lines))
        elif document.source_type == "manual_billing":
            entries.append(_entry(f"MAN-{document.internal_id}", document.issue_date, document.concept or "Facturación manual", document.source_type, document.internal_id, None, (
                {"account": "430", "account_label": "Clientes", "debit": total},
                {"account": "700", "account_label": "Ventas de mercaderías / servicios", "credit": base},
                {"account": "477", "account_label": "Hacienda Pública, IVA repercutido", "credit": vat},
            )))
        elif document.source_type == "no_invoice_expense":
            expense_type = _text(metadata.get("expense_type")) or ""
            amount = _decimal(metadata.get("amount")) or Decimal("0.00")
            account, label = suggest_expense_account(
                expense_type=expense_type,
                expense_family=metadata.get("expense_family"),
                expense_subtype=metadata.get("expense_subtype"),
                pnl_bucket=metadata.get("pnl_bucket"),
            )
            lines = []
            if expense_type == "nomina":
                gross = _decimal(metadata.get("base_amount") or metadata.get("amount")) or Decimal("0.00")
                net = _decimal(metadata.get("payroll_net_amount")) or Decimal("0.00")
                deductions = _decimal(metadata.get("payroll_total_deductions_amount")) or Decimal("0.00")
                employer_cost = _decimal(metadata.get("payroll_employer_cost_amount")) or Decimal("0.00")
                employer_social_security = max(employer_cost - gross, Decimal("0.00"))
                lines.append({"account": "640", "account_label": "Sueldos y salarios", "debit": gross})
                if employer_social_security > 0:
                    lines.append({"account": "642", "account_label": "Seguridad Social a cargo de la empresa", "debit": employer_social_security})
                if net > 0:
                    lines.append({"account": "465", "account_label": "Remuneraciones pendientes de pago", "credit": net})
                if deductions > 0:
                    lines.append({"account": "476", "account_label": "Organismos de la Seguridad Social acreedores", "credit": deductions})
                balance = sum((_decimal(line.get("debit")) or Decimal("0.00") for line in lines), Decimal("0.00")) - sum((_decimal(line.get("credit")) or Decimal("0.00") for line in lines), Decimal("0.00"))
                if abs(balance) >= CENT:
                    lines.append({"account": "410", "account_label": "Acreedores varios", "credit": balance if balance > 0 else 0, "debit": abs(balance) if balance < 0 else 0})
            elif expense_type == "seguridad_social":
                lines = [
                    {"account": "642", "account_label": "Seguridad Social a cargo de la empresa", "debit": amount},
                    {"account": "476", "account_label": "Organismos de la Seguridad Social acreedores", "credit": amount},
                ]
            elif expense_type == "prestamo":
                interest = _decimal(metadata.get("interest_amount")) or Decimal("0.00")
                principal = max(amount - interest, Decimal("0.00"))
                lines = [
                    {"account": "520", "account_label": "Deudas a corto plazo con entidades de crédito", "debit": principal},
                    {"account": "662", "account_label": "Intereses de deudas", "debit": interest},
                    {"account": "572", "account_label": "Bancos e instituciones de crédito c/c vista", "credit": amount},
                ]
            else:
                if metadata.get("vat_deductible") and vat > 0:
                    lines.extend((
                        {"account": account, "account_label": label, "debit": base},
                        {"account": "472", "account_label": "Hacienda Pública, IVA soportado", "debit": vat},
                    ))
                else:
                    lines.append({"account": account, "account_label": label, "debit": amount})
                if withholding > 0:
                    lines.append({"account": "4751", "account_label": "Hacienda Pública acreedora por retenciones practicadas", "credit": withholding})
                creditor = amount - withholding
                if creditor > 0:
                    lines.append({"account": "410", "account_label": "Acreedores por prestaciones de servicios", "credit": creditor})
            entries.append(_entry(f"EXP-{document.internal_id}", document.issue_date, document.concept or expense_type or "Gasto", document.source_type, document.internal_id, document.counterparty.name, lines))
    return entries


def _loan_entries(rows: Iterable[Mapping[str, Any]]) -> list[AccountingEntry]:
    entries = []
    for row in rows:
        source_id = int(row.get("id"))
        total = _decimal(row.get("total_amount")) or Decimal("0.00")
        interest = _decimal(row.get("interest_amount")) or Decimal("0.00")
        principal = _decimal(row.get("principal_amount"), None) or max(total - interest, Decimal("0.00"))
        entries.append(_entry(f"LOA-{source_id}", str(row.get("payment_date") or ""), _text(row.get("concept")) or "Cuota de préstamo", "loan_installment", source_id, _text(row.get("bank_name")), (
            {"account": "520", "account_label": "Deudas a corto plazo con entidades de crédito", "debit": principal},
            {"account": "662", "account_label": "Intereses de deudas", "debit": interest},
            {"account": "572", "account_label": "Bancos e instituciones de crédito c/c vista", "credit": total},
        )))
    return entries


def _attachments(rows: Iterable[Mapping[str, Any]], resolver: Optional[Callable]) -> tuple[AccountingAttachment, ...]:
    attachments = []
    for row in rows:
        payload = resolver(row) if resolver else {}
        payload = payload if isinstance(payload, Mapping) else {}
        attachments.append(
            AccountingAttachment(
                internal_id=int(row.get("id")),
                filename=_text(row.get("original_filename")),
                document_type=_text(row.get("detected_document_type")),
                validation_status=_text(row.get("validation_status")),
                linked_record_type=_text(row.get("linked_accounting_record_type")),
                linked_record_id=row.get("linked_accounting_record_id"),
                reference_date=_text(payload.get("invoice_date") or payload.get("expense_date") or payload.get("payment_date") or row.get("registered_at")),
                counterparty=_text(payload.get("provider_name") or payload.get("counterparty_name") or payload.get("employee_name") or payload.get("concept")),
                total_amount=_decimal(payload.get("total_amount") or payload.get("amount") or payload.get("payable_amount")) or Decimal("0.00"),
                period=_text(row.get("period")),
                storage_path=_text(row.get("storage_path")),
                file_url=_text(row.get("file_url")),
            )
        )
    return tuple(attachments)


def build_accounting_export_data(
    source_data: Mapping[str, Any],
    *,
    company: Optional[Mapping[str, Any]] = None,
    period: Optional[Mapping[str, Any]] = None,
    document_data_resolver: Optional[Callable] = None,
) -> AccountingExportData:
    documents = [
        *(_invoice_document(row, direction="purchase") for row in source_data.get("purchase_invoices", ())),
        *(_invoice_document(row, direction="sale") for row in source_data.get("income_invoices", ())),
        *(_no_invoice_document(row) for row in source_data.get("no_invoice_expenses", ())),
        *(_manual_sale_document(row) for row in source_data.get("manual_sales", ())),
    ]
    entries = _entries_for_documents(documents)
    entries.extend(_loan_entries(source_data.get("loan_installments", ())))
    entries.sort(key=lambda item: (item.date, item.entry_id))
    company = company or {}
    period = period or {}
    return AccountingExportData(
        company=AccountingCompany(
            internal_id=company.get("id") or company.get("company_id"),
            display_name=_text(company.get("display_name") or company.get("company_name")),
            legal_name=_text(company.get("legal_name") or company.get("company_legal_name")),
            tax_id=_text(company.get("tax_id") or company.get("company_tax_id")),
        ),
        period=AccountingPeriod(
            start_date=_text(period.get("start_date")),
            end_date=_text(period.get("end_date")),
            label=_text(period.get("label") or period.get("period_label")),
        ),
        documents=tuple(documents),
        entries=tuple(entries),
        attachments=_attachments(source_data.get("documents", ()), document_data_resolver),
    )
