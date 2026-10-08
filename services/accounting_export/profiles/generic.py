from __future__ import annotations

import csv
import io
import json
import re
import unicodedata
import zipfile
from decimal import Decimal
from typing import Callable, Mapping, Optional

import openpyxl

from ..models import AccountingDocument, AccountingExportData
from .base import ExportProfile, ExportTable


PURCHASE_COLUMNS = (
    "fecha", "documento_tipo", "origen_tipo", "origen_id", "contraparte",
    "concepto", "base", "iva", "retencion", "total", "iva_deducible",
    "cuenta_sugerida", "familia", "subtipo", "bucket_pyg", "modelos_fiscales",
)
SALES_COLUMNS = (
    "fecha", "documento_tipo", "origen_tipo", "origen_id", "cliente",
    "concepto", "base", "iva", "retencion", "total", "tipo_iva",
    "vencimiento", "estado_pago",
)
JOURNAL_COLUMNS = (
    "asiento_id", "linea", "fecha", "diario", "concepto", "cuenta",
    "descripcion_cuenta", "debe", "haber", "tercero", "documento_origen",
    "origen_tipo", "origen_id",
)
MANIFEST_COLUMNS = (
    "document_id", "archivo", "tipo_detectado", "estado",
    "referencia_contable_tipo", "referencia_contable_id", "fecha_referencia",
    "contraparte", "importe_total", "periodo",
)


def _number(value: Decimal) -> float:
    return float(value)


def _format_account(code, label):
    return f"{code or ''} {label or ''}".strip()


def _tax_targets(value):
    if not value:
        return []
    if isinstance(value, list):
        return value
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
        except (TypeError, ValueError):
            return [item.strip() for item in value.split(",") if item.strip()]
        return parsed if isinstance(parsed, list) else []
    return []


def _safe_attachment_name(value):
    ascii_value = unicodedata.normalize("NFKD", str(value or "")).encode(
        "ascii", "ignore"
    ).decode("ascii")
    normalized = re.sub(r"[^A-Za-z0-9_.-]+", "_", ascii_value).strip("._")
    return normalized or "documento"


class GenericLedgedExportProfile(ExportProfile):
    profile_id = "generic_ledged_v1"

    def purchases(self, export_data: AccountingExportData):
        rows = []
        for document in export_data.documents:
            if document.direction != "purchase":
                continue
            metadata = document.accounting_metadata
            if document.source_type == "purchase_invoice":
                account = metadata.get("suggested_account") or ("629", "Otros servicios")
                document_type = "Factura recibida"
                total = document.totals.total_amount
            elif document.source_type == "no_invoice_expense":
                account = metadata.get("suggested_account") or ("629", "Otros servicios")
                document_type = (metadata.get("expense_type") or "otro").replace("_", " ").title()
                total = document.totals.total_amount
            else:
                continue
            rows.append({
                "fecha": document.issue_date,
                "documento_tipo": document_type,
                "origen_tipo": document.source_type,
                "origen_id": document.internal_id,
                "contraparte": document.counterparty.name,
                "concepto": (
                    document.concept
                    if document.source_type == "no_invoice_expense"
                    else metadata.get("original_filename")
                ),
                "base": _number(document.totals.taxable_base),
                "iva": _number(document.totals.vat_amount),
                "retencion": _number(document.totals.withholding_amount),
                "total": _number(total),
                "iva_deducible": "Sí" if metadata.get("vat_deductible") else "No",
                "cuenta_sugerida": _format_account(*account),
                "familia": metadata.get("expense_family") or "",
                "subtipo": metadata.get("expense_subtype") or "",
                "bucket_pyg": metadata.get("pnl_bucket") or "",
                "modelos_fiscales": ", ".join(_tax_targets(metadata.get("tax_model_targets"))),
            })
        return sorted(rows, key=lambda item: (item.get("fecha") or "", str(item.get("origen_id") or "")))

    def sales(self, export_data: AccountingExportData):
        rows = []
        for document in export_data.documents:
            if document.direction != "sale":
                continue
            metadata = document.accounting_metadata
            if document.source_type == "income_invoice":
                due_date = document.due_dates[0].due_date if document.due_dates else ""
                rows.append({
                    "fecha": document.issue_date,
                    "documento_tipo": "Factura emitida",
                    "origen_tipo": document.source_type,
                    "origen_id": document.internal_id,
                    "cliente": document.counterparty.name,
                    "concepto": metadata.get("original_filename"),
                    "base": _number(document.totals.taxable_base),
                    "iva": _number(document.totals.vat_amount),
                    "retencion": _number(document.totals.withholding_amount),
                    "total": _number(document.totals.total_amount),
                    "tipo_iva": metadata.get("vat_rate") if metadata.get("vat_rate") is not None else "",
                    "vencimiento": due_date,
                    "estado_pago": "Planificado" if due_date else "",
                })
            elif document.source_type == "manual_billing":
                rows.append({
                    "fecha": document.issue_date,
                    "documento_tipo": "Registro manual",
                    "origen_tipo": document.source_type,
                    "origen_id": document.internal_id,
                    "cliente": "",
                    "concepto": document.concept or "Facturación manual",
                    "base": _number(document.totals.taxable_base),
                    "iva": _number(document.totals.vat_amount),
                    "retencion": "",
                    "total": _number(document.totals.total_amount),
                    "tipo_iva": metadata.get("tipo_iva") if metadata.get("tipo_iva") is not None else "",
                    "vencimiento": "",
                    "estado_pago": "",
                })
        return sorted(rows, key=lambda item: (item.get("fecha") or "", str(item.get("origen_id") or "")))

    def journal(self, export_data: AccountingExportData):
        rows = []
        for entry in export_data.entries:
            for line in entry.lines:
                rows.append({
                    "asiento_id": entry.entry_id,
                    "linea": line.line_number,
                    "fecha": entry.date,
                    "diario": entry.journal or "GENERAL",
                    "concepto": entry.reference,
                    "cuenta": line.account or "",
                    "descripcion_cuenta": line.account_label or "",
                    "debe": _number(line.debit),
                    "haber": _number(line.credit),
                    "tercero": entry.counterparty or "",
                    "documento_origen": entry.source_document_id,
                    "origen_tipo": entry.source_type,
                    "origen_id": entry.source_id,
                })
        return rows

    def manifest(self, export_data: AccountingExportData):
        return [{
            "document_id": item.internal_id,
            "archivo": item.filename,
            "tipo_detectado": item.document_type,
            "estado": item.validation_status,
            "referencia_contable_tipo": item.linked_record_type or "",
            "referencia_contable_id": item.linked_record_id or "",
            "fecha_referencia": item.reference_date,
            "contraparte": item.counterparty or "",
            "importe_total": _number(item.total_amount),
            "periodo": item.period or "",
        } for item in export_data.attachments]

    def table(self, export_data: AccountingExportData, kind: str) -> ExportTable:
        if kind == "purchases":
            return ExportTable(self.purchases(export_data), PURCHASE_COLUMNS, "Compras", "compras")
        if kind == "sales":
            return ExportTable(self.sales(export_data), SALES_COLUMNS, "Ventas", "ventas")
        if kind == "journal":
            return ExportTable(self.journal(export_data), JOURNAL_COLUMNS, "Asientos", "asientos")
        if kind == "manifest":
            return ExportTable(self.manifest(export_data), MANIFEST_COLUMNS, "Documentos", "manifest_documental")
        raise KeyError(kind)

    @staticmethod
    def csv_bytes(rows, columns):
        output = io.StringIO()
        writer = csv.writer(output, delimiter=";")
        writer.writerow(columns)
        for row in rows:
            values = []
            for column in columns:
                value = row.get(column, "")
                if isinstance(value, bool):
                    values.append("Sí" if value else "No")
                elif isinstance(value, (int, float, Decimal)):
                    values.append(f"{float(value or 0):.2f}".replace(".", ","))
                else:
                    values.append(str(value).strip() if value is not None else "")
            writer.writerow(values)
        return output.getvalue().encode("utf-8-sig")

    @staticmethod
    def xlsx_bytes(rows, columns, sheet_name):
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = (sheet_name or "Export")[:31]
        worksheet.append(list(columns))
        for row in rows:
            worksheet.append([row.get(column, "") for column in columns])
        for column_cells in worksheet.columns:
            max_length = max(len("" if cell.value is None else str(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(max_length + 2, 36)
        output = io.BytesIO()
        workbook.save(output)
        return output.getvalue()

    def package(
        self,
        export_data: AccountingExportData,
        metadata: Mapping,
        attachment_loader: Optional[Callable] = None,
    ) -> io.BytesIO:
        tables = {kind: self.table(export_data, kind) for kind in ("purchases", "sales", "journal", "manifest")}
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("compras.csv", self.csv_bytes(tables["purchases"].rows, tables["purchases"].columns))
            archive.writestr("ventas.csv", self.csv_bytes(tables["sales"].rows, tables["sales"].columns))
            archive.writestr("asientos.csv", self.csv_bytes(tables["journal"].rows, tables["journal"].columns))
            archive.writestr("manifest_documental.csv", self.csv_bytes(tables["manifest"].rows, tables["manifest"].columns))
            archive.writestr("manifest.json", json.dumps({
                **dict(metadata),
                "purchase_rows": len(tables["purchases"].rows),
                "sales_rows": len(tables["sales"].rows),
                "journal_rows": len(tables["journal"].rows),
                "document_rows": len(tables["manifest"].rows),
            }, ensure_ascii=False, indent=2))
            if attachment_loader:
                for attachment in export_data.attachments:
                    content = attachment_loader(attachment)
                    if not content:
                        continue
                    safe_name = _safe_attachment_name(
                        attachment.filename or f"documento_{attachment.internal_id}"
                    )
                    archive.writestr(f"documentos/{attachment.internal_id:05d}_{safe_name}", content)
        output.seek(0)
        return output
