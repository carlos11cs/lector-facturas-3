from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any, Mapping, Optional, Tuple


@dataclass(frozen=True)
class AccountingCompany:
    internal_id: Optional[int]
    display_name: Optional[str] = None
    legal_name: Optional[str] = None
    tax_id: Optional[str] = None


@dataclass(frozen=True)
class AccountingPeriod:
    start_date: Optional[str]
    end_date: Optional[str]
    label: Optional[str] = None


@dataclass(frozen=True)
class Counterparty:
    name: Optional[str]
    tax_id: Optional[str] = None
    country: Optional[str] = None
    address: Optional[str] = None
    external_id: Optional[str] = None


@dataclass(frozen=True)
class FiscalBreakdownLine:
    tax_type: str
    rate: Optional[Decimal]
    base: Optional[Decimal]
    tax_amount: Optional[Decimal]
    total: Optional[Decimal]


@dataclass(frozen=True)
class TaxComponent:
    tax_type: Optional[str]
    rate: Optional[Decimal]
    base: Optional[Decimal]
    tax_amount: Decimal


@dataclass(frozen=True)
class PaymentDue:
    due_date: str
    amount: Optional[Decimal] = None
    actual_payment_date: Optional[str] = None


@dataclass(frozen=True)
class AccountingTotals:
    taxable_base: Decimal
    vat_amount: Decimal
    withholding_amount: Decimal
    other_taxes_amount: Decimal
    total_amount: Decimal


@dataclass(frozen=True)
class AccountingReferences:
    document_reference: Optional[str] = None
    order_reference: Optional[str] = None
    delivery_note_reference: Optional[str] = None
    rectified_invoice_reference: Optional[str] = None


@dataclass(frozen=True)
class AccountingProvenance:
    source_system: Optional[str] = None
    external_document_id: Optional[str] = None
    extraction_source: Optional[str] = None
    confidence_score: Optional[Decimal] = None


@dataclass(frozen=True)
class AccountingDocument:
    internal_id: int
    source_type: str
    document_type: str
    direction: str
    issue_date: str
    counterparty: Counterparty
    totals: AccountingTotals
    fiscal_breakdown: Tuple[FiscalBreakdownLine, ...] = ()
    withholdings: Tuple[TaxComponent, ...] = ()
    other_taxes: Tuple[TaxComponent, ...] = ()
    due_dates: Tuple[PaymentDue, ...] = ()
    invoice_number: Optional[str] = None
    series: Optional[str] = None
    is_rectificative: bool = False
    currency: Optional[str] = None
    concept: Optional[str] = None
    references: AccountingReferences = field(default_factory=AccountingReferences)
    provenance: AccountingProvenance = field(default_factory=AccountingProvenance)
    accounting_metadata: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class AccountingEntryLine:
    line_number: int
    account: Optional[str]
    account_label: Optional[str]
    debit: Decimal
    credit: Decimal
    counterparty: Optional[str] = None
    tax: Optional[Mapping[str, Any]] = None
    external_reference: Optional[str] = None


@dataclass(frozen=True)
class AccountingEntry:
    entry_id: str
    date: str
    reference: str
    source_document_id: str
    source_type: str
    source_id: int
    journal: Optional[str]
    counterparty: Optional[str]
    lines: Tuple[AccountingEntryLine, ...]


@dataclass(frozen=True)
class AccountingAttachment:
    internal_id: int
    filename: Optional[str]
    document_type: Optional[str]
    validation_status: Optional[str]
    linked_record_type: Optional[str]
    linked_record_id: Optional[int]
    reference_date: Optional[str]
    counterparty: Optional[str]
    total_amount: Decimal
    period: Optional[str]
    storage_path: Optional[str] = None
    file_url: Optional[str] = None


@dataclass(frozen=True)
class AccountingExportData:
    company: AccountingCompany
    period: AccountingPeriod
    documents: Tuple[AccountingDocument, ...]
    entries: Tuple[AccountingEntry, ...]
    attachments: Tuple[AccountingAttachment, ...]
