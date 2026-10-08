# Accounting Export Model

`invoices` and `income_invoices` remain Ledged's accounting-document source
of truth. `AccountingExportData` is an in-memory DTO built from those tables;
it is not a second persisted document store.

The Phase 1 nullable, additive columns on both invoice tables are:

- Identity: `invoice_number`, `invoice_series`, `is_rectificative`,
  `rectified_invoice_reference`.
- Counterparty: `counterparty_tax_id`, `counterparty_country`,
  `counterparty_address`.
- Document: `currency`, `concept`, `document_reference`, `order_reference`,
  `delivery_note_reference`.
- Provenance: `source_system`, `external_document_id`.
- Structured fiscal and payment details: `withholding_details`, `other_taxes`,
  `payment_schedule`.

All are nullable to preserve historical records without inventing values.
`vat_breakdown` remains the canonical multi-VAT detail. New calculation code
uses `Decimal`; existing aggregate columns remain compatibility summaries.

The sole current profile is `generic_ledged_v1`. It consumes the normalized
DTO for individual CSV/XLSX exports and the ZIP package, so those outputs use
the same rows and column definitions. Destination-specific profiles belong in
the registry only once their published import specifications are verified.
