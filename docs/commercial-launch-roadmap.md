# Ledged Commercial Launch Roadmap

## Objective
Prepare Ledged for commercial launch as a B2B product for Spanish accounting firms (`gestorias`) with a stable expense workflow, reliable fiscal outputs, and operational controls.

## Commercial Scope
Ledged must be able to:

- ingest supplier invoices with OCR + AI
- register payroll, Social Security, rent, financing and manual adjustments
- distinguish accounting date vs payment date
- calculate and expose tax-relevant data for:
  - `303`
  - `111`
  - `115`
  - `190`
  - `180`
  - `390`
  - `130`
  - `200`
  - `202`
- keep P&L, balance sheet and treasury views coherent

## Deployment Plan

### Deploy 1: Expense Taxonomy and Navigation
Goal: make the expense module understandable and usable for firms.

Deliverables:

- `Gastos` tabs:
  - `Facturas proveedor`
  - `Alquileres`
  - `Personal`
  - `Financiación`
  - `Otros ajustes`
- subtype selector constrained by active tab
- consistent wording across upload, manual entry and tables
- financing isolated from generic “other expenses”

Acceptance criteria:

- users cannot confuse payroll, financing and adjustments
- same expense cannot belong visually to two categories at once
- payment calendar remains coherent

### Deploy 2: Canonical Expense Model
Goal: stop mixing document nature, fiscal treatment and accounting behavior.

Target model:

- `expense_documents`
- `expense_entries`
- `expense_tax_lines`
- `fiscal_obligations`

Minimum fields to add or backfill:

- `expense_family`
- `expense_subtype`
- `counterparty_name`
- `counterparty_tax_id`
- `withholding_amount`
- `payment_dates`
- `analysis_status`
- `pnl_bucket`
- `tax_model_targets`

Acceptance criteria:

- each saved expense has one clear accounting family
- each tax impact is traceable from source document to model

### Deploy 3: Fiscal Engine Hardening
Goal: make reports and tax models reliable enough for real firms.

Rules to close:

- rent -> `115` / `180`
- payroll and professional retentions -> `111` / `190`
- VAT purchase/sales -> `303` / `390`
- financing -> interest to P&L, principal outside P&L
- non-deductible expenses excluded from taxable views

Acceptance criteria:

- every report can explain where each amount comes from
- no model depends on free-text interpretation alone

### Deploy 4: Operational Safety
Goal: make the platform sellable and supportable.

Deliverables:

- review states:
  - `ok`
  - `partial`
  - `review_required`
  - `low_quality_scan`
- explicit warning flows for doubtful OCR
- better audit trail for manual edits
- stable batch uploads with file-level error visibility
- deployment checklist for Render

Acceptance criteria:

- doubtful invoices never silently contaminate accounting
- operators can identify which file caused a failed batch

### Deploy 5: Commercial Readiness
Goal: ship with product and business safeguards in place.

Deliverables:

- pricing and plan gating
- Stripe integration
- onboarding flow for firms
- owner panel controls
- privacy/legal copy review
- production checklists:
  - backups
  - email deliverability
  - error monitoring
  - login/session checks

Acceptance criteria:

- a new firm can sign up, create a company and use the app without manual intervention
- billing, support and error visibility are operational

## Launch Gate
Ledged should only be marketed broadly when these are true:

1. `Gastos` taxonomy is stable.
2. Tax model calculations are deterministic and reviewable.
3. OCR/AI doubtful cases degrade safely to manual review.
4. Batch uploads identify the failing file.
5. P&L, balance, calendar and tax reports agree on the same stored data.
