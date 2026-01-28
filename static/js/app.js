const pendingFiles = [];
let lineChart;
let billingLineChart;
let pieChart;
let netChart;
let expenseVatTotal = 0;
let billingVatTotal = 0;
let currentInvoices = [];
let currentNoInvoiceExpenses = [];
let billingBaseTotal = 0;
let currentSummary = null;
let currentBillingSummary = null;
let currentBillingEntries = [];
let currentDeductibleExpenses = 0;
let annualBillingBaseTotal = 0;
let annualDeductibleExpenses = 0;
let currentPayments = null;
let selectedPaymentDay = null;
let calendarMonth = null;
let calendarYear = null;
let calendarOverride = false;
let companies = [];
let selectedCompanyId = null;
let pendingIncomeFiles = [];
let currentIncomeInvoices = [];
let staffMembers = [];
const lowQualityDismissedIds = new Set();
let billingLastSource = "base";

const monthNames = [
  "Enero",
  "Febrero",
  "Marzo",
  "Abril",
  "Mayo",
  "Junio",
  "Julio",
  "Agosto",
  "Septiembre",
  "Octubre",
  "Noviembre",
  "Diciembre",
];

const expenseCategoryLabels = {
  with_invoice: "Gasto con factura",
  without_invoice: "Gasto sin factura",
  non_deductible: "No deducible",
};

const noInvoiceTypeLabels = {
  nomina: "Nómina",
  seguridad_social: "Seguridad Social",
  amortizacion: "Amortización",
  kilometraje: "Kilometraje",
  otro: "Otro",
};

const ANALYSIS_ERROR_MESSAGE =
  "No se ha podido analizar la factura automáticamente. Puedes introducir los datos manualmente.";
const LOW_QUALITY_SCAN_MESSAGE =
  "La calidad de la factura escaneada no es óptima. No se puede leer correctamente. Por favor, introduce los datos manualmente.";

function showLowQualityModal() {
  const modal = document.getElementById("lowQualityModal");
  if (!modal) {
    return;
  }
  modal.classList.add("is-visible");
  modal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
}

function hideLowQualityModal() {
  const modal = document.getElementById("lowQualityModal");
  if (!modal) {
    return;
  }
  modal.classList.remove("is-visible");
  modal.setAttribute("aria-hidden", "true");
  document.body.classList.remove("modal-open");
}

function formatCurrency(value) {
  const number = Number(value || 0);
  return `${number.toFixed(2).replace(".", ",")} €`;
}

function parseNumberInput(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const cleaned = String(value).replace(",", ".").trim();
  const numeric = Number(cleaned);
  return Number.isNaN(numeric) ? null : numeric;
}

function roundAmount(value) {
  return Math.round(Number(value) * 100) / 100;
}

function formatAmountInput(value) {
  if (value === null || value === undefined || Number.isNaN(Number(value))) {
    return "";
  }
  return Number(value).toFixed(2);
}

function getPnlInputValue(id) {
  const el = document.getElementById(id);
  if (!el) {
    return 0;
  }
  const parsed = parseNumberInput(el.value);
  return parsed === null ? 0 : parsed;
}

function setPnlInputValue(id, value, auto = false) {
  const el = document.getElementById(id);
  if (!el) {
    return;
  }
  if (!auto || !pnlManualOverrides.has(id)) {
    el.value = formatAmountInput(value);
  }
}

function bindPnlInputs() {
  pnlInputIds.forEach((id) => {
    const el = document.getElementById(id);
    if (!el) {
      return;
    }
    el.addEventListener("input", () => {
      pnlManualOverrides.add(id);
      updatePnlSummary();
    });
  });
}

function normalizeEntityName(value) {
  if (!value) {
    return "";
  }
  return String(value).toLowerCase().replace(/[^a-z0-9]/g, "");
}

function getActiveCompanyNames() {
  const company = getSelectedCompany();
  if (!company) {
    return [];
  }
  return [company.display_name, company.legal_name].filter(Boolean);
}

function isSupplierSameAsCompany(supplier) {
  const normalizedSupplier = normalizeEntityName(supplier);
  if (!normalizedSupplier) {
    return false;
  }
  return getActiveCompanyNames().some(
    (name) => normalizeEntityName(name) === normalizedSupplier
  );
}

function calculateVatFields({ baseValue, totalValue, vatRateValue, source }) {
  const vatRate = parseNumberInput(vatRateValue);
  if (vatRate === null) {
    return {
      base: baseValue,
      vatAmount: null,
      total: totalValue,
    };
  }
  const factor = 1 + vatRate / 100;

  if (source === "total" && totalValue !== null) {
    const base = roundAmount(totalValue / factor);
    return {
      base,
      vatAmount: roundAmount(totalValue - base),
      total: roundAmount(totalValue),
    };
  }

  if (baseValue !== null) {
    const vatAmount = roundAmount(baseValue * (vatRate / 100));
    return {
      base: baseValue,
      vatAmount,
      total: roundAmount(baseValue + vatAmount),
    };
  }

  if (totalValue !== null) {
    const base = roundAmount(totalValue / factor);
    return {
      base,
      vatAmount: roundAmount(totalValue - base),
      total: roundAmount(totalValue),
    };
  }

  return { base: null, vatAmount: null, total: null };
}

function applyVatCalculation(item, inputs, source) {
  const baseValue = parseNumberInput(inputs.base.value);
  const totalValue = parseNumberInput(inputs.total.value);
  const vatRateValue = resolveVatRateValue(inputs.vat.value);
  const result = calculateVatFields({
    baseValue,
    totalValue,
    vatRateValue,
    source,
  });

  if (source === "total" && result.base !== null) {
    inputs.base.value = formatAmountInput(result.base);
  } else if (source === "total" && result.base === null) {
    inputs.base.value = "";
  }
  if (result.vatAmount !== null) {
    inputs.vatAmount.value = formatAmountInput(result.vatAmount);
  } else {
    inputs.vatAmount.value = "";
  }
  if (result.total !== null) {
    inputs.total.value = formatAmountInput(result.total);
  } else {
    inputs.total.value = "";
  }

  item.base = inputs.base.value;
  item.vat = resolveVatRateValue(inputs.vat.value);
  item.vatAmount = inputs.vatAmount.value;
  item.total = inputs.total.value;
}

function syncBillingCalculation(source) {
  if (!billingBaseInput || !billingTotalInput || !billingVatSelect || !billingVatAmountInput) {
    return;
  }
  const baseValue = parseNumberInput(billingBaseInput.value);
  const totalValue = parseNumberInput(billingTotalInput.value);
  const vatRateValue = resolveVatRateValue(billingVatSelect.value);
  const result = calculateVatFields({
    baseValue,
    totalValue,
    vatRateValue,
    source,
  });

  if (result.base !== null) {
    billingBaseInput.value = formatAmountInput(result.base);
  } else if (source === "total") {
    billingBaseInput.value = "";
  }
  if (result.total !== null) {
    billingTotalInput.value = formatAmountInput(result.total);
  } else if (source === "base") {
    billingTotalInput.value = "";
  }
  if (result.vatAmount !== null) {
    billingVatAmountInput.value = formatAmountInput(result.vatAmount);
  } else {
    billingVatAmountInput.value = "";
  }
}

function normalizeInvoiceAmounts(item) {
  const baseValue = parseNumberInput(item.base);
  const totalValue = parseNumberInput(item.total);
  const vatRateValue = resolveVatRateValue(item.vat);
  const source = baseValue !== null ? "base" : "total";
  const result = calculateVatFields({
    baseValue,
    totalValue,
    vatRateValue,
    source,
  });

  return {
    base: result.base !== null ? formatAmountInput(result.base) : "",
    vatAmount: result.vatAmount !== null ? formatAmountInput(result.vatAmount) : "",
    total: result.total !== null ? formatAmountInput(result.total) : "",
  };
}

function formatMonthYear(month, year) {
  const name = monthNames[month - 1] || "";
  return `${name} ${year}`;
}

function computePaymentDate(invoiceDate, paymentDate) {
  if (paymentDate) {
    return paymentDate;
  }
  if (!invoiceDate) {
    return "";
  }
  const base = new Date(`${invoiceDate}T00:00:00`);
  if (Number.isNaN(base.getTime())) {
    return "";
  }
  base.setDate(base.getDate() + 30);
  return base.toISOString().slice(0, 10);
}

function getCalendarMonthYear() {
  if (calendarMonth && calendarYear) {
    return { month: calendarMonth, year: calendarYear };
  }
  return getSelectedMonthYear();
}

function setCalendarMonthYear(month, year, override = true) {
  calendarMonth = month;
  calendarYear = year;
  calendarOverride = override;
}

function syncCalendarWithFilters() {
  if (calendarOverride) {
    return;
  }
  const { month, year } = getSelectedMonthYear();
  if (month && year) {
    setCalendarMonthYear(month, year, false);
  }
}

function shiftCalendarMonth(delta) {
  const { month, year } = getCalendarMonthYear();
  if (!month || !year) {
    return;
  }
  let newMonth = month + delta;
  let newYear = year;
  if (newMonth < 1) {
    newMonth = 12;
    newYear -= 1;
  }
  if (newMonth > 12) {
    newMonth = 1;
    newYear += 1;
  }
  setCalendarMonthYear(newMonth, newYear, true);
  selectedPaymentDay = null;
  refreshPayments();
}

function withCompanyParam(url) {
  const companyId = getSelectedCompanyId();
  if (!companyId) {
    return url;
  }
  return `${url}${url.includes("?") ? "&" : "?"}company_id=${companyId}`;
}

const allowedExtensions = new Set([".pdf", ".jpg", ".jpeg", ".png"]);

const monthSelect = document.getElementById("monthSelect");
const yearSelect = document.getElementById("yearSelect");
const periodSelect = document.getElementById("periodSelect");
const companySelect = document.getElementById("companySelect");
const supplierSuggestions = document.getElementById("supplierSuggestions");
const billingMonthSelect = document.getElementById("billingMonthSelect");
const billingYearSelect = document.getElementById("billingYearSelect");
const billingDateInput = document.getElementById("billingDateInput");
const billingConceptInput = document.getElementById("billingConceptInput");
const billingBaseInput = document.getElementById("billingBaseInput");
const billingVatSelect = document.getElementById("billingVatSelect");
const billingVatAmountInput = document.getElementById("billingVatAmountInput");
const billingTotalInput = document.getElementById("billingTotalInput");
const billingSaveBtn = document.getElementById("billingSaveBtn");
const billingEntriesBody = document.querySelector("#billingEntriesTable tbody");
const billingEntriesEmpty = document.getElementById("billingEntriesEmpty");
const invoicesTableBody = document.querySelector("#invoicesTable tbody");
const invoicesEmpty = document.getElementById("invoicesEmpty");
const taxPeriodBadge = document.getElementById("taxPeriodBadge");
const noInvoiceDate = document.getElementById("noInvoiceDate");
const noInvoiceConcept = document.getElementById("noInvoiceConcept");
const noInvoiceAmount = document.getElementById("noInvoiceAmount");
const noInvoiceType = document.getElementById("noInvoiceType");
const noInvoiceDeductible = document.getElementById("noInvoiceDeductible");
const noInvoiceSaveBtn = document.getElementById("noInvoiceSaveBtn");
const noInvoiceTableBody = document.querySelector("#noInvoiceTable tbody");
const noInvoiceEmpty = document.getElementById("noInvoiceEmpty");
const fileInput = document.getElementById("fileInput");
const folderInput = document.getElementById("folderInput");
const dropZone = document.getElementById("dropZone");
const uploadTableBody = document.querySelector("#uploadTable tbody");
const emptyMessage = document.getElementById("emptyMessage");
const uploadBtn = document.getElementById("uploadBtn");
const navLinks = document.querySelectorAll(".nav-link");
const sections = document.querySelectorAll(".page-section");
const sidebarToggle = document.getElementById("sidebarToggle");
const sidebarOverlay = document.getElementById("sidebarOverlay");
const exportPnlBtn = document.getElementById("exportPnlBtn");
const pnlName = document.getElementById("pnlName");
const pnlTaxId = document.getElementById("pnlTaxId");
const globalProcessing = document.getElementById("globalProcessing");
const globalProcessingText = document.getElementById("globalProcessingText");
const companyDisplayName = document.getElementById("companyDisplayName");
const companyLegalName = document.getElementById("companyLegalName");
const companyTaxId = document.getElementById("companyTaxId");
const companyType = document.getElementById("companyType");
const companyAssignedSelect = document.getElementById("companyAssignedSelect");
const companyEmail = document.getElementById("companyEmail");
const companyPhone = document.getElementById("companyPhone");
const companySaveBtn = document.getElementById("companySaveBtn");
const companiesTableBody = document.querySelector("#companiesTable tbody");
const companiesEmpty = document.getElementById("companiesEmpty");
const staffEmail = document.getElementById("staffEmail");
const staffPassword = document.getElementById("staffPassword");
const staffSaveBtn = document.getElementById("staffSaveBtn");
const staffTableBody = document.querySelector("#staffTable tbody");
const staffEmpty = document.getElementById("staffEmpty");
const headerCompanyLabel = document.getElementById("headerCompanyLabel");
const headerPeriodLabel = document.getElementById("headerPeriodLabel");
const headerUserEmail = document.getElementById("headerUserEmail");
const incomeUploadBtn = document.getElementById("incomeUploadBtn");
const incomeDropZone = document.getElementById("incomeDropZone");
const incomeFileInput = document.getElementById("incomeFileInput");
const incomeUploadTableBody = document.querySelector("#incomeUploadTable tbody");
const incomeEmptyMessage = document.getElementById("incomeEmptyMessage");
const incomeInvoicesTableBody = document.querySelector("#incomeInvoicesTable tbody");
const incomeInvoicesEmpty = document.getElementById("incomeInvoicesEmpty");
const reportYearSelect = document.getElementById("reportYearSelect");
const reportQuarterSelect = document.getElementById("reportQuarterSelect");
const reportStartMonthSelect = document.getElementById("reportStartMonthSelect");
const reportEndMonthSelect = document.getElementById("reportEndMonthSelect");
const reportDownloadBtn = document.getElementById("reportDownloadBtn");
const reportEmailBtn = document.getElementById("reportEmailBtn");
const reportStatus = document.getElementById("reportStatus");
const currentUserRole = document.body ? document.body.dataset.userRole : null;
const paymentCalendar = document.getElementById("paymentCalendar");
const paymentCalendarTitle = document.getElementById("paymentCalendarTitle");
const paymentPrevMonth = document.getElementById("paymentPrevMonth");
const paymentNextMonth = document.getElementById("paymentNextMonth");
const paymentMonthTotal = document.getElementById("paymentMonthTotal");
const paymentMonthEmpty = document.getElementById("paymentMonthEmpty");
const paymentDayTitle = document.getElementById("paymentDayTitle");
const paymentDayList = document.getElementById("paymentDayList");
const paymentDayTotal = document.getElementById("paymentDayTotal");
const pnlInputIds = [
  "pnlLine1",
  "pnlLine2",
  "pnlLine3",
  "pnlLine4",
  "pnlLine5",
  "pnlLine6",
  "pnlLine7",
  "pnlLine8",
  "pnlLine9",
  "pnlLine10",
  "pnlLine11",
  "pnlLine12",
  "pnlLine13a",
  "pnlLine13b",
  "pnlLine14",
  "pnlLine15",
  "pnlLine16",
  "pnlLine17",
  "pnlLine18a",
  "pnlLine18b",
  "pnlLine18c",
  "pnlLine19",
];
const pnlManualOverrides = new Set();

function isAllowedFile(fileName) {
  const lower = fileName.toLowerCase();
  const dotIndex = lower.lastIndexOf(".");
  if (dotIndex === -1) {
    return false;
  }
  return allowedExtensions.has(lower.slice(dotIndex));
}

function populateMonthSelects() {
  if (!monthSelect || !billingMonthSelect) {
    return;
  }
  monthSelect.innerHTML = "";
  billingMonthSelect.innerHTML = "";
  monthNames.forEach((name, index) => {
    const option = document.createElement("option");
    option.value = String(index + 1);
    option.textContent = name;
    monthSelect.appendChild(option);
    const billingOption = option.cloneNode(true);
    billingMonthSelect.appendChild(billingOption);
  });
}

function createVatSelect(selected) {
  const select = document.createElement("select");
  ["0", "4", "10", "21"].forEach((rate) => {
    const option = document.createElement("option");
    option.value = rate;
    option.textContent = `${rate}%`;
    select.appendChild(option);
  });
  applyVatSelection(select, selected);
  return select;
}

function normalizeVatRateValue(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const cleaned = String(value).replace("%", "").trim();
  const numeric = Number(cleaned.replace(",", "."));
  if (Number.isNaN(numeric) || numeric < 0) {
    return null;
  }
  const roundedInt = Math.round(numeric);
  if (Math.abs(numeric - roundedInt) < 0.001) {
    return String(roundedInt);
  }
  return String(Number(numeric.toFixed(2)));
}

function resolveVatRateValue(value, fallback = "21") {
  const normalized = normalizeVatRateValue(value);
  return normalized === null ? fallback : normalized;
}

function applyVatSelection(select, value, fallback = "21") {
  const normalized = resolveVatRateValue(value, fallback);
  const exists = [...select.options].some((option) => option.value === normalized);
  if (!exists) {
    const option = document.createElement("option");
    option.value = normalized;
    option.textContent = `${normalized}%`;
    select.appendChild(option);
  }
  select.value = normalized;
}

function parseVatBreakdown(value) {
  if (!value) {
    return [];
  }
  if (Array.isArray(value)) {
    return value;
  }
  if (typeof value === "string") {
    try {
      const parsed = JSON.parse(value);
      return Array.isArray(parsed) ? parsed : [];
    } catch (err) {
      return [];
    }
  }
  return [];
}

function normalizeBreakdownLine(line) {
  const rateValue = normalizeVatRateValue(line.rate);
  const baseValue = parseNumberInput(line.base);
  if (rateValue === null || baseValue === null) {
    return null;
  }
  const rate = Number(rateValue);
  const vatAmount = roundAmount(baseValue * (rate / 100));
  const total = roundAmount(baseValue + vatAmount);
  return {
    rate,
    base: roundAmount(baseValue),
    vat_amount: vatAmount,
    total,
  };
}

function summarizeVatBreakdown(lines) {
  if (!lines.length) {
    return null;
  }
  let baseTotal = 0;
  let vatTotal = 0;
  let totalTotal = 0;
  lines.forEach((line) => {
    const normalized = normalizeBreakdownLine(line);
    if (!normalized) {
      return;
    }
    baseTotal += normalized.base;
    vatTotal += normalized.vat_amount;
    totalTotal += normalized.total;
  });
  return {
    base: roundAmount(baseTotal),
    vatAmount: roundAmount(vatTotal),
    total: roundAmount(totalTotal),
  };
}

function getBreakdownRates(lines) {
  const rates = new Set();
  lines.forEach((line) => {
    const normalized = normalizeBreakdownLine(line);
    if (normalized) {
      rates.add(normalized.rate);
    }
  });
  return [...rates];
}

function buildVatBreakdownPayload(lines) {
  const payload = [];
  lines.forEach((line) => {
    const normalized = normalizeBreakdownLine(line);
    if (!normalized) {
      return;
    }
    payload.push({
      rate: normalized.rate,
      base: normalized.base,
      vat_amount: normalized.vat_amount,
      total: normalized.total,
    });
  });
  return payload;
}

function getPrimaryVatRateFromBreakdown(lines, fallback = "21") {
  const normalized = buildVatBreakdownPayload(lines);
  if (!normalized.length) {
    return fallback;
  }
  return String(normalized[0].rate);
}

function getVatDisplayFromInvoice(invoice) {
  const breakdown = parseVatBreakdown(invoice.vat_breakdown || invoice.vatBreakdown);
  const rates = getBreakdownRates(breakdown);
  if (rates.length > 1) {
    return {
      label: "Mixto",
      title: `IVA: ${rates.map((rate) => `${rate}%`).join(" · ")}`,
    };
  }
  if (rates.length === 1) {
    return { label: `${rates[0]}%`, title: "" };
  }
  return { label: `${invoice.vat_rate}%`, title: "" };
}

function formatExpenseCategory(category) {
  return expenseCategoryLabels[category] || expenseCategoryLabels.with_invoice;
}

function createExpenseCategorySelect(selected) {
  const select = document.createElement("select");
  [
    { value: "with_invoice", label: expenseCategoryLabels.with_invoice },
    { value: "without_invoice", label: expenseCategoryLabels.without_invoice },
    { value: "non_deductible", label: expenseCategoryLabels.non_deductible },
  ].forEach((optionData) => {
    const option = document.createElement("option");
    option.value = optionData.value;
    option.textContent = optionData.label;
    select.appendChild(option);
  });
  select.value = selected || "with_invoice";
  return select;
}

function formatNoInvoiceType(type) {
  return noInvoiceTypeLabels[type] || noInvoiceTypeLabels.otro;
}

function createNoInvoiceTypeSelect(selected) {
  const select = document.createElement("select");
  Object.keys(noInvoiceTypeLabels).forEach((key) => {
    const option = document.createElement("option");
    option.value = key;
    option.textContent = noInvoiceTypeLabels[key];
    select.appendChild(option);
  });
  select.value = selected || "otro";
  return select;
}

function updateSupplierSuggestions() {
  if (!supplierSuggestions) {
    return;
  }
  const names = new Set();
  currentInvoices.forEach((invoice) => {
    if (invoice.supplier) {
      names.add(String(invoice.supplier).trim());
    }
  });
  supplierSuggestions.innerHTML = "";
  [...names]
    .filter((name) => name.length > 0)
    .sort((a, b) => a.localeCompare(b))
    .forEach((name) => {
      const option = document.createElement("option");
      option.value = name;
      supplierSuggestions.appendChild(option);
    });
}

function createDeductibleSelect(selected) {
  const select = document.createElement("select");
  [
    { value: "true", label: "Sí" },
    { value: "false", label: "No" },
  ].forEach((optionData) => {
    const option = document.createElement("option");
    option.value = optionData.value;
    option.textContent = optionData.label;
    select.appendChild(option);
  });
  select.value = selected ? "true" : "false";
  return select;
}

function setYearOptions(select, years) {
  if (!select) {
    return;
  }
  const current = select.value;
  select.innerHTML = "";
  years.forEach((year) => {
    const option = document.createElement("option");
    option.value = String(year);
    option.textContent = String(year);
    select.appendChild(option);
  });
  if (current && [...select.options].some((o) => o.value === current)) {
    select.value = current;
  } else {
    select.value = String(years[years.length - 1]);
  }
}

function loadYears() {
  const companyId = getSelectedCompanyId();
  const suffix = companyId ? `?company_id=${companyId}` : "";
  return fetch(`/api/years${suffix}`)
    .then((res) => res.json())
    .then((data) => {
      const currentYear = new Date().getFullYear();
      const yearSet = new Set((data.years || []).map(Number));
      yearSet.add(currentYear);
      const years = Array.from(yearSet).sort((a, b) => a - b);
      setYearOptions(yearSelect, years);
      setYearOptions(billingYearSelect, years);
      if (reportYearSelect) {
        setYearOptions(reportYearSelect, years);
      }
    });
}

function setCompanyOptions(list) {
  if (!companySelect) {
    return;
  }
  companySelect.innerHTML = "";
  if (!list.length) {
    companySelect.disabled = true;
    selectedCompanyId = null;
    return;
  }
  companySelect.disabled = false;
  list.forEach((company) => {
    const option = document.createElement("option");
    option.value = String(company.id);
    option.textContent = company.display_name;
    companySelect.appendChild(option);
  });
  if (selectedCompanyId && list.some((c) => String(c.id) === String(selectedCompanyId))) {
    companySelect.value = String(selectedCompanyId);
  } else {
    selectedCompanyId = String(list[0].id);
    companySelect.value = selectedCompanyId;
    persistFilters();
  }
  applyCompanyTaxModules();
  updateHeaderContext();
}

function loadStaff() {
  if (currentUserRole === "staff") {
    return Promise.resolve([]);
  }
  return fetch("/api/staff")
    .then((res) => res.json())
    .then((data) => {
      staffMembers = data.staff || [];
      renderStaffTable(staffMembers);
      populateStaffSelectOptions();
      return staffMembers;
    })
    .catch(() => {
      staffMembers = [];
      renderStaffTable([]);
      populateStaffSelectOptions();
      return [];
    });
}

function populateStaffSelectOptions() {
  if (!companyAssignedSelect) {
    return;
  }
  companyAssignedSelect.innerHTML = "";
  const emptyOption = document.createElement("option");
  emptyOption.value = "";
  emptyOption.textContent = "Sin asignar";
  companyAssignedSelect.appendChild(emptyOption);
  staffMembers.forEach((member) => {
    const option = document.createElement("option");
    option.value = String(member.id);
    option.textContent = member.email;
    companyAssignedSelect.appendChild(option);
  });
}

function renderStaffTable(list) {
  if (!staffTableBody || !staffEmpty) {
    return;
  }
  staffTableBody.innerHTML = "";
  if (!list.length) {
    staffEmpty.style.display = "block";
    return;
  }
  staffEmpty.style.display = "none";
  list.forEach((member) => {
    const tr = document.createElement("tr");
    const emailTd = document.createElement("td");
    emailTd.textContent = member.email;
    const statusTd = document.createElement("td");
    statusTd.textContent = member.is_active ? "Activo" : "Inactivo";
    const actionsTd = document.createElement("td");
    actionsTd.classList.add("billing-actions");

    const toggleBtn = document.createElement("button");
    toggleBtn.type = "button";
    toggleBtn.className = "button ghost";
    toggleBtn.textContent = member.is_active ? "Desactivar" : "Activar";
    toggleBtn.addEventListener("click", () => {
      updateStaff(member.id, { is_active: !member.is_active });
    });

    actionsTd.appendChild(toggleBtn);

    tr.appendChild(emailTd);
    tr.appendChild(statusTd);
    tr.appendChild(actionsTd);
    staffTableBody.appendChild(tr);
  });
}

function createStaffSelect(selectedId) {
  const select = document.createElement("select");
  const emptyOption = document.createElement("option");
  emptyOption.value = "";
  emptyOption.textContent = "Sin asignar";
  select.appendChild(emptyOption);
  staffMembers.forEach((member) => {
    const option = document.createElement("option");
    option.value = String(member.id);
    option.textContent = member.email;
    select.appendChild(option);
  });
  if (selectedId) {
    select.value = String(selectedId);
  }
  return select;
}

function loadCompanies() {
  return fetch("/api/companies")
    .then((res) => res.json())
    .then((data) => {
      companies = data.companies || [];
      setCompanyOptions(companies);
      renderCompaniesTable(companies);
      applyCompanyTaxModules();
      updatePnlSummary();
      updateHeaderContext();
      return companies;
    });
}

function renderCompaniesTable(list) {
  if (!companiesTableBody || !companiesEmpty) {
    return;
  }
  companiesTableBody.innerHTML = "";
  if (!list.length) {
    companiesEmpty.style.display = "block";
    return;
  }
  companiesEmpty.style.display = "none";
  list.forEach((company) => {
    const tr = document.createElement("tr");
    tr.dataset.id = company.id;

    const displayTd = document.createElement("td");
    displayTd.textContent = company.display_name;
    const legalTd = document.createElement("td");
    legalTd.textContent = company.legal_name;
    const taxTd = document.createElement("td");
    taxTd.textContent = company.tax_id;
    const emailTd = document.createElement("td");
    emailTd.textContent = company.email || "-";
    const phoneTd = document.createElement("td");
    phoneTd.textContent = company.phone || "-";
    const assignedTd = document.createElement("td");
    if (currentUserRole === "staff") {
      assignedTd.textContent = "Asignado a ti";
    } else {
      assignedTd.textContent = getStaffEmail(company.assigned_user_id) || "Sin asignar";
    }
    const typeTd = document.createElement("td");
    typeTd.textContent = company.company_type === "individual" ? "Autónomo" : "Sociedad";
    const actionsTd = document.createElement("td");
    actionsTd.classList.add("billing-actions");

    if (currentUserRole !== "staff") {
      const editBtn = document.createElement("button");
      editBtn.type = "button";
      editBtn.className = "button ghost";
      editBtn.textContent = "Editar";
      editBtn.addEventListener("click", () => {
        enterCompanyEditMode(tr, company);
      });

      const deleteBtn = document.createElement("button");
      deleteBtn.type = "button";
      deleteBtn.className = "button danger";
      deleteBtn.textContent = "Eliminar";
      deleteBtn.addEventListener("click", () => {
        if (!confirm("¿Seguro que deseas eliminar esta empresa?")) {
          return;
        }
        deleteCompany(company.id);
      });

      actionsTd.appendChild(editBtn);
      actionsTd.appendChild(deleteBtn);
    } else {
      actionsTd.textContent = "-";
    }

    tr.appendChild(displayTd);
    tr.appendChild(legalTd);
    tr.appendChild(taxTd);
    tr.appendChild(emailTd);
    tr.appendChild(phoneTd);
    tr.appendChild(assignedTd);
    tr.appendChild(typeTd);
    tr.appendChild(actionsTd);
    companiesTableBody.appendChild(tr);
  });
}

function getStaffEmail(staffId) {
  if (!staffId) {
    return "";
  }
  const found = staffMembers.find((member) => String(member.id) === String(staffId));
  return found ? found.email : "";
}

function enterCompanyEditMode(row, company) {
  const displayTd = row.children[0];
  const legalTd = row.children[1];
  const taxTd = row.children[2];
  const emailTd = row.children[3];
  const phoneTd = row.children[4];
  const assignedTd = row.children[5];
  const typeTd = row.children[6];
  const actionsTd = row.children[7];

  const displayInput = document.createElement("input");
  displayInput.type = "text";
  displayInput.value = company.display_name;
  const legalInput = document.createElement("input");
  legalInput.type = "text";
  legalInput.value = company.legal_name;
  const taxInput = document.createElement("input");
  taxInput.type = "text";
  taxInput.value = company.tax_id;
  const emailInput = document.createElement("input");
  emailInput.type = "email";
  emailInput.value = company.email || "";
  const phoneInput = document.createElement("input");
  phoneInput.type = "text";
  phoneInput.value = company.phone || "";
  const assignedSelect = createStaffSelect(company.assigned_user_id);
  const typeSelect = document.createElement("select");
  const optionIndividual = document.createElement("option");
  optionIndividual.value = "individual";
  optionIndividual.textContent = "Autónomo";
  const optionCompany = document.createElement("option");
  optionCompany.value = "company";
  optionCompany.textContent = "Sociedad";
  typeSelect.appendChild(optionIndividual);
  typeSelect.appendChild(optionCompany);
  typeSelect.value = company.company_type;

  displayTd.textContent = "";
  displayTd.appendChild(displayInput);
  legalTd.textContent = "";
  legalTd.appendChild(legalInput);
  taxTd.textContent = "";
  taxTd.appendChild(taxInput);
  emailTd.textContent = "";
  emailTd.appendChild(emailInput);
  phoneTd.textContent = "";
  phoneTd.appendChild(phoneInput);
  assignedTd.textContent = "";
  assignedTd.appendChild(assignedSelect);
  typeTd.textContent = "";
  typeTd.appendChild(typeSelect);

  actionsTd.innerHTML = "";
  const saveBtn = document.createElement("button");
  saveBtn.type = "button";
  saveBtn.className = "button primary";
  saveBtn.textContent = "Guardar";
  saveBtn.addEventListener("click", () => {
    updateCompany(company.id, {
      display_name: displayInput.value,
      legal_name: legalInput.value,
      tax_id: taxInput.value,
      email: emailInput.value,
      phone: phoneInput.value,
      assigned_user_id: assignedSelect.value,
      company_type: typeSelect.value,
    });
  });
  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "button ghost";
  cancelBtn.textContent = "Cancelar";
  cancelBtn.addEventListener("click", () => {
    renderCompaniesTable(companies);
  });
  actionsTd.appendChild(saveBtn);
  actionsTd.appendChild(cancelBtn);
}

function saveCompany() {
  if (
    !companyDisplayName ||
    !companyLegalName ||
    !companyTaxId ||
    !companyType ||
    !companyEmail ||
    !companyPhone ||
    !companyAssignedSelect
  ) {
    return;
  }
  companySaveBtn.disabled = true;
  fetch("/api/companies", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      display_name: companyDisplayName.value,
      legal_name: companyLegalName.value,
      tax_id: companyTaxId.value,
      company_type: companyType.value,
      email: companyEmail.value,
      phone: companyPhone.value,
      assigned_user_id: companyAssignedSelect.value,
    }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al guardar."]).join("\n"));
        return;
      }
      companyDisplayName.value = "";
      companyLegalName.value = "";
      companyTaxId.value = "";
      companyEmail.value = "";
      companyPhone.value = "";
      companyAssignedSelect.value = "";
      loadCompanies().then(() => {
        refreshAllData();
      });
    })
    .catch(() => {
      alert("No se pudo guardar la empresa.");
    })
    .finally(() => {
      companySaveBtn.disabled = false;
    });
}

function updateCompany(companyId, payload) {
  fetch(`/api/companies/${companyId}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al actualizar."]).join("\n"));
        return;
      }
      loadCompanies().then(() => {
        refreshAllData();
      });
    })
    .catch(() => {
      alert("No se pudo actualizar la empresa.");
    });
}

function deleteCompany(companyId) {
  fetch(`/api/companies/${companyId}`, {
    method: "DELETE",
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al eliminar."]).join("\n"));
        return;
      }
      if (String(selectedCompanyId) === String(companyId)) {
        selectedCompanyId = null;
      }
      loadCompanies().then(() => {
        refreshAllData();
      });
    })
    .catch(() => {
      alert("No se pudo eliminar la empresa.");
    });
}

function saveStaff() {
  if (!staffEmail || !staffPassword || !staffSaveBtn) {
    return;
  }
  if (!staffEmail.value || !staffPassword.value) {
    alert("Email y contraseña son obligatorios.");
    return;
  }
  staffSaveBtn.disabled = true;
  fetch("/api/staff", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      email: staffEmail.value,
      password: staffPassword.value,
    }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al crear el trabajador."]).join("\n"));
        return;
      }
      staffEmail.value = "";
      staffPassword.value = "";
      loadStaff().then(() => renderCompaniesTable(companies));
    })
    .catch(() => {
      alert("No se pudo crear el trabajador.");
    })
    .finally(() => {
      staffSaveBtn.disabled = false;
    });
}

function updateStaff(staffId, payload) {
  fetch(`/api/staff/${staffId}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["No se pudo actualizar el trabajador."]).join("\n"));
        return;
      }
      loadStaff().then(() => renderCompaniesTable(companies));
    })
    .catch(() => {
      alert("No se pudo actualizar el trabajador.");
    });
}

function getSelectedPeriod() {
  if (!periodSelect) {
    return "monthly";
  }
  return periodSelect.value || "monthly";
}

function getSelectedMonthYear() {
  if (!monthSelect || !yearSelect) {
    const now = new Date();
    return {
      month: now.getMonth() + 1,
      year: now.getFullYear(),
    };
  }
  return {
    month: Number(monthSelect.value),
    year: Number(yearSelect.value),
  };
}

function getSelectedCompanyId() {
  return selectedCompanyId ? Number(selectedCompanyId) : null;
}

function getSelectedCompany() {
  if (!selectedCompanyId) {
    return null;
  }
  return (
    companies.find((company) => String(company.id) === String(selectedCompanyId)) ||
    null
  );
}

function getSelectedCompanyType() {
  const company = getSelectedCompany();
  return company ? company.company_type : null;
}

// Los módulos fiscales se muestran según el tipo de la empresa seleccionada.
function applyCompanyTaxModules() {
  const companyType = getSelectedCompanyType();
  const target =
    companyType === "company" ? "is" : companyType === "individual" ? "irpf" : null;
  document.querySelectorAll("[data-tax-module]").forEach((panel) => {
    if (!target) {
      panel.style.display = "none";
      return;
    }
    panel.style.display = panel.dataset.taxModule === target ? "" : "none";
  });
}

function getQuarterMonths(month) {
  const quarterIndex = Math.floor((month - 1) / 3);
  const start = quarterIndex * 3 + 1;
  return [start, start + 1, start + 2];
}

function getPeriodLabel() {
  const { month, year } = getSelectedMonthYear();
  if (!month || !year) {
    return "";
  }
  const period = getSelectedPeriod();
  if (period === "quarterly") {
    const quarterIndex = Math.floor((month - 1) / 3) + 1;
    return `T${quarterIndex} ${year}`;
  }
  return `${monthNames[month - 1]} ${year}`;
}

function updateHeaderContext() {
  if (!monthSelect || !yearSelect || !periodSelect) {
    return;
  }
  if (headerPeriodLabel) {
    const label = getPeriodLabel();
    headerPeriodLabel.textContent = label ? `Periodo: ${label}` : "";
  }
  if (headerCompanyLabel) {
    const company = getSelectedCompany();
    headerCompanyLabel.textContent = company
      ? `Empresa: ${company.display_name}`
      : "Empresa: -";
  }
}

function applyRoleVisibility() {
  if (currentUserRole === "staff") {
    const companyForm = document.querySelector(".companies-panel .billing-form");
    if (companyForm) {
      companyForm.style.display = "none";
    }
    const staffPanel = document.querySelector(".staff-panel");
    if (staffPanel) {
      staffPanel.style.display = "none";
    }
  }
}

function persistFilters() {
  if (!monthSelect || !yearSelect || !periodSelect) {
    return;
  }
  localStorage.setItem("selectedMonth", monthSelect.value);
  localStorage.setItem("selectedYear", yearSelect.value);
  localStorage.setItem("selectedPeriod", periodSelect.value);
  if (selectedCompanyId) {
    localStorage.setItem("selectedCompanyId", String(selectedCompanyId));
  }
}

function restoreFilters(now) {
  if (!monthSelect || !yearSelect || !periodSelect) {
    return;
  }
  const storedMonth = localStorage.getItem("selectedMonth");
  const storedYear = localStorage.getItem("selectedYear");
  const storedPeriod = localStorage.getItem("selectedPeriod");

  if (storedMonth && [...monthSelect.options].some((opt) => opt.value === storedMonth)) {
    monthSelect.value = storedMonth;
  } else {
    monthSelect.value = String(now.getMonth() + 1);
  }

  if (storedYear && [...yearSelect.options].some((opt) => opt.value === storedYear)) {
    yearSelect.value = storedYear;
  } else {
    yearSelect.value = String(now.getFullYear());
  }

  if (storedPeriod && [...periodSelect.options].some((opt) => opt.value === storedPeriod)) {
    periodSelect.value = storedPeriod;
  } else {
    periodSelect.value = "monthly";
  }

  const storedCompany = localStorage.getItem("selectedCompanyId");
  if (storedCompany) {
    selectedCompanyId = storedCompany;
  }
}

function addFiles(fileList) {
  Array.from(fileList).forEach((file) => {
    if (!isAllowedFile(file.name)) {
      return;
    }
    const item = {
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      file,
      originalFilename: file.name,
      storedFilename: "",
      date: new Date().toISOString().slice(0, 10),
      paymentDate: "",
      paymentDates: [],
      supplier: "",
      base: "",
      vat: "21",
      vatAmount: "",
      total: "",
      vatBreakdown: [],
      vatBreakdownOpen: false,
      analysisText: "",
      analysisPending: true,
      analysisError: false,
      analysisErrorMessage: "",
      analysisStatus: "ok",
      touched: {
        date: false,
        supplier: false,
        base: false,
        vat: false,
        vatAmount: false,
        total: false,
      },
    };
    pendingFiles.push(item);
    analyzeInvoiceForItem(item);
  });
  renderTable();
}

function addIncomeFiles(fileList) {
  Array.from(fileList).forEach((file) => {
    if (!isAllowedFile(file.name)) {
      return;
    }
    const item = {
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      file,
      originalFilename: file.name,
      storedFilename: "",
      date: new Date().toISOString().slice(0, 10),
      paymentDate: "",
      paymentDates: [],
      client: "",
      base: "",
      vat: "21",
      vatAmount: "",
      total: "",
      vatBreakdown: [],
      vatBreakdownOpen: false,
      analysisText: "",
      analysisPending: true,
      analysisError: false,
      analysisErrorMessage: "",
      analysisStatus: "ok",
      touched: {
        date: false,
        client: false,
        base: false,
        vat: false,
        vatAmount: false,
        total: false,
      },
    };
    pendingIncomeFiles.push(item);
    analyzeIncomeForItem(item);
  });
  renderIncomeTable();
}

function renderTable() {
  uploadTableBody.innerHTML = "";
  if (pendingFiles.length === 0) {
    emptyMessage.style.display = "block";
    uploadBtn.disabled = false;
    if (globalProcessing) {
      globalProcessing.classList.remove("is-visible", "error");
    }
    return;
  }
  emptyMessage.style.display = "none";
  uploadBtn.disabled = pendingFiles.some((item) => item.analysisPending);

  if (globalProcessing && globalProcessingText) {
    const hasPending = pendingFiles.some((item) => item.analysisPending);
    const hasError = pendingFiles.some((item) => item.analysisError);
    if (hasPending) {
      globalProcessing.classList.add("is-visible");
      globalProcessing.classList.remove("error");
      globalProcessingText.textContent =
        "Procesando factura… Puede tardar hasta 2 minutos.";
    } else if (hasError) {
      globalProcessing.classList.add("is-visible", "error");
      globalProcessingText.textContent =
        "El análisis automático no ha sido posible. Puedes editar los datos manualmente.";
    } else {
      globalProcessing.classList.remove("is-visible", "error");
    }
  }

  pendingFiles.forEach((item) => {
    const tr = document.createElement("tr");
    tr.dataset.id = item.id;
    if (item.analysisPending) {
      tr.classList.add("is-processing");
    }

    const nameTd = document.createElement("td");
    nameTd.textContent = item.file.name;

    const dateTd = document.createElement("td");
    const dateInput = document.createElement("input");
    dateInput.type = "date";
    dateInput.value = item.date;
    dateInput.disabled = item.analysisPending;
    dateInput.addEventListener("change", () => {
      item.date = dateInput.value;
      item.touched.date = true;
    });
    dateTd.appendChild(dateInput);

    const supplierTd = document.createElement("td");
    const supplierInput = document.createElement("input");
    supplierInput.type = "text";
    supplierInput.placeholder = "Proveedor";
    supplierInput.value = item.supplier;
    supplierInput.disabled = item.analysisPending;
    if (supplierSuggestions) {
      supplierInput.setAttribute("list", "supplierSuggestions");
    }
    const supplierWarning = document.createElement("div");
    supplierWarning.className = "field-warning";
    const updateSupplierWarning = () => {
      const value = supplierInput.value.trim();
      if (!value) {
        supplierWarning.textContent = "Proveedor pendiente de completar.";
        supplierWarning.style.display = "block";
        supplierInput.classList.add("input-warning");
        return;
      }
      if (isSupplierSameAsCompany(value)) {
        supplierWarning.textContent =
          "El proveedor no puede ser la empresa activa.";
        supplierWarning.style.display = "block";
        supplierInput.classList.add("input-warning");
        return;
      }
      supplierWarning.textContent = "";
      supplierWarning.style.display = "none";
      supplierInput.classList.remove("input-warning");
    };
    supplierInput.addEventListener("input", () => {
      item.supplier = supplierInput.value;
      item.touched.supplier = true;
      updateSupplierWarning();
    });
    supplierTd.appendChild(supplierInput);
    supplierTd.appendChild(supplierWarning);

    const baseTd = document.createElement("td");
    const baseInput = document.createElement("input");
    baseInput.type = "number";
    baseInput.step = "0.01";
    baseInput.min = "0";
    baseInput.placeholder = "0,00";
    baseInput.value = item.base;
    const breakdownActive =
      Array.isArray(item.vatBreakdown) && item.vatBreakdown.length > 0;
    baseInput.disabled = item.analysisPending;
    baseInput.readOnly = breakdownActive;
    baseInput.addEventListener("input", () => {
      item.base = baseInput.value;
      item.touched.base = true;
      applyVatCalculation(item, {
        base: baseInput,
        vat: vatSelect,
        vatAmount: vatAmountInput,
        total: totalInput,
      }, "base");
    });
    baseTd.appendChild(baseInput);

    const vatTd = document.createElement("td");
    const vatSelect = document.createElement("select");
    ["0", "4", "10", "21"].forEach((rate) => {
      const option = document.createElement("option");
      option.value = rate;
      option.textContent = `${rate}%`;
      vatSelect.appendChild(option);
    });
    applyVatSelection(vatSelect, item.vat);
    vatSelect.disabled = item.analysisPending || breakdownActive;
    vatSelect.addEventListener("change", () => {
      item.vat = resolveVatRateValue(vatSelect.value);
      item.touched.vat = true;
      applyVatCalculation(item, {
        base: baseInput,
        vat: vatSelect,
        vatAmount: vatAmountInput,
        total: totalInput,
      }, "vat");
    });
    vatTd.appendChild(vatSelect);
    if (breakdownActive && getBreakdownRates(item.vatBreakdown).length > 1) {
      const mixedBadge = document.createElement("span");
      mixedBadge.className = "vat-mixed-badge";
      mixedBadge.textContent = "Mixto";
      vatTd.appendChild(mixedBadge);
    }
    const addBreakdownBtn = document.createElement("button");
    addBreakdownBtn.type = "button";
    addBreakdownBtn.className = "link-button vat-breakdown-toggle";
    addBreakdownBtn.textContent = breakdownActive
      ? "Añadir línea IVA"
      : "Añadir línea IVA";
    addBreakdownBtn.disabled = item.analysisPending;
    addBreakdownBtn.addEventListener("click", () => {
      item.vatBreakdown = item.vatBreakdown || [];
      item.vatBreakdown.push({ rate: item.vat || "21", base: "", vat_amount: "", total: "" });
      item.vatBreakdownOpen = true;
      renderTable();
    });
    vatTd.appendChild(addBreakdownBtn);

    const vatAmountTd = document.createElement("td");
    const vatAmountInput = document.createElement("input");
    vatAmountInput.type = "number";
    vatAmountInput.step = "0.01";
    vatAmountInput.min = "0";
    vatAmountInput.placeholder = "0,00";
    vatAmountInput.value = item.vatAmount;
    vatAmountInput.disabled = item.analysisPending;
    vatAmountInput.readOnly = breakdownActive;
    vatAmountInput.addEventListener("input", () => {
      item.vatAmount = vatAmountInput.value;
      item.touched.vatAmount = true;
      applyVatCalculation(item, {
        base: baseInput,
        vat: vatSelect,
        vatAmount: vatAmountInput,
        total: totalInput,
      }, "vatAmount");
    });
    vatAmountTd.appendChild(vatAmountInput);

    const totalTd = document.createElement("td");
    const totalInput = document.createElement("input");
    totalInput.type = "number";
    totalInput.step = "0.01";
    totalInput.min = "0";
    totalInput.placeholder = "0,00";
    totalInput.value = item.total;
    totalInput.disabled = item.analysisPending;
    totalInput.readOnly = breakdownActive;
    totalInput.addEventListener("input", () => {
      item.total = totalInput.value;
      item.touched.total = true;
      applyVatCalculation(item, {
        base: baseInput,
        vat: vatSelect,
        vatAmount: vatAmountInput,
        total: totalInput,
      }, "total");
    });
    totalTd.appendChild(totalInput);

    const actionsTd = document.createElement("td");
    actionsTd.classList.add("row-actions");
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.textContent = "Quitar";
    removeBtn.disabled = item.analysisPending;
    removeBtn.addEventListener("click", () => {
      const index = pendingFiles.findIndex((entry) => entry.id === item.id);
      if (index !== -1) {
        pendingFiles.splice(index, 1);
        renderTable();
      }
    });
    actionsTd.appendChild(removeBtn);

    tr.appendChild(nameTd);
    tr.appendChild(dateTd);
    tr.appendChild(supplierTd);
    tr.appendChild(baseTd);
    tr.appendChild(vatTd);
    tr.appendChild(vatAmountTd);
    tr.appendChild(totalTd);
    tr.appendChild(actionsTd);
    uploadTableBody.appendChild(tr);

    if (item.vatBreakdown && item.vatBreakdown.length) {
      item.vatBreakdown.forEach((line, lineIndex) => {
        const row = document.createElement("tr");
        row.className = "vat-breakdown-inline-row";
        row.innerHTML = `
          <td></td>
          <td></td>
          <td></td>
        `;
        const baseCell = document.createElement("td");
        const rateCell = document.createElement("td");
        const vatCell = document.createElement("td");
        const totalCell = document.createElement("td");
        const actionsCell = document.createElement("td");
        actionsCell.className = "row-actions";

        const rateSelect = document.createElement("select");
        ["0", "4", "10", "21"].forEach((rate) => {
          const option = document.createElement("option");
          option.value = rate;
          option.textContent = `${rate}%`;
          rateSelect.appendChild(option);
        });
        rateSelect.value = resolveVatRateValue(line.rate);
        rateSelect.disabled = item.analysisPending;
        rateCell.appendChild(rateSelect);

        const baseLineInput = document.createElement("input");
        baseLineInput.type = "number";
        baseLineInput.step = "0.01";
        baseLineInput.min = "0";
        baseLineInput.placeholder = "0,00";
        baseLineInput.value = line.base || "";
        baseLineInput.disabled = item.analysisPending;
        baseCell.appendChild(baseLineInput);

        const vatLineInput = document.createElement("input");
        vatLineInput.type = "number";
        vatLineInput.step = "0.01";
        vatLineInput.min = "0";
        vatLineInput.readOnly = true;
        vatLineInput.value = line.vat_amount || "";
        vatCell.appendChild(vatLineInput);

        const totalLineInput = document.createElement("input");
        totalLineInput.type = "number";
        totalLineInput.step = "0.01";
        totalLineInput.min = "0";
        totalLineInput.readOnly = true;
        totalLineInput.value = line.total || "";
        totalCell.appendChild(totalLineInput);

        const removeLineBtn = document.createElement("button");
        removeLineBtn.type = "button";
        removeLineBtn.className = "button ghost";
        removeLineBtn.textContent = "Quitar";
        removeLineBtn.disabled = item.analysisPending;
        removeLineBtn.addEventListener("click", () => {
          item.vatBreakdown.splice(lineIndex, 1);
          if (!item.vatBreakdown.length) {
            item.vatBreakdownOpen = false;
          }
          renderTable();
        });
        actionsCell.appendChild(removeLineBtn);

        const syncLine = () => {
          line.rate = rateSelect.value;
          line.base = baseLineInput.value;
          const normalized = normalizeBreakdownLine(line);
          if (normalized) {
            line.vat_amount = formatAmountInput(normalized.vat_amount);
            line.total = formatAmountInput(normalized.total);
            vatLineInput.value = line.vat_amount;
            totalLineInput.value = line.total;
          } else {
            line.vat_amount = "";
            line.total = "";
            vatLineInput.value = "";
            totalLineInput.value = "";
          }
          const totals = summarizeVatBreakdown(item.vatBreakdown || []);
          if (totals) {
            item.base = formatAmountInput(totals.base);
            item.vatAmount = formatAmountInput(totals.vatAmount);
            item.total = formatAmountInput(totals.total);
            baseInput.value = item.base;
            vatAmountInput.value = item.vatAmount;
            totalInput.value = item.total;
          }
        };
        rateSelect.addEventListener("change", syncLine);
        baseLineInput.addEventListener("input", syncLine);
        syncLine();

        row.appendChild(baseCell);
        row.appendChild(rateCell);
        row.appendChild(vatCell);
        row.appendChild(totalCell);
        row.appendChild(actionsCell);
        uploadTableBody.appendChild(row);
      });
    }

    updateSupplierWarning();
    const initialSource =
      parseNumberInput(baseInput.value) !== null ? "base" : "total";
    if (breakdownActive) {
      const totals = summarizeVatBreakdown(item.vatBreakdown || []);
      if (totals) {
        item.base = formatAmountInput(totals.base);
        item.vatAmount = formatAmountInput(totals.vatAmount);
        item.total = formatAmountInput(totals.total);
        baseInput.value = item.base;
        vatAmountInput.value = item.vatAmount;
        totalInput.value = item.total;
      }
    } else {
      applyVatCalculation(
        item,
        {
          base: baseInput,
          vat: vatSelect,
          vatAmount: vatAmountInput,
          total: totalInput,
        },
        initialSource
      );
    }

    if (item.analysisPending || item.analysisError) {
      const statusRow = document.createElement("tr");
      statusRow.className = "processing-row";
      if (item.analysisError) {
        statusRow.classList.add("error");
      }
      const statusTd = document.createElement("td");
      statusTd.colSpan = 8;
      const statusWrapper = document.createElement("div");
      statusWrapper.className = "processing-status";
      if (item.analysisPending) {
        const spinner = document.createElement("span");
        spinner.className = "spinner";
        statusWrapper.appendChild(spinner);
      }
      const message = document.createElement("span");
      message.textContent = item.analysisError
        ? item.analysisErrorMessage || ANALYSIS_ERROR_MESSAGE
        : "Analizando factura… Las facturas escaneadas pueden tardar hasta 1 minuto.";
      statusWrapper.appendChild(message);
      statusTd.appendChild(statusWrapper);
      statusRow.appendChild(statusTd);
      uploadTableBody.appendChild(statusRow);
    }
  });
}

function renderIncomeTable() {
  if (!incomeUploadTableBody || !incomeEmptyMessage) {
    return;
  }
  incomeUploadTableBody.innerHTML = "";
  if (pendingIncomeFiles.length === 0) {
    incomeEmptyMessage.style.display = "block";
    if (incomeUploadBtn) {
      incomeUploadBtn.disabled = false;
    }
    return;
  }
  incomeEmptyMessage.style.display = "none";
  if (incomeUploadBtn) {
    incomeUploadBtn.disabled = pendingIncomeFiles.some((item) => item.analysisPending);
  }

  pendingIncomeFiles.forEach((item) => {
    const tr = document.createElement("tr");
    tr.dataset.id = item.id;
    if (item.analysisPending) {
      tr.classList.add("is-processing");
    }

    const nameTd = document.createElement("td");
    nameTd.textContent = item.file.name;

    const dateTd = document.createElement("td");
    const dateInput = document.createElement("input");
    dateInput.type = "date";
    dateInput.value = item.date;
    dateInput.disabled = item.analysisPending;
    dateInput.addEventListener("change", () => {
      item.date = dateInput.value;
      item.touched.date = true;
    });
    dateTd.appendChild(dateInput);

    const clientTd = document.createElement("td");
    const clientInput = document.createElement("input");
    clientInput.type = "text";
    clientInput.placeholder = "Cliente";
    clientInput.value = item.client;
    clientInput.disabled = item.analysisPending;
    clientInput.addEventListener("input", () => {
      item.client = clientInput.value;
      item.touched.client = true;
    });
    clientTd.appendChild(clientInput);

    const baseTd = document.createElement("td");
    const baseInput = document.createElement("input");
    baseInput.type = "number";
    baseInput.step = "0.01";
    baseInput.min = "0";
    baseInput.placeholder = "0,00";
    baseInput.value = item.base;
    baseInput.disabled = item.analysisPending;
    baseInput.addEventListener("input", () => {
      item.base = baseInput.value;
      item.touched.base = true;
      applyVatCalculation(item, {
        base: baseInput,
        vat: vatSelect,
        vatAmount: vatAmountInput,
        total: totalInput,
      }, "base");
    });
    baseTd.appendChild(baseInput);

    const vatTd = document.createElement("td");
    const vatSelect = document.createElement("select");
    ["0", "4", "10", "21"].forEach((rate) => {
      const option = document.createElement("option");
      option.value = rate;
      option.textContent = `${rate}%`;
      vatSelect.appendChild(option);
    });
    applyVatSelection(vatSelect, item.vat);
    vatSelect.disabled = item.analysisPending || breakdownActive;
    vatSelect.addEventListener("change", () => {
      item.vat = resolveVatRateValue(vatSelect.value);
      item.touched.vat = true;
      applyVatCalculation(item, {
        base: baseInput,
        vat: vatSelect,
        vatAmount: vatAmountInput,
        total: totalInput,
      }, "vat");
    });
    vatTd.appendChild(vatSelect);
    if (breakdownActive && getBreakdownRates(item.vatBreakdown).length > 1) {
      const mixedBadge = document.createElement("span");
      mixedBadge.className = "vat-mixed-badge";
      mixedBadge.textContent = "Mixto";
      vatTd.appendChild(mixedBadge);
    }
    const addBreakdownBtn = document.createElement("button");
    addBreakdownBtn.type = "button";
    addBreakdownBtn.className = "link-button vat-breakdown-toggle";
    addBreakdownBtn.textContent = "Añadir línea IVA";
    addBreakdownBtn.disabled = item.analysisPending;
    addBreakdownBtn.addEventListener("click", () => {
      item.vatBreakdown = item.vatBreakdown || [];
      item.vatBreakdown.push({ rate: item.vat || "21", base: "", vat_amount: "", total: "" });
      item.vatBreakdownOpen = true;
      renderIncomeTable();
    });
    vatTd.appendChild(addBreakdownBtn);

    const vatAmountTd = document.createElement("td");
    const vatAmountInput = document.createElement("input");
    vatAmountInput.type = "number";
    vatAmountInput.step = "0.01";
    vatAmountInput.min = "0";
    vatAmountInput.placeholder = "0,00";
    vatAmountInput.value = item.vatAmount;
    vatAmountInput.disabled = item.analysisPending;
    vatAmountInput.readOnly = breakdownActive;
    vatAmountInput.addEventListener("input", () => {
      item.vatAmount = vatAmountInput.value;
      item.touched.vatAmount = true;
      applyVatCalculation(item, {
        base: baseInput,
        vat: vatSelect,
        vatAmount: vatAmountInput,
        total: totalInput,
      }, "vatAmount");
    });
    vatAmountTd.appendChild(vatAmountInput);

    const totalTd = document.createElement("td");
    const totalInput = document.createElement("input");
    totalInput.type = "number";
    totalInput.step = "0.01";
    totalInput.min = "0";
    totalInput.placeholder = "0,00";
    totalInput.value = item.total;
    totalInput.disabled = item.analysisPending;
    totalInput.readOnly = breakdownActive;
    totalInput.addEventListener("input", () => {
      item.total = totalInput.value;
      item.touched.total = true;
      applyVatCalculation(item, {
        base: baseInput,
        vat: vatSelect,
        vatAmount: vatAmountInput,
        total: totalInput,
      }, "total");
    });
    totalTd.appendChild(totalInput);

    const actionsTd = document.createElement("td");
    actionsTd.classList.add("row-actions");
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.textContent = "Quitar";
    removeBtn.disabled = item.analysisPending;
    removeBtn.addEventListener("click", () => {
      const index = pendingIncomeFiles.findIndex((entry) => entry.id === item.id);
      if (index !== -1) {
        pendingIncomeFiles.splice(index, 1);
        renderIncomeTable();
      }
    });
    actionsTd.appendChild(removeBtn);

    tr.appendChild(nameTd);
    tr.appendChild(dateTd);
    tr.appendChild(clientTd);
    tr.appendChild(baseTd);
    tr.appendChild(vatTd);
    tr.appendChild(vatAmountTd);
    tr.appendChild(totalTd);
    tr.appendChild(actionsTd);
    incomeUploadTableBody.appendChild(tr);

    if (item.vatBreakdown && item.vatBreakdown.length) {
      item.vatBreakdown.forEach((line, lineIndex) => {
        const row = document.createElement("tr");
        row.className = "vat-breakdown-inline-row";
        row.innerHTML = `
          <td></td>
          <td></td>
          <td></td>
        `;
        const baseCell = document.createElement("td");
        const rateCell = document.createElement("td");
        const vatCell = document.createElement("td");
        const totalCell = document.createElement("td");
        const actionsCell = document.createElement("td");
        actionsCell.className = "row-actions";

        const rateSelect = document.createElement("select");
        ["0", "4", "10", "21"].forEach((rate) => {
          const option = document.createElement("option");
          option.value = rate;
          option.textContent = `${rate}%`;
          rateSelect.appendChild(option);
        });
        rateSelect.value = resolveVatRateValue(line.rate);
        rateSelect.disabled = item.analysisPending;
        rateCell.appendChild(rateSelect);

        const baseLineInput = document.createElement("input");
        baseLineInput.type = "number";
        baseLineInput.step = "0.01";
        baseLineInput.min = "0";
        baseLineInput.placeholder = "0,00";
        baseLineInput.value = line.base || "";
        baseLineInput.disabled = item.analysisPending;
        baseCell.appendChild(baseLineInput);

        const vatLineInput = document.createElement("input");
        vatLineInput.type = "number";
        vatLineInput.step = "0.01";
        vatLineInput.min = "0";
        vatLineInput.readOnly = true;
        vatLineInput.value = line.vat_amount || "";
        vatCell.appendChild(vatLineInput);

        const totalLineInput = document.createElement("input");
        totalLineInput.type = "number";
        totalLineInput.step = "0.01";
        totalLineInput.min = "0";
        totalLineInput.readOnly = true;
        totalLineInput.value = line.total || "";
        totalCell.appendChild(totalLineInput);

        const removeLineBtn = document.createElement("button");
        removeLineBtn.type = "button";
        removeLineBtn.className = "button ghost";
        removeLineBtn.textContent = "Quitar";
        removeLineBtn.disabled = item.analysisPending;
        removeLineBtn.addEventListener("click", () => {
          item.vatBreakdown.splice(lineIndex, 1);
          if (!item.vatBreakdown.length) {
            item.vatBreakdownOpen = false;
          }
          renderIncomeTable();
        });
        actionsCell.appendChild(removeLineBtn);

        const syncLine = () => {
          line.rate = rateSelect.value;
          line.base = baseLineInput.value;
          const normalized = normalizeBreakdownLine(line);
          if (normalized) {
            line.vat_amount = formatAmountInput(normalized.vat_amount);
            line.total = formatAmountInput(normalized.total);
            vatLineInput.value = line.vat_amount;
            totalLineInput.value = line.total;
          } else {
            line.vat_amount = "";
            line.total = "";
            vatLineInput.value = "";
            totalLineInput.value = "";
          }
          const totals = summarizeVatBreakdown(item.vatBreakdown || []);
          if (totals) {
            item.base = formatAmountInput(totals.base);
            item.vatAmount = formatAmountInput(totals.vatAmount);
            item.total = formatAmountInput(totals.total);
            baseInput.value = item.base;
            vatAmountInput.value = item.vatAmount;
            totalInput.value = item.total;
          }
        };
        rateSelect.addEventListener("change", syncLine);
        baseLineInput.addEventListener("input", syncLine);
        syncLine();

        row.appendChild(baseCell);
        row.appendChild(rateCell);
        row.appendChild(vatCell);
        row.appendChild(totalCell);
        row.appendChild(actionsCell);
        incomeUploadTableBody.appendChild(row);
      });
    }

    const initialSource =
      parseNumberInput(baseInput.value) !== null ? "base" : "total";
    if (breakdownActive) {
      const totals = summarizeVatBreakdown(item.vatBreakdown || []);
      if (totals) {
        item.base = formatAmountInput(totals.base);
        item.vatAmount = formatAmountInput(totals.vatAmount);
        item.total = formatAmountInput(totals.total);
        baseInput.value = item.base;
        vatAmountInput.value = item.vatAmount;
        totalInput.value = item.total;
      }
    } else {
      applyVatCalculation(
        item,
        {
          base: baseInput,
          vat: vatSelect,
          vatAmount: vatAmountInput,
          total: totalInput,
        },
        initialSource
      );
    }

    if (item.analysisPending || item.analysisError) {
      const statusRow = document.createElement("tr");
      statusRow.className = "processing-row";
      if (item.analysisError) {
        statusRow.classList.add("error");
      }
      const statusTd = document.createElement("td");
      statusTd.colSpan = 8;
      const statusWrapper = document.createElement("div");
      statusWrapper.className = "processing-status";
      if (item.analysisPending) {
        const spinner = document.createElement("span");
        spinner.className = "spinner";
        statusWrapper.appendChild(spinner);
      }
      const message = document.createElement("span");
      message.textContent = item.analysisError
        ? item.analysisErrorMessage || ANALYSIS_ERROR_MESSAGE
        : "Analizando factura… Las facturas escaneadas pueden tardar hasta 1 minuto.";
      statusWrapper.appendChild(message);
      statusTd.appendChild(statusWrapper);
      statusRow.appendChild(statusTd);
      incomeUploadTableBody.appendChild(statusRow);
    }
  });
}

function analyzeIncomeForItem(item) {
  const formData = new FormData();
  formData.append("file", item.file);
  formData.append("document_type", "income");
  const companyId = getSelectedCompanyId();
  if (companyId) {
    formData.append("company_id", companyId);
  }

  fetch("/api/analyze-invoice", {
    method: "POST",
    body: formData,
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        item.analysisPending = false;
        item.analysisError = true;
        item.analysisErrorMessage = ANALYSIS_ERROR_MESSAGE;
        renderIncomeTable();
        return;
      }
      const extracted = data.extracted || {};
      item.storedFilename = data.storedFilename || "";
      item.analysisText = extracted.analysis_text || "";
      item.analysisStatus = extracted.analysis_status || "ok";
      const extractedBreakdown = parseVatBreakdown(
        extracted.vat_breakdown || extracted.vatBreakdown
      );
      if (extractedBreakdown.length) {
        item.vatBreakdown = extractedBreakdown;
        item.vatBreakdownOpen = extractedBreakdown.length > 1;
      }
      if (extracted.analysis_status === "low_quality_scan") {
        item.analysisPending = false;
        item.analysisError = true;
        item.analysisErrorMessage = LOW_QUALITY_SCAN_MESSAGE;
        if (!lowQualityDismissedIds.has(item.id)) {
          showLowQualityModal();
          lowQualityDismissedIds.add(item.id);
        }
        renderIncomeTable();
        return;
      }
      const detectedClient = extracted.client_name || extracted.provider_name;

      if (!item.touched.client && detectedClient) {
        item.client = detectedClient;
      }
      if (!item.touched.date && extracted.invoice_date) {
        item.date = extracted.invoice_date;
      }
      if (Array.isArray(extracted.payment_dates) && extracted.payment_dates.length) {
        item.paymentDates = extracted.payment_dates.slice();
      }
      if (!item.paymentDate) {
        const primaryDate = item.paymentDates[0] || extracted.payment_date;
        item.paymentDate = computePaymentDate(item.date, primaryDate);
      }
      if (!item.touched.base && extracted.base_amount !== null && extracted.base_amount !== undefined) {
        item.base = String(extracted.base_amount);
      }
      if (!item.touched.vat) {
        const detectedVat = normalizeVatRateValue(extracted.vat_rate);
        if (detectedVat !== null) {
          item.vat = detectedVat;
        } else if (!item.vat) {
          item.vat = "21";
        }
      }
      if (
        !item.touched.vatAmount &&
        extracted.vat_amount !== null &&
        extracted.vat_amount !== undefined
      ) {
        item.vatAmount = String(extracted.vat_amount);
      }
      if (!item.touched.total && extracted.total_amount !== null && extracted.total_amount !== undefined) {
        item.total = String(extracted.total_amount);
      }

      const normalizedAmounts = normalizeInvoiceAmounts(item);
      if (!item.touched.base && normalizedAmounts.base) {
        item.base = normalizedAmounts.base;
      }
      if (!item.touched.vatAmount && normalizedAmounts.vatAmount) {
        item.vatAmount = normalizedAmounts.vatAmount;
      }
      if (!item.touched.total && normalizedAmounts.total) {
        item.total = normalizedAmounts.total;
      }

      item.analysisPending = false;
      renderIncomeTable();
    })
    .catch(() => {
      item.analysisPending = false;
      item.analysisError = true;
      item.analysisErrorMessage = ANALYSIS_ERROR_MESSAGE;
      renderIncomeTable();
    });
}

function validateIncomePending() {
  const errors = [];
  pendingIncomeFiles.forEach((item) => {
    if (item.analysisPending) {
      errors.push(`Análisis en proceso: ${item.file.name}`);
    }
    if (!item.storedFilename) {
      errors.push(`Análisis pendiente: ${item.file.name}`);
    }
    if (!item.client.trim()) {
      errors.push(`Cliente obligatorio: ${item.file.name}`);
    }
    const baseValue = parseNumberInput(item.base);
    const totalValue = parseNumberInput(item.total);
    if (baseValue === null && totalValue === null) {
      errors.push(`Base imponible o total obligatorio: ${item.file.name}`);
    }
    if (baseValue !== null && baseValue < 0) {
      errors.push(`Base imponible inválida: ${item.file.name}`);
    }
    if (totalValue !== null && totalValue < 0) {
      errors.push(`Total inválido: ${item.file.name}`);
    }
    if (!item.date) {
      errors.push(`Fecha obligatoria: ${item.file.name}`);
    }
  });
  return errors;
}

function uploadIncomePending() {
  if (pendingIncomeFiles.length === 0) {
    alert("No hay facturas emitidas para subir.");
    return;
  }
  if (!getSelectedCompanyId()) {
    alert("Selecciona una empresa antes de subir facturas emitidas.");
    return;
  }
  const errors = validateIncomePending();
  if (errors.length) {
    alert(errors.slice(0, 3).join("\n"));
    return;
  }

  if (incomeUploadBtn) {
    incomeUploadBtn.disabled = true;
  }
  const payload = {
    companyId: getSelectedCompanyId(),
    entries: pendingIncomeFiles.map((item) => {
      const normalized = normalizeInvoiceAmounts(item);
      const breakdownPayload = buildVatBreakdownPayload(item.vatBreakdown || []);
      const breakdownTotals = breakdownPayload.length
        ? summarizeVatBreakdown(item.vatBreakdown || [])
        : null;
      return {
        storedFilename: item.storedFilename,
        originalFilename: item.originalFilename,
        date: item.date,
        paymentDate: computePaymentDate(item.date, item.paymentDate),
        paymentDates: item.paymentDates || [],
        analysisStatus: item.analysisStatus || "ok",
        client: item.client.trim(),
        base: breakdownTotals ? breakdownTotals.base : normalized.base || item.base,
        vat: breakdownPayload.length
          ? getPrimaryVatRateFromBreakdown(breakdownPayload)
          : resolveVatRateValue(item.vat),
        vatAmount: breakdownTotals ? breakdownTotals.vatAmount : normalized.vatAmount || item.vatAmount,
        total: breakdownTotals ? breakdownTotals.total : normalized.total || item.total,
        vatBreakdown: breakdownPayload,
        analysisText: item.analysisText,
        companyId: getSelectedCompanyId(),
      };
    }),
  };

  fetch("/api/income-invoices", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al guardar."]).join("\n"));
        return;
      }
      pendingIncomeFiles = [];
      renderIncomeTable();
      refreshIncomeInvoices();
      refreshPayments();
    })
    .catch(() => {
      alert("No se pudo subir la factura emitida.");
    })
    .finally(() => {
      if (incomeUploadBtn) {
        incomeUploadBtn.disabled = false;
      }
    });
}

function validatePending() {
  const errors = [];
  pendingFiles.forEach((item) => {
    if (item.analysisPending) {
      errors.push(`Análisis en proceso: ${item.file.name}`);
    }
    if (!item.storedFilename) {
      errors.push(`Análisis pendiente: ${item.file.name}`);
    }
    if (!item.supplier.trim()) {
      errors.push(`Proveedor obligatorio: ${item.file.name}`);
    }
    if (isSupplierSameAsCompany(item.supplier)) {
      errors.push(`El proveedor no puede ser la empresa activa: ${item.file.name}`);
    }
    const baseValue = parseNumberInput(item.base);
    const totalValue = parseNumberInput(item.total);
    if (baseValue === null && totalValue === null) {
      errors.push(`Base imponible o total obligatorio: ${item.file.name}`);
    }
    if (baseValue !== null && baseValue < 0) {
      errors.push(`Base imponible inválida: ${item.file.name}`);
    }
    if (totalValue !== null && totalValue < 0) {
      errors.push(`Total inválido: ${item.file.name}`);
    }
    if (!item.date) {
      errors.push(`Fecha obligatoria: ${item.file.name}`);
    }
  });
  return errors;
}

function analyzeInvoiceForItem(item) {
  const formData = new FormData();
  formData.append("file", item.file);
  const companyId = getSelectedCompanyId();
  if (companyId) {
    formData.append("company_id", companyId);
  }

  fetch("/api/analyze-invoice", {
    method: "POST",
    body: formData,
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        item.analysisPending = false;
        item.analysisError = true;
        item.analysisErrorMessage = ANALYSIS_ERROR_MESSAGE;
        renderTable();
        return;
      }
      const extracted = data.extracted || {};
      item.storedFilename = data.storedFilename || "";
      item.analysisText = extracted.analysis_text || "";
      item.analysisStatus = extracted.analysis_status || "ok";
      const extractedBreakdown = parseVatBreakdown(
        extracted.vat_breakdown || extracted.vatBreakdown
      );
      if (extractedBreakdown.length) {
        item.vatBreakdown = extractedBreakdown;
        item.vatBreakdownOpen = extractedBreakdown.length > 1;
      }
      if (extracted.analysis_status === "low_quality_scan") {
        item.analysisPending = false;
        item.analysisError = true;
        item.analysisErrorMessage = LOW_QUALITY_SCAN_MESSAGE;
        if (!lowQualityDismissedIds.has(item.id)) {
          showLowQualityModal();
          lowQualityDismissedIds.add(item.id);
        }
        renderTable();
        return;
      }

      if (!item.touched.supplier && extracted.provider_name) {
        if (!isSupplierSameAsCompany(extracted.provider_name)) {
          item.supplier = extracted.provider_name;
        } else {
          item.supplier = "";
        }
      }
      if (!item.touched.date && extracted.invoice_date) {
        item.date = extracted.invoice_date;
      }
      if (Array.isArray(extracted.payment_dates) && extracted.payment_dates.length) {
        item.paymentDates = extracted.payment_dates.slice();
      }
      if (!item.paymentDate) {
        const primaryDate = item.paymentDates[0] || extracted.payment_date;
        item.paymentDate = computePaymentDate(item.date, primaryDate);
      }
      if (!item.touched.base && extracted.base_amount !== null && extracted.base_amount !== undefined) {
        item.base = String(extracted.base_amount);
      }
      if (!item.touched.vat) {
        const detectedVat = normalizeVatRateValue(extracted.vat_rate);
        if (detectedVat !== null) {
          item.vat = detectedVat;
        } else if (!item.vat) {
          item.vat = "21";
        }
      }
      if (
        !item.touched.vatAmount &&
        extracted.vat_amount !== null &&
        extracted.vat_amount !== undefined
      ) {
        item.vatAmount = String(extracted.vat_amount);
      }
      if (!item.touched.total && extracted.total_amount !== null && extracted.total_amount !== undefined) {
        item.total = String(extracted.total_amount);
      }

      const normalizedAmounts = normalizeInvoiceAmounts(item);
      if (!item.touched.base && normalizedAmounts.base) {
        item.base = normalizedAmounts.base;
      }
      if (!item.touched.vatAmount && normalizedAmounts.vatAmount) {
        item.vatAmount = normalizedAmounts.vatAmount;
      }
      if (!item.touched.total && normalizedAmounts.total) {
        item.total = normalizedAmounts.total;
      }

      item.analysisPending = false;
      const hasExtractedValue = [
        extracted.provider_name,
        extracted.invoice_date,
        extracted.base_amount,
        extracted.vat_rate,
        extracted.vat_amount,
        extracted.total_amount,
        extracted.vat_breakdown,
      ].some((value) => value !== null && value !== undefined && value !== "");
      item.analysisError = !hasExtractedValue && !item.analysisText;
      item.analysisErrorMessage = item.analysisError ? ANALYSIS_ERROR_MESSAGE : "";
      renderTable();
    })
    .catch(() => {
      item.analysisPending = false;
      item.analysisError = true;
      item.analysisErrorMessage = ANALYSIS_ERROR_MESSAGE;
      renderTable();
    });
}

function uploadPending() {
  if (pendingFiles.length === 0) {
    alert("No hay facturas para subir.");
    return;
  }
  if (!getSelectedCompanyId()) {
    alert("Selecciona una empresa antes de subir facturas.");
    return;
  }

  const errors = validatePending();
  if (errors.length) {
    alert(errors.slice(0, 3).join("\n"));
    return;
  }

  uploadBtn.disabled = true;
  const payload = {
    companyId: getSelectedCompanyId(),
    entries: pendingFiles.map((item) => {
      const normalized = normalizeInvoiceAmounts(item);
      const breakdownPayload = buildVatBreakdownPayload(item.vatBreakdown || []);
      const breakdownTotals = breakdownPayload.length
        ? summarizeVatBreakdown(item.vatBreakdown || [])
        : null;
      return {
        storedFilename: item.storedFilename,
        originalFilename: item.originalFilename,
        date: item.date,
        paymentDate: computePaymentDate(item.date, item.paymentDate),
        paymentDates: item.paymentDates || [],
        analysisStatus: item.analysisStatus || "ok",
        companyId: getSelectedCompanyId(),
        supplier: item.supplier.trim(),
        base: breakdownTotals ? breakdownTotals.base : normalized.base || item.base,
        vat: breakdownPayload.length
          ? getPrimaryVatRateFromBreakdown(breakdownPayload)
          : resolveVatRateValue(item.vat),
        vatAmount: breakdownTotals ? breakdownTotals.vatAmount : normalized.vatAmount || item.vatAmount,
        total: breakdownTotals ? breakdownTotals.total : normalized.total || item.total,
        vatBreakdown: breakdownPayload,
        analysisText: item.analysisText,
      };
    }),
  };

  fetch("/api/upload", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al subir."]).join("\n"));
        return;
      }
      if (data.errors && data.errors.length) {
        alert(data.errors.join("\n"));
      }
      pendingFiles.length = 0;
      renderTable();
      return loadYears().then(() => {
        refreshAllData();
      });
    })
    .catch(() => {
      alert("No se pudo subir la factura.");
    })
    .finally(() => {
      uploadBtn.disabled = false;
    });
}

function getPeriodMonths() {
  const { month } = getSelectedMonthYear();
  if (!month) {
    return [];
  }
  if (getSelectedPeriod() === "quarterly") {
    return getQuarterMonths(month);
  }
  return [month];
}

function fetchSummary(month, year) {
  const companyId = getSelectedCompanyId();
  const suffix = companyId ? `&company_id=${companyId}` : "";
  return fetch(`/api/summary?month=${month}&year=${year}${suffix}`).then((res) => res.json());
}

function mergeSummaries(summaries, months) {
  const vatTotals = { "0": 0, "4": 0, "10": 0, "21": 0 };
  const supplierTotals = {};
  let totalSpent = 0;
  const monthlyTotals = [];

  summaries.forEach((summary, index) => {
    const month = months[index];
    const monthTotal = Number(summary.totalSpent) || 0;
    totalSpent += monthTotal;
    monthlyTotals.push({ month, total: monthTotal });

    ["0", "4", "10", "21"].forEach((rate) => {
      vatTotals[rate] += Number(summary.vatTotals?.[rate]) || 0;
    });

    (summary.suppliers || []).forEach((supplier, supplierIndex) => {
      const value = Number(summary.supplierTotals?.[supplierIndex]) || 0;
      supplierTotals[supplier] = (supplierTotals[supplier] || 0) + value;
    });
  });

  const suppliers = Object.keys(supplierTotals);
  const supplierValues = suppliers.map((name) => Number(supplierTotals[name].toFixed(2)));
  const vatTotalDeductible = Number(
    (vatTotals["0"] + vatTotals["4"] + vatTotals["10"] + vatTotals["21"]).toFixed(2)
  );

  return {
    days: summaries[0]?.days || [],
    cumulative: summaries[0]?.cumulative || [],
    suppliers,
    supplierTotals: supplierValues,
    totalSpent: Number(totalSpent.toFixed(2)),
    vatTotals: {
      "0": Number(vatTotals["0"].toFixed(2)),
      "4": Number(vatTotals["4"].toFixed(2)),
      "10": Number(vatTotals["10"].toFixed(2)),
      "21": Number(vatTotals["21"].toFixed(2)),
    },
    vatTotalDeductible,
    monthlyTotals,
  };
}

function refreshSummary() {
  const { month, year } = getSelectedMonthYear();
  if (!month || !year) {
    return Promise.resolve();
  }
  const months = getPeriodMonths();

  return Promise.all(months.map((targetMonth) => fetchSummary(targetMonth, year)))
    .then((summaries) => {
      const merged = mergeSummaries(summaries, months);
      currentSummary = merged;
      updateSummary(merged);
      updateCharts(merged, getSelectedPeriod());
      updateNetChart();
      updateTaxSummary();
      updatePnlSummary();
      updatePeriodBadge();
    });
}

function updateSummary(data) {
  document.getElementById("totalSpent").textContent = formatCurrency(
    data.totalSpent || 0
  );
  document.getElementById("vat0").textContent = formatCurrency(
    data.vatTotals["0"] || 0
  );
  document.getElementById("vat4").textContent = formatCurrency(
    data.vatTotals["4"] || 0
  );
  document.getElementById("vat10").textContent = formatCurrency(
    data.vatTotals["10"] || 0
  );
  document.getElementById("vat21").textContent = formatCurrency(
    data.vatTotals["21"] || 0
  );
  document.getElementById("vatTotal").textContent = formatCurrency(
    data.vatTotalDeductible || 0
  );

  expenseVatTotal =
    (Number(data.vatTotals["0"]) || 0) +
    (Number(data.vatTotals["4"]) || 0) +
    (Number(data.vatTotals["10"]) || 0) +
    (Number(data.vatTotals["21"]) || 0);
  updateVatResult();
}

function updateCharts(data, period) {
  if (period === "quarterly") {
    const labels = data.monthlyTotals.map((item) => monthNames[item.month - 1]);
    const values = data.monthlyTotals.map((item) => item.total);
    updateLineChart(labels, values, "Total por mes");
  } else {
    updateLineChart(data.days, data.cumulative, "Gasto acumulado");
  }
  updatePieChart(data.suppliers, data.supplierTotals);
}

function toggleChartEmpty(id, isEmpty) {
  const node = document.getElementById(id);
  if (!node) {
    return;
  }
  node.classList.toggle("is-visible", isEmpty);
}

function updateLineChart(labels, values, datasetLabel) {
  const ctx = document.getElementById("lineChart");
  const hasData = Array.isArray(values) && values.some((value) => Number(value) > 0);
  toggleChartEmpty("lineChartEmpty", !hasData);
  if (!lineChart) {
    lineChart = new Chart(ctx, {
      type: "line",
      data: {
        labels,
        datasets: [
          {
            label: datasetLabel,
            data: values,
            borderColor: "#227c65",
            backgroundColor: "rgba(34, 124, 101, 0.15)",
            fill: true,
            tension: 0.35,
            pointRadius: 0,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: false,
          },
          tooltip: {
            callbacks: {
              label: (context) =>
                `${context.label} · ${formatCurrency(context.parsed.y)}`,
            },
          },
        },
        scales: {
          x: {
            grid: {
              display: false,
            },
          },
          y: {
            ticks: {
              callback: (value) => `${value} €`,
            },
          },
        },
      },
    });
  } else {
    lineChart.data.labels = labels;
    lineChart.data.datasets[0].data = values;
    lineChart.data.datasets[0].label = datasetLabel;
    lineChart.update();
  }
}

function updatePieChart(labels, values) {
  const ctx = document.getElementById("pieChart");
  let chartLabels = labels;
  let chartValues = values;
  let colors = [
    "#1b5d4b",
    "#e0a458",
    "#4c7a9f",
    "#c97b63",
    "#7b9e89",
    "#5c6f90",
  ];

  const hasData = Array.isArray(values) && values.some((value) => Number(value) > 0);
  toggleChartEmpty("pieChartEmpty", !hasData);

  if (!labels || labels.length === 0 || !hasData) {
    chartLabels = ["Sin datos"];
    chartValues = [1];
    colors = ["#dce5df"];
  }

  if (!pieChart) {
    pieChart = new Chart(ctx, {
      type: "pie",
      data: {
        labels: chartLabels,
        datasets: [
          {
            data: chartValues,
            backgroundColor: colors,
            borderWidth: 0,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: "bottom",
          },
        },
      },
    });
  } else {
    pieChart.data.labels = chartLabels;
    pieChart.data.datasets[0].data = chartValues;
    pieChart.data.datasets[0].backgroundColor = colors;
    pieChart.update();
  }
}

function updateBillingChart() {
  const { month, year } = getSelectedMonthYear();
  if (!month || !year) {
    return;
  }
  const months = getPeriodMonths();
  if (!months.length) {
    return;
  }
  const period = getSelectedPeriod();
  const start = new Date(year, months[0] - 1, 1);
  const end = new Date(year, months[months.length - 1], 0);
  const dayTotals = {};

  const addRecord = (dateValue, amount) => {
    if (!dateValue) {
      return;
    }
    const dateObj = new Date(`${dateValue}T00:00:00`);
    if (Number.isNaN(dateObj.getTime())) {
      return;
    }
    if (dateObj < start || dateObj > end) {
      return;
    }
    const key = dateObj.toISOString().slice(0, 10);
    dayTotals[key] = (dayTotals[key] || 0) + (Number(amount) || 0);
  };

  currentBillingEntries.forEach((entry) => {
    const dateValue =
      entry.invoice_date ||
      `${entry.year}-${String(entry.month).padStart(2, "0")}-01`;
    addRecord(dateValue, entry.base);
  });

  currentIncomeInvoices.forEach((invoice) => {
    addRecord(invoice.invoice_date, invoice.base_amount);
  });

  const labels = [];
  const values = [];
  let cumulative = 0;
  let cursor = new Date(start);
  while (cursor <= end) {
    const key = cursor.toISOString().slice(0, 10);
    const daily = Number((dayTotals[key] || 0).toFixed(2));
    cumulative = Number((cumulative + daily).toFixed(2));
    values.push(cumulative);
    if (period === "quarterly") {
      labels.push(`${cursor.getDate()} ${monthNames[cursor.getMonth()].slice(0, 3)}`);
    } else {
      labels.push(String(cursor.getDate()));
    }
    cursor.setDate(cursor.getDate() + 1);
  }

  const hasData = values.some((value) => value > 0);
  toggleChartEmpty("billingChartEmpty", !hasData);

  const ctx = document.getElementById("billingLineChart");
  if (!billingLineChart) {
    billingLineChart = new Chart(ctx, {
      type: "line",
      data: {
        labels,
        datasets: [
          {
            label: "Ingresos acumulados",
            data: values,
            borderColor: "#1b5d4b",
            backgroundColor: "rgba(27, 93, 75, 0.12)",
            fill: true,
            tension: 0.35,
            pointRadius: 0,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: false,
          },
          tooltip: {
            callbacks: {
              label: (context) =>
                `${context.label} · Acumulado: ${formatCurrency(context.parsed.y)}`,
            },
          },
        },
        scales: {
          x: {
            grid: {
              display: false,
            },
          },
          y: {
            ticks: {
              callback: (value) => `${value} €`,
            },
          },
        },
      },
    });
  } else {
    billingLineChart.data.labels = labels;
    billingLineChart.data.datasets[0].data = values;
    billingLineChart.data.datasets[0].label = "Ingresos acumulados";
    billingLineChart.update();
  }
}

function updateNetChart() {
  if (!currentSummary || !currentBillingSummary) {
    return;
  }
  const period = getSelectedPeriod();
  let labels = [];
  let values = [];
  if (period === "quarterly") {
    const months = getPeriodMonths();
    const expensesMap = {};
    months.forEach((month) => {
      expensesMap[month] = 0;
    });
    currentInvoices.forEach((invoice) => {
      if (invoice.expense_category === "non_deductible") {
        return;
      }
      const month = Number(String(invoice.invoice_date || "").slice(5, 7));
      if (!expensesMap[month]) {
        expensesMap[month] = 0;
      }
      expensesMap[month] += Number(invoice.base_amount) || 0;
    });
    currentNoInvoiceExpenses.forEach((expense) => {
      if (!expense.deductible) {
        return;
      }
      const month = Number(String(expense.expense_date || "").slice(5, 7));
      if (!expensesMap[month]) {
        expensesMap[month] = 0;
      }
      expensesMap[month] += Number(expense.amount) || 0;
    });

    labels = months.map((month) => monthNames[month - 1]);
    values = months.map((month) => {
      const income = currentBillingSummary.monthlyTotals.find(
        (entry) => entry.month === month
      )?.total || 0;
      const expenses = expensesMap[month] || 0;
      return income - expenses;
    });
  } else {
    const netValue = billingBaseTotal - currentDeductibleExpenses;
    const { month } = getSelectedMonthYear();
    labels = month ? [monthNames[month - 1]] : [];
    values = [netValue];
  }

  const hasData = values.some((value) => Number(value) !== 0);
  toggleChartEmpty("netChartEmpty", !hasData);

  const ctx = document.getElementById("netChart");
  if (!netChart) {
    netChart = new Chart(ctx, {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: "Resultado neto",
            data: values,
            backgroundColor: "rgba(76, 122, 159, 0.2)",
            borderColor: "#4c7a9f",
            borderWidth: 1,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: false,
          },
        },
        scales: {
          x: {
            grid: {
              display: false,
            },
          },
          y: {
            ticks: {
              callback: (value) => `${value} €`,
            },
          },
        },
      },
    });
  } else {
    netChart.data.labels = labels;
    netChart.data.datasets[0].data = values;
    netChart.update();
  }
}

function fetchBillingSummary(month, year) {
  const companyId = getSelectedCompanyId();
  const suffix = companyId ? `&company_id=${companyId}` : "";
  return fetch(`/api/billing/summary?month=${month}&year=${year}${suffix}`).then((res) => res.json());
}

function mergeBillingSummaries(summaries, months) {
  const baseTotals = { "0": 0, "4": 0, "10": 0, "21": 0 };
  const vatTotals = { "0": 0, "4": 0, "10": 0, "21": 0 };
  const monthlyTotals = [];

  summaries.forEach((summary, index) => {
    const month = months[index];
    let monthBaseTotal = 0;
    ["0", "4", "10", "21"].forEach((rate) => {
      const baseValue = Number(summary.baseTotals?.[rate]) || 0;
      const vatValue = Number(summary.vatTotals?.[rate]) || 0;
      baseTotals[rate] += baseValue;
      vatTotals[rate] += vatValue;
      monthBaseTotal += baseValue;
    });
    monthlyTotals.push({ month, total: monthBaseTotal });
  });

  const totalVat = Number(
    (vatTotals["0"] + vatTotals["4"] + vatTotals["10"] + vatTotals["21"]).toFixed(2)
  );

  return {
    baseTotals: {
      "0": Number(baseTotals["0"].toFixed(2)),
      "4": Number(baseTotals["4"].toFixed(2)),
      "10": Number(baseTotals["10"].toFixed(2)),
      "21": Number(baseTotals["21"].toFixed(2)),
    },
    vatTotals: {
      "0": Number(vatTotals["0"].toFixed(2)),
      "4": Number(vatTotals["4"].toFixed(2)),
      "10": Number(vatTotals["10"].toFixed(2)),
      "21": Number(vatTotals["21"].toFixed(2)),
    },
    totalVat,
    monthlyTotals,
  };
}

function refreshBillingSummary() {
  const { month, year } = getSelectedMonthYear();
  if (!month || !year) {
    return Promise.resolve();
  }
  const months = getPeriodMonths();

  return Promise.all(months.map((targetMonth) => fetchBillingSummary(targetMonth, year)))
    .then((summaries) => {
      const merged = mergeBillingSummaries(summaries, months);
      currentBillingSummary = merged;
      updateBillingSummary(merged);
      updateBillingChart();
      updateNetChart();
      updatePnlSummary();
    });
}

function updateBillingSummary(data) {
  const baseTotals = data.baseTotals || {};
  const vatTotals = data.vatTotals || {};

  document.getElementById("billingBase0").textContent = formatCurrency(
    baseTotals["0"] || 0
  );
  document.getElementById("billingBase4").textContent = formatCurrency(
    baseTotals["4"] || 0
  );
  document.getElementById("billingBase10").textContent = formatCurrency(
    baseTotals["10"] || 0
  );
  document.getElementById("billingBase21").textContent = formatCurrency(
    baseTotals["21"] || 0
  );

  document.getElementById("billingVat0").textContent = formatCurrency(
    vatTotals["0"] || 0
  );
  document.getElementById("billingVat4").textContent = formatCurrency(
    vatTotals["4"] || 0
  );
  document.getElementById("billingVat10").textContent = formatCurrency(
    vatTotals["10"] || 0
  );
  document.getElementById("billingVat21").textContent = formatCurrency(
    vatTotals["21"] || 0
  );

  billingBaseTotal =
    (Number(baseTotals["0"]) || 0) +
    (Number(baseTotals["4"]) || 0) +
    (Number(baseTotals["10"]) || 0) +
    (Number(baseTotals["21"]) || 0);
  billingVatTotal = Number(data.totalVat) || 0;
  updateVatResult();
  updateTaxSummary();
}

function renderPaymentCalendar(month, year, data) {
  if (!paymentCalendar) {
    return;
  }
  paymentCalendar.innerHTML = "";
  if (paymentCalendarTitle) {
    paymentCalendarTitle.textContent = formatMonthYear(month, year);
  }
  const monthTotal = Object.values(data.dayTotals || {}).reduce(
    (sum, value) => sum + (Number(value) || 0),
    0
  );
  if (paymentMonthTotal) {
    paymentMonthTotal.textContent = `Total a pagar del mes: ${formatCurrency(monthTotal)}`;
  }
  if (paymentMonthEmpty) {
    paymentMonthEmpty.style.display = monthTotal > 0 ? "none" : "block";
  }

  const dayNames = ["L", "M", "X", "J", "V", "S", "D"];
  const headerRow = document.createElement("div");
  headerRow.className = "calendar-row calendar-header";
  dayNames.forEach((label) => {
    const cell = document.createElement("div");
    cell.className = "calendar-cell header";
    cell.textContent = label;
    headerRow.appendChild(cell);
  });
  paymentCalendar.appendChild(headerRow);

  const firstDay = new Date(year, month - 1, 1);
  const offset = (firstDay.getDay() + 6) % 7;
  const daysInMonth = new Date(year, month, 0).getDate();
  if (selectedPaymentDay && selectedPaymentDay > daysInMonth) {
    selectedPaymentDay = null;
  }
  const grid = document.createElement("div");
  grid.className = "calendar-grid-body";

  for (let i = 0; i < offset; i += 1) {
    const emptyCell = document.createElement("div");
    emptyCell.className = "calendar-cell empty";
    grid.appendChild(emptyCell);
  }

  const itemsByDay = {};
  (data.items || []).forEach((item) => {
    if (!item.payment_date) {
      return;
    }
    const day = Number(item.payment_date.slice(8, 10));
    if (!itemsByDay[day]) {
      itemsByDay[day] = [];
    }
    itemsByDay[day].push(item);
  });
  currentPayments = { ...data, itemsByDay };

  for (let day = 1; day <= daysInMonth; day += 1) {
    const cell = document.createElement("button");
    cell.type = "button";
    cell.className = "calendar-cell day";
    const total = Number(data.dayTotals?.[day] || 0);
    cell.innerHTML = `<span class="day-number">${day}</span><span class="day-total">${formatCurrency(
      total
    )}</span>`;
    if (total > 0) {
      cell.classList.add("has-payments");
    }
    if (selectedPaymentDay === day) {
      cell.classList.add("selected");
    }
    cell.addEventListener("click", () => {
      selectedPaymentDay = day;
      renderPaymentCalendar(month, year, data);
      renderPaymentDayDetails(day);
    });
    grid.appendChild(cell);
  }

  paymentCalendar.appendChild(grid);
  renderPaymentDayDetails(selectedPaymentDay);
}

function updatePaymentDateFromCalendar(item, newDate) {
  if (!newDate) {
    alert("Selecciona una fecha válida.");
    return Promise.resolve();
  }
  if (item.type === "no_invoice") {
    const payload = {
      expense_date: newDate,
      concept: item.concept || "Gasto sin factura",
      amount: item.total_amount ?? item.amount ?? 0,
      expense_type: item.expense_type || "otro",
      deductible: item.deductible !== undefined ? item.deductible : true,
      company_id: getSelectedCompanyId(),
    };
    const url = withCompanyParam(`/api/expenses/no-invoice/${item.id}`);
    return fetch(url, {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    })
      .then((res) => res.json())
      .then((data) => {
        if (!data.ok) {
          alert((data.errors || ["Error al actualizar."]).join("\n"));
          return;
        }
        refreshPayments();
      })
      .catch(() => {
        alert("No se pudo actualizar la fecha de vencimiento.");
      });
  }
  const existingDates = Array.isArray(item.payment_dates) && item.payment_dates.length
    ? item.payment_dates.slice()
    : item.payment_date
    ? [item.payment_date]
    : [];
  const updatedDates = existingDates.map((date) =>
    date === item.payment_date ? newDate : date
  );
  if (!updatedDates.includes(newDate)) {
    updatedDates.push(newDate);
  }
  const payload = {
    invoice_date: item.invoice_date,
    payment_date: newDate,
    payment_dates: updatedDates,
    base_amount: item.base_amount,
    vat_rate: item.vat_rate,
    vat_amount: item.vat_amount,
    total_amount: item.total_amount,
  };
  if (item.type === "income") {
    payload.client = item.counterparty || "";
  } else {
    payload.supplier = item.counterparty || "";
    payload.expense_category = item.expense_category || "with_invoice";
  }
  const url = withCompanyParam(
    item.type === "income"
      ? `/api/income-invoices/${item.id}`
      : `/api/invoices/${item.id}`
  );
  return fetch(url, {
    method: "PUT",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al actualizar."]).join("\n"));
        return;
      }
      refreshPayments();
    })
    .catch(() => {
      alert("No se pudo actualizar la fecha de vencimiento.");
    });
}

function renderPaymentDayDetails(day) {
  if (!paymentDayTitle || !paymentDayList || !paymentDayTotal) {
    return;
  }
  paymentDayList.innerHTML = "";
  if (!day || !currentPayments?.itemsByDay?.[day]) {
    paymentDayTitle.textContent = "Selecciona un día para ver el detalle";
    paymentDayTotal.textContent = "";
    return;
  }
  const items = currentPayments.itemsByDay[day];
  const { month: calendarMonthValue, year: calendarYearValue } = getCalendarMonthYear();
  const monthLabel = formatMonthYear(
    Number(calendarMonthValue || monthSelect?.value),
    Number(calendarYearValue || yearSelect?.value)
  );
  paymentDayTitle.textContent = `Pagos del ${day} ${monthLabel}`;

  let total = 0;
  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "payment-day-item";
    const supplier = document.createElement("span");
    let label = "Proveedor";
    if (item.type === "income") {
      label = "Cliente";
    } else if (item.type === "no_invoice") {
      label = "Concepto";
    }
    supplier.textContent = `${label}: ${item.counterparty || "-"}`;
    const concept = document.createElement("span");
    concept.textContent = item.concept || "Factura";
    const dateLabel = document.createElement("span");
    dateLabel.textContent = item.payment_date;
    const amount = document.createElement("span");
    const amountLabel =
      item.type === "income"
        ? "Ingreso"
        : item.type === "no_invoice"
        ? "Gasto sin factura"
        : "Gasto";
    amount.textContent = `${formatCurrency(item.amount)} (${amountLabel})`;
    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "button ghost small";
    editBtn.textContent = "Editar fecha";
    const editContainer = document.createElement("div");
    editContainer.className = "payment-edit";
    editBtn.addEventListener("click", () => {
      editContainer.innerHTML = "";
      const dateInput = document.createElement("input");
      dateInput.type = "date";
      dateInput.value = item.payment_date || "";
      const saveBtn = document.createElement("button");
      saveBtn.type = "button";
      saveBtn.className = "button primary small";
      saveBtn.textContent = "Guardar";
      const cancelBtn = document.createElement("button");
      cancelBtn.type = "button";
      cancelBtn.className = "button ghost small";
      cancelBtn.textContent = "Cancelar";
      cancelBtn.addEventListener("click", () => {
        editContainer.innerHTML = "";
      });
      saveBtn.addEventListener("click", () => {
        updatePaymentDateFromCalendar(item, dateInput.value);
      });
      editContainer.appendChild(dateInput);
      editContainer.appendChild(saveBtn);
      editContainer.appendChild(cancelBtn);
    });
    row.appendChild(supplier);
    row.appendChild(concept);
    row.appendChild(dateLabel);
    row.appendChild(amount);
    row.appendChild(editBtn);
    row.appendChild(editContainer);
    paymentDayList.appendChild(row);
    total += Number(item.amount || 0);
  });
  paymentDayTotal.textContent = `Total del día: ${formatCurrency(total)}`;
}

function fetchBillingEntries(month, year) {
  const companyId = getSelectedCompanyId();
  const suffix = companyId ? `&company_id=${companyId}` : "";
  return fetch(`/api/billing/entries?month=${month}&year=${year}${suffix}`)
    .then((res) => res.json())
    .then((data) =>
      (data.entries || []).map((entry) => ({
        ...entry,
        month,
        year,
      }))
    );
}

function refreshBillingEntries() {
  const { month, year } = getSelectedMonthYear();
  if (!month || !year) {
    return Promise.resolve();
  }
  const months = getPeriodMonths();

  return Promise.all(months.map((targetMonth) => fetchBillingEntries(targetMonth, year)))
    .then((entriesByMonth) => {
      const entries = entriesByMonth.flat();
      entries.sort((a, b) => {
        if (a.year !== b.year) {
          return b.year - a.year;
        }
        if (a.month !== b.month) {
          return b.month - a.month;
        }
        return b.id - a.id;
      });
      renderBillingEntries(entries);
    });
}

function refreshBillingData() {
  return Promise.all([refreshBillingSummary(), refreshBillingEntries()]);
}

function refreshAllData() {
  if (!getSelectedCompanyId()) {
    return Promise.resolve();
  }
  return Promise.all([
    refreshSummary(),
    refreshBillingData(),
    refreshInvoices(),
    refreshIncomeInvoices(),
    refreshPayments(),
    refreshNoInvoiceExpenses(),
    refreshAnnualTaxData(),
  ]).then(() => {
    updateDashboardEmptyState();
  });
}

function updateDashboardEmptyState() {
  const emptyNode = document.getElementById("dashboardEmptyMessage");
  if (!emptyNode) {
    return;
  }
  const hasExpenses = Array.isArray(currentInvoices) && currentInvoices.length > 0;
  const hasBilling =
    (Array.isArray(currentBillingEntries) && currentBillingEntries.length > 0) ||
    (Array.isArray(currentIncomeInvoices) && currentIncomeInvoices.length > 0);
  const hasNoInvoice =
    Array.isArray(currentNoInvoiceExpenses) && currentNoInvoiceExpenses.length > 0;
  const hasData = hasExpenses || hasBilling || hasNoInvoice;
  emptyNode.style.display = hasData ? "none" : "block";
}

function updatePeriodBadge() {
  if (!taxPeriodBadge) {
    return;
  }
  const label = getPeriodLabel();
  taxPeriodBadge.textContent = label ? `Periodo: ${label}` : "";
}

// IRPF e IS se calculan siempre a nivel anual, ignorando mes/trimestre.
function refreshAnnualTaxData() {
  const { year } = getSelectedMonthYear();
  if (!year) {
    return Promise.resolve();
  }

  const months = Array.from({ length: 12 }, (_, index) => index + 1);

  return Promise.all([
    Promise.all(months.map((targetMonth) => fetchInvoices(targetMonth, year))).then(
      (invoicesByMonth) => invoicesByMonth.flat()
    ),
    Promise.all(months.map((targetMonth) => fetchNoInvoiceExpenses(targetMonth, year))).then(
      (expensesByMonth) => expensesByMonth.flat()
    ),
    Promise.all(months.map((targetMonth) => fetchBillingSummary(targetMonth, year))).then(
      (summaries) => mergeBillingSummaries(summaries, months)
    ),
  ]).then(([invoices, expenses, billingSummary]) => {
    const baseTotals = billingSummary.baseTotals || {};
    annualBillingBaseTotal =
      (Number(baseTotals["0"]) || 0) +
      (Number(baseTotals["4"]) || 0) +
      (Number(baseTotals["10"]) || 0) +
      (Number(baseTotals["21"]) || 0);

    const annualInvoices = invoices.reduce((total, invoice) => {
      if (invoice.expense_category === "non_deductible") {
        return total;
      }
      return total + (Number(invoice.base_amount) || 0);
    }, 0);

    const annualNoInvoice = expenses.reduce((total, expense) => {
      if (!expense.deductible) {
        return total;
      }
      return total + (Number(expense.amount) || 0);
    }, 0);

    annualDeductibleExpenses = annualInvoices + annualNoInvoice;
    updateTaxSummary();
  });
}

function fetchInvoices(month, year) {
  const companyId = getSelectedCompanyId();
  const suffix = companyId ? `&company_id=${companyId}` : "";
  return fetch(`/api/invoices?month=${month}&year=${year}${suffix}`)
    .then((res) => res.json())
    .then((data) => data.invoices || []);
}

function refreshInvoices() {
  const { month, year } = getSelectedMonthYear();
  if (!month || !year) {
    return Promise.resolve();
  }
  const months = getPeriodMonths();

  return Promise.all(months.map((targetMonth) => fetchInvoices(targetMonth, year)))
    .then((invoicesByMonth) => {
      const invoices = invoicesByMonth.flat();
      invoices.sort((a, b) => b.invoice_date.localeCompare(a.invoice_date));
      renderInvoices(invoices);
    });
}

function fetchPayments(month, year) {
  const companyId = getSelectedCompanyId();
  const suffix = companyId ? `&company_id=${companyId}` : "";
  return fetch(`/api/payments?month=${month}&year=${year}${suffix}`)
    .then((res) => res.json())
    .then((data) => ({
      items: data.items || [],
      dayTotals: data.dayTotals || {},
    }));
}

function refreshPayments() {
  if (!paymentCalendar) {
    return Promise.resolve();
  }
  syncCalendarWithFilters();
  const { month, year } = getCalendarMonthYear();
  if (!month || !year) {
    return Promise.resolve();
  }
  return fetchPayments(month, year).then((data) => {
    currentPayments = data;
    renderPaymentCalendar(month, year, data);
  });
}

function renderInvoices(invoices) {
  invoicesTableBody.innerHTML = "";
  currentInvoices = invoices;
  if (!invoices.length) {
    invoicesEmpty.style.display = "block";
    updateTaxSummary();
    updateSupplierSuggestions();
    return;
  }
  invoicesEmpty.style.display = "none";

  invoices.forEach((invoice) => {
    const tr = document.createElement("tr");
    tr.dataset.id = invoice.id;

    const dateTd = document.createElement("td");
    dateTd.textContent = invoice.invoice_date;

    const supplierTd = document.createElement("td");
    supplierTd.textContent = invoice.supplier;

    const baseTd = document.createElement("td");
    baseTd.textContent = formatCurrency(invoice.base_amount);

    const vatTd = document.createElement("td");
    const vatDisplay = getVatDisplayFromInvoice(invoice);
    vatTd.textContent = vatDisplay.label;
    if (vatDisplay.title) {
      vatTd.title = vatDisplay.title;
    }

    const vatAmountTd = document.createElement("td");
    vatAmountTd.textContent = formatCurrency(invoice.vat_amount || 0);

    const totalTd = document.createElement("td");
    totalTd.textContent = formatCurrency(invoice.total_amount);

    const categoryTd = document.createElement("td");
    categoryTd.textContent = formatExpenseCategory(
      invoice.expense_category || "with_invoice"
    );

    const actionsTd = document.createElement("td");
    actionsTd.classList.add("invoice-actions");

    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "button ghost";
    editBtn.textContent = "Editar";
    editBtn.addEventListener("click", () => {
      enterInvoiceEditMode(tr, invoice);
    });

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "button danger";
    deleteBtn.textContent = "Eliminar";
    deleteBtn.addEventListener("click", () => {
      if (!confirm("¿Seguro que deseas eliminar esta factura?")) {
        return;
      }
      deleteInvoice(invoice.id);
    });

    actionsTd.appendChild(editBtn);
    actionsTd.appendChild(deleteBtn);

    tr.appendChild(dateTd);
    tr.appendChild(supplierTd);
    tr.appendChild(baseTd);
    tr.appendChild(vatTd);
    tr.appendChild(vatAmountTd);
    tr.appendChild(totalTd);
    tr.appendChild(categoryTd);
    tr.appendChild(actionsTd);
    invoicesTableBody.appendChild(tr);
  });

  updateTaxSummary();
  updateSupplierSuggestions();
}

function fetchIncomeInvoices(month, year) {
  const companyId = getSelectedCompanyId();
  const suffix = companyId ? `&company_id=${companyId}` : "";
  return fetch(`/api/income-invoices?month=${month}&year=${year}${suffix}`)
    .then((res) => res.json())
    .then((data) => data.invoices || []);
}

function refreshIncomeInvoices() {
  if (!incomeInvoicesTableBody) {
    return Promise.resolve();
  }
  const { month, year } = getSelectedMonthYear();
  if (!month || !year) {
    return Promise.resolve();
  }
  const months = getPeriodMonths();
  return Promise.all(months.map((targetMonth) => fetchIncomeInvoices(targetMonth, year)))
    .then((invoicesByMonth) => {
      const invoices = invoicesByMonth.flat();
      invoices.sort((a, b) => b.invoice_date.localeCompare(a.invoice_date));
      renderIncomeInvoices(invoices);
    });
}

function renderIncomeInvoices(invoices) {
  if (!incomeInvoicesTableBody || !incomeInvoicesEmpty) {
    return;
  }
  incomeInvoicesTableBody.innerHTML = "";
  currentIncomeInvoices = invoices;
  if (!invoices.length) {
    incomeInvoicesEmpty.style.display = "block";
    updateBillingChart();
    return;
  }
  incomeInvoicesEmpty.style.display = "none";

  invoices.forEach((invoice) => {
    const tr = document.createElement("tr");
    tr.dataset.id = invoice.id;

    const dateTd = document.createElement("td");
    dateTd.textContent = invoice.invoice_date;

    const clientTd = document.createElement("td");
    clientTd.textContent = invoice.client;

    const baseTd = document.createElement("td");
    baseTd.textContent = formatCurrency(invoice.base_amount);

    const vatTd = document.createElement("td");
    const vatDisplay = getVatDisplayFromInvoice(invoice);
    vatTd.textContent = vatDisplay.label;
    if (vatDisplay.title) {
      vatTd.title = vatDisplay.title;
    }

    const vatAmountTd = document.createElement("td");
    vatAmountTd.textContent = formatCurrency(invoice.vat_amount || 0);

    const totalTd = document.createElement("td");
    totalTd.textContent = formatCurrency(invoice.total_amount);

    const actionsTd = document.createElement("td");
    actionsTd.classList.add("billing-actions");

    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "button ghost";
    editBtn.textContent = "Editar";
    editBtn.addEventListener("click", () => {
      enterIncomeInvoiceEditMode(tr, invoice);
    });

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "button danger";
    deleteBtn.textContent = "Eliminar";
    deleteBtn.addEventListener("click", () => {
      if (!confirm("¿Seguro que deseas eliminar esta factura emitida?")) {
        return;
      }
      deleteIncomeInvoice(invoice.id);
    });

    actionsTd.appendChild(editBtn);
    actionsTd.appendChild(deleteBtn);

    tr.appendChild(dateTd);
    tr.appendChild(clientTd);
    tr.appendChild(baseTd);
    tr.appendChild(vatTd);
    tr.appendChild(vatAmountTd);
    tr.appendChild(totalTd);
    tr.appendChild(actionsTd);
    incomeInvoicesTableBody.appendChild(tr);
  });
  updateBillingChart();
}

function enterIncomeInvoiceEditMode(row, invoice) {
  const dateTd = row.children[0];
  const clientTd = row.children[1];
  const baseTd = row.children[2];
  const vatTd = row.children[3];
  const vatAmountTd = row.children[4];
  const totalTd = row.children[5];
  const actionsTd = row.children[6];

  const dateInput = document.createElement("input");
  dateInput.type = "date";
  dateInput.value = invoice.invoice_date;

  const clientInput = document.createElement("input");
  clientInput.type = "text";
  clientInput.value = invoice.client;

  const baseInput = document.createElement("input");
  baseInput.type = "number";
  baseInput.step = "0.01";
  baseInput.min = "0";
  baseInput.value = invoice.base_amount;

  const vatSelect = createVatSelect(invoice.vat_rate);

  const vatAmountInput = document.createElement("input");
  vatAmountInput.type = "number";
  vatAmountInput.step = "0.01";
  vatAmountInput.min = "0";
  vatAmountInput.value = invoice.vat_amount || 0;

  const totalInput = document.createElement("input");
  totalInput.type = "number";
  totalInput.step = "0.01";
  totalInput.min = "0";
  totalInput.value = invoice.total_amount;

  const calcInputs = {
    base: baseInput,
    vat: vatSelect,
    vatAmount: vatAmountInput,
    total: totalInput,
  };
  baseInput.addEventListener("input", () => {
    applyVatCalculation(invoice, calcInputs, "base");
  });
  vatSelect.addEventListener("change", () => {
    applyVatCalculation(invoice, calcInputs, "vat");
  });
  totalInput.addEventListener("input", () => {
    applyVatCalculation(invoice, calcInputs, "total");
  });

  dateTd.textContent = "";
  dateTd.appendChild(dateInput);
  clientTd.textContent = "";
  clientTd.appendChild(clientInput);
  baseTd.textContent = "";
  baseTd.appendChild(baseInput);
  vatTd.textContent = "";
  vatTd.appendChild(vatSelect);
  vatAmountTd.textContent = "";
  vatAmountTd.appendChild(vatAmountInput);
  totalTd.textContent = "";
  totalTd.appendChild(totalInput);

  const initialSource =
    parseNumberInput(baseInput.value) !== null ? "base" : "total";
  applyVatCalculation(invoice, calcInputs, initialSource);

  actionsTd.innerHTML = "";
  const saveBtn = document.createElement("button");
  saveBtn.type = "button";
  saveBtn.className = "button primary";
  saveBtn.textContent = "Guardar";
  saveBtn.addEventListener("click", () => {
    updateIncomeInvoice(invoice.id, {
      invoice_date: dateInput.value,
      client: clientInput.value,
      base_amount: baseInput.value,
      vat_rate: vatSelect.value,
      vat_amount: vatAmountInput.value,
      total_amount: totalInput.value,
    });
  });

  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "button ghost";
  cancelBtn.textContent = "Cancelar";
  cancelBtn.addEventListener("click", () => {
    refreshIncomeInvoices();
  });

  actionsTd.appendChild(saveBtn);
  actionsTd.appendChild(cancelBtn);
}

function updateIncomeInvoice(invoiceId, payload) {
  const normalized = normalizeInvoiceAmounts({
    base: payload.base_amount,
    total: payload.total_amount,
    vat: payload.vat_rate,
  });
  const normalizedPayload = {
    ...payload,
    base_amount: normalized.base || payload.base_amount,
    vat_amount: normalized.vatAmount || payload.vat_amount,
    total_amount: normalized.total || payload.total_amount,
  };
  const url = withCompanyParam(`/api/income-invoices/${invoiceId}`);
  fetch(url, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      ...normalizedPayload,
      company_id: getSelectedCompanyId(),
    }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al actualizar."]).join("\n"));
        return;
      }
      refreshIncomeInvoices();
      refreshPayments();
    })
    .catch(() => {
      alert("No se pudo actualizar la factura emitida.");
    });
}

function deleteIncomeInvoice(invoiceId) {
  const url = withCompanyParam(`/api/income-invoices/${invoiceId}`);
  fetch(url, {
    method: "DELETE",
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al eliminar."]).join("\n"));
        return;
      }
      refreshIncomeInvoices();
      refreshPayments();
    })
    .catch(() => {
      alert("No se pudo eliminar la factura emitida.");
    });
}

function enterInvoiceEditMode(row, invoice) {
  const dateTd = row.children[0];
  const supplierTd = row.children[1];
  const baseTd = row.children[2];
  const vatTd = row.children[3];
  const vatAmountTd = row.children[4];
  const totalTd = row.children[5];
  const categoryTd = row.children[6];
  const actionsTd = row.children[7];

  const dateInput = document.createElement("input");
  dateInput.type = "date";
  dateInput.value = invoice.invoice_date;

  const supplierInput = document.createElement("input");
  supplierInput.type = "text";
  supplierInput.value = invoice.supplier;
  if (supplierSuggestions) {
    supplierInput.setAttribute("list", "supplierSuggestions");
  }
  const supplierWarning = document.createElement("div");
  supplierWarning.className = "field-warning";
  const updateSupplierWarning = () => {
    const value = supplierInput.value.trim();
    if (!value) {
      supplierWarning.textContent = "Proveedor pendiente de completar.";
      supplierWarning.style.display = "block";
      supplierInput.classList.add("input-warning");
      return;
    }
    if (isSupplierSameAsCompany(value)) {
      supplierWarning.textContent =
        "El proveedor no puede ser la empresa activa.";
      supplierWarning.style.display = "block";
      supplierInput.classList.add("input-warning");
      return;
    }
    supplierWarning.textContent = "";
    supplierWarning.style.display = "none";
    supplierInput.classList.remove("input-warning");
  };
  supplierInput.addEventListener("input", updateSupplierWarning);

  const baseInput = document.createElement("input");
  baseInput.type = "number";
  baseInput.step = "0.01";
  baseInput.min = "0";
  baseInput.value = invoice.base_amount;

  const vatSelect = createVatSelect(invoice.vat_rate);

  const vatAmountInput = document.createElement("input");
  vatAmountInput.type = "number";
  vatAmountInput.step = "0.01";
  vatAmountInput.min = "0";
  vatAmountInput.value = invoice.vat_amount || 0;

  const totalInput = document.createElement("input");
  totalInput.type = "number";
  totalInput.step = "0.01";
  totalInput.min = "0";
  totalInput.value = invoice.total_amount;

  const calcInputs = {
    base: baseInput,
    vat: vatSelect,
    vatAmount: vatAmountInput,
    total: totalInput,
  };
  baseInput.addEventListener("input", () => {
    applyVatCalculation(invoice, calcInputs, "base");
  });
  vatSelect.addEventListener("change", () => {
    applyVatCalculation(invoice, calcInputs, "vat");
  });
  totalInput.addEventListener("input", () => {
    applyVatCalculation(invoice, calcInputs, "total");
  });

  const categorySelect = createExpenseCategorySelect(invoice.expense_category);

  dateTd.textContent = "";
  dateTd.appendChild(dateInput);
  supplierTd.textContent = "";
  supplierTd.appendChild(supplierInput);
  supplierTd.appendChild(supplierWarning);
  updateSupplierWarning();
  baseTd.textContent = "";
  baseTd.appendChild(baseInput);
  vatTd.textContent = "";
  vatTd.appendChild(vatSelect);
  vatAmountTd.textContent = "";
  vatAmountTd.appendChild(vatAmountInput);
  totalTd.textContent = "";
  totalTd.appendChild(totalInput);
  categoryTd.textContent = "";
  categoryTd.appendChild(categorySelect);
  const initialSource =
    parseNumberInput(baseInput.value) !== null ? "base" : "total";
  applyVatCalculation(invoice, calcInputs, initialSource);

  actionsTd.innerHTML = "";
  const saveBtn = document.createElement("button");
  saveBtn.type = "button";
  saveBtn.className = "button primary";
  saveBtn.textContent = "Guardar";
  saveBtn.addEventListener("click", () => {
    updateInvoice(invoice.id, {
      invoice_date: dateInput.value,
      supplier: supplierInput.value,
      base_amount: baseInput.value,
      vat_rate: vatSelect.value,
      vat_amount: vatAmountInput.value,
      total_amount: totalInput.value,
      expense_category: categorySelect.value,
    });
  });

  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "button ghost";
  cancelBtn.textContent = "Cancelar";
  cancelBtn.addEventListener("click", () => {
    refreshInvoices();
  });

  actionsTd.appendChild(saveBtn);
  actionsTd.appendChild(cancelBtn);
}

function updateInvoice(invoiceId, payload) {
  if (payload.supplier && isSupplierSameAsCompany(payload.supplier)) {
    alert("El proveedor no puede ser la empresa activa.");
    return;
  }
  const normalized = normalizeInvoiceAmounts({
    base: payload.base_amount,
    total: payload.total_amount,
    vat: payload.vat_rate,
  });
  const normalizedPayload = {
    ...payload,
    base_amount: normalized.base || payload.base_amount,
    vat_amount: normalized.vatAmount || payload.vat_amount,
    total_amount: normalized.total || payload.total_amount,
  };
  const url = withCompanyParam(`/api/invoices/${invoiceId}`);
  fetch(url, {
    method: "PUT",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      ...normalizedPayload,
      company_id: getSelectedCompanyId(),
    }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al actualizar."]).join("\n"));
        return;
      }
      refreshAllData();
    })
    .catch(() => {
      alert("No se pudo actualizar la factura.");
    });
}

function deleteInvoice(invoiceId) {
  const url = withCompanyParam(`/api/invoices/${invoiceId}`);
  fetch(url, {
    method: "DELETE",
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al eliminar."]).join("\n"));
        return;
      }
      refreshAllData();
    })
    .catch(() => {
      alert("No se pudo eliminar la factura.");
    });
}

function fetchNoInvoiceExpenses(month, year) {
  const companyId = getSelectedCompanyId();
  const suffix = companyId ? `&company_id=${companyId}` : "";
  return fetch(`/api/expenses/no-invoice?month=${month}&year=${year}${suffix}`)
    .then((res) => res.json())
    .then((data) => data.expenses || []);
}

function refreshNoInvoiceExpenses() {
  const { month, year } = getSelectedMonthYear();
  if (!month || !year) {
    return Promise.resolve();
  }
  const months = getPeriodMonths();

  return Promise.all(months.map((targetMonth) => fetchNoInvoiceExpenses(targetMonth, year)))
    .then((expensesByMonth) => {
      const expenses = expensesByMonth.flat();
      expenses.sort((a, b) => b.expense_date.localeCompare(a.expense_date));
      renderNoInvoiceExpenses(expenses);
    });
}

function renderNoInvoiceExpenses(expenses) {
  noInvoiceTableBody.innerHTML = "";
  currentNoInvoiceExpenses = expenses;
  if (!expenses.length) {
    noInvoiceEmpty.style.display = "block";
    updateTaxSummary();
    return;
  }
  noInvoiceEmpty.style.display = "none";

  expenses.forEach((expense) => {
    const tr = document.createElement("tr");
    tr.dataset.id = expense.id;

    const dateTd = document.createElement("td");
    dateTd.textContent = expense.expense_date;

    const conceptTd = document.createElement("td");
    conceptTd.textContent = expense.concept;

    const amountTd = document.createElement("td");
    amountTd.textContent = formatCurrency(expense.amount);

    const typeTd = document.createElement("td");
    typeTd.textContent = formatNoInvoiceType(expense.expense_type);

    const deductibleTd = document.createElement("td");
    deductibleTd.textContent = expense.deductible ? "Sí" : "No";

    const actionsTd = document.createElement("td");
    actionsTd.classList.add("billing-actions");

    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "button ghost";
    editBtn.textContent = "Editar";
    editBtn.addEventListener("click", () => {
      enterNoInvoiceEditMode(tr, expense);
    });

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "button danger";
    deleteBtn.textContent = "Eliminar";
    deleteBtn.addEventListener("click", () => {
      if (!confirm("¿Seguro que deseas eliminar este gasto?")) {
        return;
      }
      deleteNoInvoiceExpense(expense.id);
    });

    actionsTd.appendChild(editBtn);
    actionsTd.appendChild(deleteBtn);

    tr.appendChild(dateTd);
    tr.appendChild(conceptTd);
    tr.appendChild(amountTd);
    tr.appendChild(typeTd);
    tr.appendChild(deductibleTd);
    tr.appendChild(actionsTd);
    noInvoiceTableBody.appendChild(tr);
  });

  updateTaxSummary();
}

function enterNoInvoiceEditMode(row, expense) {
  const dateTd = row.children[0];
  const conceptTd = row.children[1];
  const amountTd = row.children[2];
  const typeTd = row.children[3];
  const deductibleTd = row.children[4];
  const actionsTd = row.children[5];

  const dateInput = document.createElement("input");
  dateInput.type = "date";
  dateInput.value = expense.expense_date;

  const conceptInput = document.createElement("input");
  conceptInput.type = "text";
  conceptInput.value = expense.concept;

  const amountInput = document.createElement("input");
  amountInput.type = "number";
  amountInput.step = "0.01";
  amountInput.min = "0";
  amountInput.value = expense.amount;

  const typeSelect = createNoInvoiceTypeSelect(expense.expense_type);
  const deductibleSelect = createDeductibleSelect(expense.deductible);

  dateTd.textContent = "";
  dateTd.appendChild(dateInput);
  conceptTd.textContent = "";
  conceptTd.appendChild(conceptInput);
  amountTd.textContent = "";
  amountTd.appendChild(amountInput);
  typeTd.textContent = "";
  typeTd.appendChild(typeSelect);
  deductibleTd.textContent = "";
  deductibleTd.appendChild(deductibleSelect);

  actionsTd.innerHTML = "";
  const saveBtn = document.createElement("button");
  saveBtn.type = "button";
  saveBtn.className = "button primary";
  saveBtn.textContent = "Guardar";
  saveBtn.addEventListener("click", () => {
    updateNoInvoiceExpense(expense.id, {
      expense_date: dateInput.value,
      concept: conceptInput.value,
      amount: amountInput.value,
      expense_type: typeSelect.value,
      deductible: deductibleSelect.value === "true",
    });
  });

  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "button ghost";
  cancelBtn.textContent = "Cancelar";
  cancelBtn.addEventListener("click", () => {
    refreshNoInvoiceExpenses();
  });

  actionsTd.appendChild(saveBtn);
  actionsTd.appendChild(cancelBtn);
}

function saveNoInvoiceExpense() {
  if (!getSelectedCompanyId()) {
    alert("Selecciona una empresa antes de guardar gastos.");
    return;
  }
  const dateValue = noInvoiceDate.value;
  const conceptValue = noInvoiceConcept.value.trim();
  const amountValue = noInvoiceAmount.value;
  const typeValue = noInvoiceType.value;
  const deductibleValue = noInvoiceDeductible.value === "true";

  if (!dateValue) {
    alert("Fecha obligatoria.");
    return;
  }
  if (!conceptValue) {
    alert("Concepto obligatorio.");
    return;
  }
  if (!amountValue || Number(amountValue) < 0) {
    alert("Importe inválido.");
    return;
  }

  noInvoiceSaveBtn.disabled = true;
  fetch("/api/expenses/no-invoice", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      company_id: getSelectedCompanyId(),
      expense_date: dateValue,
      concept: conceptValue,
      amount: amountValue,
      expense_type: typeValue,
      deductible: deductibleValue,
    }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al guardar."]).join("\n"));
        return;
      }
      noInvoiceConcept.value = "";
      noInvoiceAmount.value = "";
      refreshNoInvoiceExpenses();
    })
    .catch(() => {
      alert("No se pudo guardar el gasto.");
    })
    .finally(() => {
      noInvoiceSaveBtn.disabled = false;
    });
}

function updateNoInvoiceExpense(expenseId, payload) {
  const url = withCompanyParam(`/api/expenses/no-invoice/${expenseId}`);
  fetch(url, {
    method: "PUT",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      ...payload,
      company_id: getSelectedCompanyId(),
    }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al actualizar."]).join("\n"));
        return;
      }
      refreshNoInvoiceExpenses();
    })
    .catch(() => {
      alert("No se pudo actualizar el gasto.");
    });
}

function deleteNoInvoiceExpense(expenseId) {
  const url = withCompanyParam(`/api/expenses/no-invoice/${expenseId}`);
  fetch(url, {
    method: "DELETE",
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al eliminar."]).join("\n"));
        return;
      }
      refreshNoInvoiceExpenses();
    })
    .catch(() => {
      alert("No se pudo eliminar el gasto.");
    });
}

function renderBillingEntries(entries) {
  billingEntriesBody.innerHTML = "";
  currentBillingEntries = entries;
  if (!entries.length) {
    billingEntriesEmpty.style.display = "block";
    updateBillingChart();
    return;
  }
  billingEntriesEmpty.style.display = "none";

  const showMonth = getSelectedPeriod() === "quarterly";
  entries.forEach((entry) => {
    const tr = document.createElement("tr");
    tr.dataset.id = entry.id;

    if (showMonth) {
      const monthTd = document.createElement("td");
      monthTd.classList.add("period-month");
      monthTd.textContent = monthNames[Number(entry.month) - 1] || "";
      tr.appendChild(monthTd);
    }

    const baseTd = document.createElement("td");
    baseTd.textContent = formatCurrency(entry.base);

    const vatTd = document.createElement("td");
    vatTd.textContent = `${entry.vat}%`;

    const vatAmountTd = document.createElement("td");
    vatAmountTd.textContent = formatCurrency(entry.vatAmount);

    const actionsTd = document.createElement("td");
    actionsTd.classList.add("billing-actions");

    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "button ghost";
    editBtn.textContent = "Editar";
    editBtn.addEventListener("click", () => {
      enterEditMode(tr, entry);
    });

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "button danger";
    deleteBtn.textContent = "Eliminar";
    deleteBtn.addEventListener("click", () => {
      deleteBillingEntry(entry.id);
    });

    actionsTd.appendChild(editBtn);
    actionsTd.appendChild(deleteBtn);

    tr.appendChild(baseTd);
    tr.appendChild(vatTd);
    tr.appendChild(vatAmountTd);
    tr.appendChild(actionsTd);
    billingEntriesBody.appendChild(tr);
  });
  updateBillingChart();
}

function enterEditMode(row, entry) {
  const showMonth = getSelectedPeriod() === "quarterly";
  const baseIndex = showMonth ? 1 : 0;
  const vatIndex = showMonth ? 2 : 1;
  const actionsIndex = showMonth ? 4 : 3;

  const baseTd = row.children[baseIndex];
  const vatTd = row.children[vatIndex];
  const actionsTd = row.children[actionsIndex];

  const baseInput = document.createElement("input");
  baseInput.type = "number";
  baseInput.step = "0.01";
  baseInput.min = "0";
  baseInput.value = Number(entry.base).toFixed(2);

  const vatSelect = createVatSelect(entry.vat);

  baseTd.textContent = "";
  baseTd.appendChild(baseInput);
  vatTd.textContent = "";
  vatTd.appendChild(vatSelect);

  actionsTd.innerHTML = "";
  const saveBtn = document.createElement("button");
  saveBtn.type = "button";
  saveBtn.className = "button primary";
  saveBtn.textContent = "Guardar";
  saveBtn.addEventListener("click", () => {
    updateBillingEntry(entry.id, baseInput.value, vatSelect.value);
  });

  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "button ghost";
  cancelBtn.textContent = "Cancelar";
  cancelBtn.addEventListener("click", () => {
    refreshBillingEntries();
  });

  actionsTd.appendChild(saveBtn);
  actionsTd.appendChild(cancelBtn);
}

function updateBillingEntry(entryId, baseValue, vatValue) {
  if (!baseValue || Number(baseValue) < 0) {
    alert("Base facturada inválida.");
    return;
  }

  const url = withCompanyParam(`/api/billing/${entryId}`);
  fetch(url, {
    method: "PUT",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      company_id: getSelectedCompanyId(),
      base: baseValue,
      vat: vatValue,
    }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al actualizar."]).join("\n"));
        return;
      }
      refreshBillingData();
    })
    .catch(() => {
      alert("No se pudo actualizar la facturación.");
    });
}

function deleteBillingEntry(entryId) {
  const url = withCompanyParam(`/api/billing/${entryId}`);
  fetch(url, {
    method: "DELETE",
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al eliminar."]).join("\n"));
        return;
      }
      refreshBillingData();
    })
    .catch(() => {
      alert("No se pudo eliminar la facturación.");
    });
}

function updateVatResult() {
  const result = billingVatTotal - expenseVatTotal;
  document.getElementById("vatOutputTotal").textContent = formatCurrency(
    billingVatTotal
  );
  document.getElementById("vatInputTotal").textContent = formatCurrency(
    expenseVatTotal
  );
  document.getElementById("vatResultLabel").textContent =
    result >= 0 ? "IVA A PAGAR" : "IVA A DEVOLVER";
  document.getElementById("vatResultValue").textContent = formatCurrency(
    Math.abs(result)
  );
}

function updateTaxSummary() {
  const deductibleInvoices = currentInvoices.reduce((total, invoice) => {
    if (invoice.expense_category === "non_deductible") {
      return total;
    }
    return total + (Number(invoice.base_amount) || 0);
  }, 0);

  const deductibleNoInvoice = currentNoInvoiceExpenses.reduce((total, expense) => {
    if (!expense.deductible) {
      return total;
    }
    return total + (Number(expense.amount) || 0);
  }, 0);

  const periodExpenses = deductibleInvoices + deductibleNoInvoice;
  currentDeductibleExpenses = periodExpenses;

  const annualIncome = annualBillingBaseTotal;
  const annualExpenses = annualDeductibleExpenses;
  const annualNet = annualIncome - annualExpenses;

  document.getElementById("irpfIncome").textContent = formatCurrency(annualIncome);
  document.getElementById("irpfExpenses").textContent = formatCurrency(annualExpenses);
  document.getElementById("irpfNet").textContent = formatCurrency(annualNet);

  document.getElementById("isIncome").textContent = formatCurrency(annualIncome);
  document.getElementById("isExpenses").textContent = formatCurrency(annualExpenses);
  document.getElementById("isResult").textContent = formatCurrency(annualNet);
  document.getElementById("isBase").textContent = formatCurrency(annualNet);

  updatePnlSummary();
  updateNetChart();
}

function updatePnlSummary() {
  const incomeTotal = billingBaseTotal;
  const expensesTotal = currentDeductibleExpenses;
  const operatingResultEl = document.getElementById("pnlOperatingResult");
  const financialResultEl = document.getElementById("pnlFinancialResult");
  const preTaxEl = document.getElementById("pnlPreTax");
  const netEl = document.getElementById("pnlNet");
  if (!operatingResultEl || !financialResultEl || !preTaxEl || !netEl) {
    return;
  }

  setPnlInputValue("pnlLine1", incomeTotal, true);
  setPnlInputValue("pnlLine4", expensesTotal, true);

  const opIncome =
    getPnlInputValue("pnlLine1") +
    getPnlInputValue("pnlLine2") +
    getPnlInputValue("pnlLine3") +
    getPnlInputValue("pnlLine5") +
    getPnlInputValue("pnlLine9") +
    getPnlInputValue("pnlLine10") +
    getPnlInputValue("pnlLine11") +
    getPnlInputValue("pnlLine12");
  const opExpenses =
    getPnlInputValue("pnlLine4") +
    getPnlInputValue("pnlLine6") +
    getPnlInputValue("pnlLine7") +
    getPnlInputValue("pnlLine8");
  const operatingResult = opIncome - opExpenses;

  const financialIncome =
    getPnlInputValue("pnlLine13a") +
    getPnlInputValue("pnlLine13b") +
    getPnlInputValue("pnlLine18a") +
    getPnlInputValue("pnlLine18b") +
    getPnlInputValue("pnlLine18c");
  const financialExpenses =
    getPnlInputValue("pnlLine14") +
    getPnlInputValue("pnlLine15") +
    getPnlInputValue("pnlLine16") +
    getPnlInputValue("pnlLine17");
  const financialResult = financialIncome - financialExpenses;
  const preTax = operatingResult + financialResult;

  const companyType = getSelectedCompanyType();
  const taxRate =
    companyType === "company" ? 0.25 : companyType === "individual" ? 0.15 : 0;
  const defaultTaxes = preTax > 0 ? preTax * taxRate : 0;
  setPnlInputValue("pnlLine19", defaultTaxes, true);
  const taxes = getPnlInputValue("pnlLine19");
  const netResult = preTax - taxes;

  operatingResultEl.textContent = formatCurrency(operatingResult);
  financialResultEl.textContent = formatCurrency(financialResult);
  preTaxEl.textContent = formatCurrency(preTax);
  netEl.textContent = formatCurrency(netResult);
}

function exportPnlPdf() {
  const { jsPDF } = window.jspdf || {};
  if (!jsPDF) {
    alert("No se pudo cargar el módulo de exportación PDF.");
    return;
  }

  const nameValue = pnlName.value.trim();
  const taxIdValue = pnlTaxId.value.trim();
  const periodLabel = getPeriodLabel();

  const incomeValue = formatCurrency(getPnlInputValue("pnlLine1"));
  const expensesValue = formatCurrency(getPnlInputValue("pnlLine4"));
  const preTaxValue = document.getElementById("pnlPreTax").textContent;
  const taxesValue = formatCurrency(getPnlInputValue("pnlLine19"));
  const netValue = document.getElementById("pnlNet").textContent;

  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const margin = 48;
  let y = margin;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.text("Cuenta de pérdidas y ganancias (estimada)", margin, y);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(11);
  y += 24;
  if (nameValue) {
    doc.text(`Nombre: ${nameValue}`, margin, y);
    y += 16;
  }
  if (taxIdValue) {
    doc.text(`CIF/NIF: ${taxIdValue}`, margin, y);
    y += 16;
  }
  if (periodLabel) {
    doc.text(`Periodo: ${periodLabel}`, margin, y);
    y += 16;
  }

  const dateLabel = new Date().toLocaleDateString("es-ES");
  doc.text(`Fecha de generación: ${dateLabel}`, margin, y);
  y += 24;

  doc.setFont("helvetica", "bold");
  doc.text("Resumen", margin, y);
  y += 16;

  doc.setFont("helvetica", "normal");
  const rows = [
    ["1. Importe neto de la cifra de negocios", incomeValue],
    ["2. Variación de existencias de productos terminados y en curso", formatCurrency(getPnlInputValue("pnlLine2"))],
    ["3. Trabajos realizados por la empresa para su activo", formatCurrency(getPnlInputValue("pnlLine3"))],
    ["4. Aprovisionamientos", expensesValue],
    ["5. Otros ingresos de explotación", formatCurrency(getPnlInputValue("pnlLine5"))],
    ["6. Gastos de personal", formatCurrency(getPnlInputValue("pnlLine6"))],
    ["7. Otros gastos de explotación", formatCurrency(getPnlInputValue("pnlLine7"))],
    ["8. Amortización del inmovilizado", formatCurrency(getPnlInputValue("pnlLine8"))],
    ["9. Imputación de subvenciones de inmovilizado no financiero y otras", formatCurrency(getPnlInputValue("pnlLine9"))],
    ["10. Excesos de provisiones", formatCurrency(getPnlInputValue("pnlLine10"))],
    ["11. Deterioro y resultado por enajenación del inmovilizado", formatCurrency(getPnlInputValue("pnlLine11"))],
    ["12. Otros resultados", formatCurrency(getPnlInputValue("pnlLine12"))],
    ["A) RESULTADO DE EXPLOTACIÓN", document.getElementById("pnlOperatingResult").textContent],
    ["13.a Imputación de subvenciones, donaciones y legados de carácter financiero", formatCurrency(getPnlInputValue("pnlLine13a"))],
    ["13.b Otros ingresos financieros", formatCurrency(getPnlInputValue("pnlLine13b"))],
    ["14. Gastos financieros", formatCurrency(getPnlInputValue("pnlLine14"))],
    ["15. Variación de valor razonable en instrumentos financieros", formatCurrency(getPnlInputValue("pnlLine15"))],
    ["16. Diferencias de cambio", formatCurrency(getPnlInputValue("pnlLine16"))],
    ["17. Deterioro y resultado por enajenación de instrumentos financieros", formatCurrency(getPnlInputValue("pnlLine17"))],
    ["18.a Incorporación al activo de gastos financieros", formatCurrency(getPnlInputValue("pnlLine18a"))],
    ["18.b Ingresos financieros derivados de convenios de acreedores", formatCurrency(getPnlInputValue("pnlLine18b"))],
    ["18.c Resto de ingresos y gastos", formatCurrency(getPnlInputValue("pnlLine18c"))],
    ["B) RESULTADO FINANCIERO", document.getElementById("pnlFinancialResult").textContent],
    ["C) RESULTADO ANTES DE IMPUESTOS", preTaxValue],
    ["19. Impuestos sobre beneficios", taxesValue],
    ["D) RESULTADO DEL EJERCICIO", netValue],
  ];

  rows.forEach(([label, value]) => {
    doc.text(label, margin, y);
    doc.text(value, margin + 320, y, { align: "right" });
    y += 18;
  });

  const filename = `pnl_${periodLabel || "periodo"}.pdf`.replace(/\s+/g, "_");
  doc.save(filename);
}

function populateReportMonthSelect(select) {
  if (!select) {
    return;
  }
  select.innerHTML = "";
  monthNames.forEach((name, index) => {
    const option = document.createElement("option");
    option.value = String(index + 1);
    option.textContent = name;
    select.appendChild(option);
  });
}

function toggleReportCustomRange() {
  if (!reportQuarterSelect || !reportStartMonthSelect || !reportEndMonthSelect) {
    return;
  }
  const isCustom = reportQuarterSelect.value === "custom";
  reportStartMonthSelect.disabled = !isCustom;
  reportEndMonthSelect.disabled = !isCustom;
  reportStartMonthSelect.parentElement.style.display = isCustom ? "" : "none";
  reportEndMonthSelect.parentElement.style.display = isCustom ? "" : "none";
}

function buildReportParams() {
  const params = {};
  if (reportYearSelect) {
    params.year = reportYearSelect.value;
  }
  if (reportQuarterSelect) {
    if (reportQuarterSelect.value === "custom") {
      params.start_month = reportStartMonthSelect.value;
      params.end_month = reportEndMonthSelect.value;
    } else {
      params.quarter = reportQuarterSelect.value;
    }
  }
  return params;
}

function downloadQuarterlyReport() {
  if (!getSelectedCompanyId()) {
    alert("Selecciona una empresa antes de generar el informe.");
    return;
  }
  const params = buildReportParams();
  const query = new URLSearchParams(params).toString();
  const url = withCompanyParam(`/api/reports/quarterly?${query}`);
  if (reportStatus) {
    reportStatus.textContent = "Generando informe...";
  }
  fetch(url)
    .then((res) => {
      if (!res.ok) {
        return res.json().then((data) => {
          throw new Error((data.errors || ["Error al generar informe."]).join("\n"));
        });
      }
      return res.blob();
    })
    .then((blob) => {
      const link = document.createElement("a");
      const href = URL.createObjectURL(blob);
      link.href = href;
      link.download = "informe_fiscal.html";
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(href);
      if (reportStatus) {
        reportStatus.textContent = "Informe descargado.";
      }
    })
    .catch((error) => {
      if (reportStatus) {
        reportStatus.textContent = error.message || "No se pudo generar el informe.";
      }
    });
}

function sendQuarterlyReportEmail() {
  if (!getSelectedCompanyId()) {
    alert("Selecciona una empresa antes de enviar el informe.");
    return;
  }
  const payload = buildReportParams();
  const url = withCompanyParam("/api/reports/quarterly/email");
  if (reportStatus) {
    reportStatus.textContent = "Enviando informe por email...";
  }
  fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        throw new Error((data.errors || ["No se pudo enviar el informe."]).join("\n"));
      }
      if (reportStatus) {
        reportStatus.textContent = "Informe enviado correctamente.";
      }
    })
    .catch((error) => {
      if (reportStatus) {
        reportStatus.textContent = error.message || "No se pudo enviar el informe.";
      }
    });
}

function saveBillingEntry() {
  const month = Number(billingMonthSelect.value || monthSelect.value);
  const year = Number(billingYearSelect.value || yearSelect.value);
  const baseValue = billingBaseInput ? billingBaseInput.value : "";
  const totalValue = billingTotalInput ? billingTotalInput.value : "";
  const vatValue = billingVatSelect ? billingVatSelect.value : "";
  const conceptValue = billingConceptInput ? billingConceptInput.value.trim() : "";
  const dateValue = billingDateInput ? billingDateInput.value : "";

  if (!getSelectedCompanyId()) {
    alert("Selecciona una empresa antes de guardar ingresos.");
    return;
  }
  if (!month || !year) {
    alert("Selecciona mes y año.");
    return;
  }
  const resolvedVat = resolveVatRateValue(vatValue);
  const computed = calculateVatFields({
    baseValue: parseNumberInput(baseValue),
    totalValue: parseNumberInput(totalValue),
    vatRateValue: resolvedVat,
    source: billingLastSource,
  });
  const finalBase = computed.base !== null ? computed.base : null;
  if (finalBase === null || finalBase < 0) {
    alert("Base facturada inválida.");
    return;
  }
  if (billingBaseInput) {
    billingBaseInput.value = formatAmountInput(finalBase);
  }
  if (billingVatAmountInput && computed.vatAmount !== null) {
    billingVatAmountInput.value = formatAmountInput(computed.vatAmount);
  }
  if (billingTotalInput && computed.total !== null) {
    billingTotalInput.value = formatAmountInput(computed.total);
  }

  billingSaveBtn.disabled = true;
  fetch("/api/billing", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      company_id: getSelectedCompanyId(),
      month,
      year,
      base: finalBase,
      vat: resolvedVat,
      concept: conceptValue,
      invoice_date: dateValue,
    }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al guardar."]).join("\n"));
        return;
      }
      billingBaseInput.value = "";
      if (billingVatAmountInput) {
        billingVatAmountInput.value = "";
      }
      if (billingTotalInput) {
        billingTotalInput.value = "";
      }
      if (billingConceptInput) {
        billingConceptInput.value = "";
      }
      if (billingDateInput) {
        billingDateInput.value = "";
      }
      refreshBillingData();
    })
    .catch(() => {
      alert("No se pudo guardar la facturación.");
    })
    .finally(() => {
      billingSaveBtn.disabled = false;
    });
}

function setActiveSection(sectionId) {
  sections.forEach((section) => {
    section.classList.toggle("active", section.dataset.section === sectionId);
  });
  navLinks.forEach((link) => {
    link.classList.toggle("active", link.dataset.section === sectionId);
  });
  localStorage.setItem("activeSection", sectionId);
  document.body.classList.remove("sidebar-open");
}

function initNavigation() {
  const storedSection = localStorage.getItem("activeSection");
  const availableSections = Array.from(sections || []).map(
    (section) => section.dataset.section
  );
  const defaultSection = availableSections.includes(storedSection)
    ? storedSection
    : availableSections[0] || "dashboard";
  setActiveSection(defaultSection);

  navLinks.forEach((link) => {
    link.addEventListener("click", () => {
      setActiveSection(link.dataset.section);
    });
  });

  if (sidebarToggle) {
    sidebarToggle.addEventListener("click", () => {
      document.body.classList.toggle("sidebar-open");
    });
  }
  if (sidebarOverlay) {
    sidebarOverlay.addEventListener("click", () => {
      document.body.classList.remove("sidebar-open");
    });
  }
}

function bindEvents() {
  const selectFilesBtn = document.getElementById("selectFiles");
  const selectFolderBtn = document.getElementById("selectFolder");
  const incomeSelectFilesBtn = document.getElementById("incomeSelectFiles");
  const lowQualityAccept = document.getElementById("lowQualityAccept");
  const lowQualityClose = document.getElementById("lowQualityClose");
  if (lowQualityAccept) {
    lowQualityAccept.addEventListener("click", hideLowQualityModal);
  }
  if (lowQualityClose) {
    lowQualityClose.addEventListener("click", hideLowQualityModal);
  }
  if (paymentPrevMonth) {
    paymentPrevMonth.addEventListener("click", () => shiftCalendarMonth(-1));
  }
  if (paymentNextMonth) {
    paymentNextMonth.addEventListener("click", () => shiftCalendarMonth(1));
  }
  if (selectFilesBtn && fileInput) {
    selectFilesBtn.addEventListener("click", () => {
      fileInput.click();
    });
  }
  if (selectFolderBtn && folderInput) {
    selectFolderBtn.addEventListener("click", () => {
      folderInput.click();
    });
  }
  if (incomeFileInput && incomeSelectFilesBtn) {
    incomeSelectFilesBtn.addEventListener("click", () => {
      incomeFileInput.click();
    });
  }
  if (fileInput) {
    fileInput.addEventListener("change", (event) => {
      addFiles(event.target.files);
      fileInput.value = "";
    });
  }
  if (folderInput) {
    folderInput.addEventListener("change", (event) => {
      addFiles(event.target.files);
      folderInput.value = "";
    });
  }
  if (incomeFileInput) {
    incomeFileInput.addEventListener("change", (event) => {
      addIncomeFiles(event.target.files);
      incomeFileInput.value = "";
    });
  }

  if (dropZone) {
    dropZone.addEventListener("dragover", (event) => {
      event.preventDefault();
      dropZone.classList.add("dragover");
    });
    dropZone.addEventListener("dragleave", () => {
      dropZone.classList.remove("dragover");
    });
    dropZone.addEventListener("drop", (event) => {
      event.preventDefault();
      dropZone.classList.remove("dragover");
      if (event.dataTransfer.files) {
        addFiles(event.dataTransfer.files);
      }
    });
  }
  if (incomeDropZone) {
    incomeDropZone.addEventListener("dragover", (event) => {
      event.preventDefault();
      incomeDropZone.classList.add("dragover");
    });
    incomeDropZone.addEventListener("dragleave", () => {
      incomeDropZone.classList.remove("dragover");
    });
    incomeDropZone.addEventListener("drop", (event) => {
      event.preventDefault();
      incomeDropZone.classList.remove("dragover");
      if (event.dataTransfer.files) {
        addIncomeFiles(event.dataTransfer.files);
      }
    });
  }

  if (uploadBtn) {
    uploadBtn.addEventListener("click", uploadPending);
  }
  if (incomeUploadBtn) {
    incomeUploadBtn.addEventListener("click", uploadIncomePending);
  }
  if (companySaveBtn) {
    companySaveBtn.addEventListener("click", saveCompany);
  }
  if (staffSaveBtn) {
    staffSaveBtn.addEventListener("click", saveStaff);
  }
  if (monthSelect) {
    monthSelect.addEventListener("change", () => {
      persistFilters();
      updateHeaderContext();
      calendarOverride = false;
      syncCalendarWithFilters();
      refreshAllData();
    });
  }
  if (yearSelect) {
    yearSelect.addEventListener("change", () => {
      persistFilters();
      updateHeaderContext();
      calendarOverride = false;
      syncCalendarWithFilters();
      refreshAllData();
    });
  }
  if (periodSelect) {
    periodSelect.addEventListener("change", () => {
      document.body.classList.toggle(
        "period-quarterly",
        getSelectedPeriod() === "quarterly"
      );
      persistFilters();
      updateHeaderContext();
      calendarOverride = false;
      syncCalendarWithFilters();
      refreshAllData();
    });
  }
  if (companySelect) {
    companySelect.addEventListener("change", () => {
      selectedCompanyId = companySelect.value;
      persistFilters();
      applyCompanyTaxModules();
      updatePnlSummary();
      updateHeaderContext();
      calendarOverride = false;
      syncCalendarWithFilters();
      renderTable();
      renderIncomeTable();
      loadYears().then(() => refreshAllData());
    });
  }
  if (billingBaseInput) {
    billingBaseInput.addEventListener("input", () => {
      billingLastSource = "base";
      syncBillingCalculation("base");
    });
  }
  if (billingTotalInput) {
    billingTotalInput.addEventListener("input", () => {
      billingLastSource = "total";
      syncBillingCalculation("total");
    });
  }
  if (billingVatSelect) {
    billingVatSelect.addEventListener("change", () => {
      syncBillingCalculation(billingLastSource);
    });
  }
  if (billingSaveBtn) {
    billingSaveBtn.addEventListener("click", saveBillingEntry);
  }
  if (noInvoiceSaveBtn) {
    noInvoiceSaveBtn.addEventListener("click", saveNoInvoiceExpense);
  }
  if (exportPnlBtn) {
    exportPnlBtn.addEventListener("click", exportPnlPdf);
  }
  if (reportQuarterSelect) {
    reportQuarterSelect.addEventListener("change", () => {
      toggleReportCustomRange();
    });
  }
  if (reportDownloadBtn) {
    reportDownloadBtn.addEventListener("click", downloadQuarterlyReport);
  }
  if (reportEmailBtn) {
    reportEmailBtn.addEventListener("click", sendQuarterlyReportEmail);
  }
  bindPnlInputs();
}

function init() {
  const now = new Date();
  populateMonthSelects();
  populateReportMonthSelect(reportStartMonthSelect);
  populateReportMonthSelect(reportEndMonthSelect);
  toggleReportCustomRange();
  bindEvents();
  initNavigation();
  applyRoleVisibility();
  restoreFilters(now);
  syncCalendarWithFilters();
  loadStaff()
    .then(() => loadCompanies())
    .then(() => loadYears())
    .then(() => {
      document.body.classList.toggle("period-quarterly", getSelectedPeriod() === "quarterly");
      billingMonthSelect.value = monthSelect.value;
      billingYearSelect.value = yearSelect.value;
      if (reportStartMonthSelect) {
        reportStartMonthSelect.value = monthSelect.value;
      }
      if (reportEndMonthSelect) {
        reportEndMonthSelect.value = monthSelect.value;
      }
      if (billingDateInput && !billingDateInput.value) {
        billingDateInput.value = now.toISOString().slice(0, 10);
      }
      if (!noInvoiceDate.value) {
        noInvoiceDate.value = now.toISOString().slice(0, 10);
      }
      updateHeaderContext();
      refreshAllData();
    });
}

document.addEventListener("DOMContentLoaded", init);
