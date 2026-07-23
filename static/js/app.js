const pendingFiles = [];
let lineChart;
let billingLineChart;
let pieChart;
let netChart;
let expenseVatTotal = 0;
let expenseVatSupportedTotal = 0;
let expenseVatDeductibleTotal = 0;
let billingVatTotal = 0;
let incomeVatOutputTotal = 0;
let incomeGrossTotal = 0;
let expenseGrossTotal = 0;
let currentInvoices = [];
let currentNoInvoiceExpenses = [];
let currentLoanInstallments = [];
let billingBaseTotal = 0;
let currentSummary = null;
let currentBillingSummary = null;
let currentFinancialMetrics = null;
let currentBillingEntries = [];
let currentDeductibleExpenses = 0;
let annualBillingBaseTotal = 0;
let annualDeductibleExpenses = 0;
let annualLoanInterestTotal = 0;
let annualTaxEstimateTotal = 0;
let pnlInvoices = [];
let pnlNoInvoiceExpenses = [];
let pnlLoanInstallments = [];
let pnlIncomeInvoices = [];
let pnlDataReady = false;
let currentPayments = null;
let selectedPaymentDay = null;
let calendarMonth = null;
let calendarYear = null;
let calendarOverride = false;
let currentBalanceManualData = {};
let currentDocumentBatches = [];
let currentDocumentCenterDocuments = [];
let currentAccountingIntegrationSummary = null;
let currentArchiveRecords = [];
let selectedDocumentBatchId = "";
let selectedDocumentCenterDocumentId = null;
let companies = [];
let selectedCompanyId = null;
let pendingIncomeFiles = [];
let currentIncomeInvoices = [];
let staffMembers = [];
let selectedLoanPlanFile = null;
let loanPlanDraft = [];
let currentExpenseSubview = "received";
let currentExpenseMode = "adjustments";
let currentReportsSubview = "summary";
let currentStatementsSubview = "pnl";
const lowQualityDismissedIds = new Set();
const vatWarningDismissedIds = new Set();
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
  alquiler_local: "Alquiler local",
  alquiler_cabina: "Alquiler cabina",
  nomina: "Nómina",
  seguridad_social: "Seguridad Social",
  amortizacion: "Amortización",
  kilometraje: "Kilometraje",
  prestamo: "Financiación bancaria",
  otro: "Otro ajuste",
};

const expenseModeTypes = {
  rent: ["alquiler_local", "alquiler_cabina"],
  payroll: ["nomina", "seguridad_social"],
  financing: ["prestamo"],
  adjustments: ["amortizacion", "kilometraje", "otro"],
};

const ANALYSIS_ERROR_MESSAGE =
  "No se ha podido analizar la factura automáticamente. Puedes introducir los datos manualmente.";
const LOW_QUALITY_SCAN_MESSAGE =
  "La calidad de la factura escaneada no es óptima. No se puede leer correctamente. Por favor, introduce los datos manualmente.";
const TIMEOUT_MESSAGE =
  "El análisis tardó demasiado y se detuvo. Puedes introducir los datos manualmente.";
const VAT_WARNING_MESSAGE =
  "Puede que la calidad de la imagen o la información sea dudosa. Por favor, revisa siempre las cantidades y los tipos de IVA.";
const ANALYSIS_MAX_CONCURRENCY = 1;
const ANALYSIS_PENDING_TIMEOUT_MS = 70 * 1000;

const analysisTaskQueue = [];
let activeAnalysisTasks = 0;

function isPendingUploadItemPresent(item) {
  return (
    pendingFiles.some((entry) => entry.id === item.id) ||
    pendingIncomeFiles.some((entry) => entry.id === item.id)
  );
}

function removeQueuedAnalysisTask(itemId) {
  const index = analysisTaskQueue.findIndex((task) => task.item.id === itemId);
  if (index !== -1) {
    analysisTaskQueue.splice(index, 1);
  }
}

function abortPendingAnalysis(item) {
  item._analysisCancelled = true;
  item.analysisPending = false;
  item.analysisQueued = false;
  removeQueuedAnalysisTask(item.id);
  if (item._analysisTimeoutId) {
    clearTimeout(item._analysisTimeoutId);
    item._analysisTimeoutId = null;
  }
  if (item._analysisController) {
    item._analysisController.abort();
  }
}

function scheduleAnalysisTimeout(item, render) {
  if (item._analysisTimeoutId) {
    clearTimeout(item._analysisTimeoutId);
  }
  item._analysisTimeoutId = window.setTimeout(() => {
    if (item._analysisCancelled || !isPendingUploadItemPresent(item) || !item.analysisPending) {
      return;
    }
    removeQueuedAnalysisTask(item.id);
    if (item._analysisController) {
      item._analysisController.abort();
    }
    item.analysisPending = false;
    item.analysisQueued = false;
    item.analysisError = true;
    item.analysisErrorMessage = TIMEOUT_MESSAGE;
    item.analysisStatus = "timeout";
    if (typeof render === "function") {
      render();
    }
  }, ANALYSIS_PENDING_TIMEOUT_MS);
}

function processAnalysisQueue() {
  while (activeAnalysisTasks < ANALYSIS_MAX_CONCURRENCY && analysisTaskQueue.length) {
    const task = analysisTaskQueue.shift();
    if (!task || !task.item || task.item._analysisCancelled || !isPendingUploadItemPresent(task.item)) {
      continue;
    }
    activeAnalysisTasks += 1;
    task.item.analysisQueued = false;
    if (typeof task.render === "function") {
      task.render();
    }
    Promise.resolve()
      .then(() => task.run(task.item))
      .catch(() => undefined)
      .finally(() => {
        activeAnalysisTasks = Math.max(activeAnalysisTasks - 1, 0);
        processAnalysisQueue();
      });
  }
}

function enqueueAnalysisTask(item, run, render) {
  item.analysisPending = true;
  item.analysisQueued = true;
  item._analysisCancelled = false;
  scheduleAnalysisTimeout(item, render);
  analysisTaskQueue.push({ item, run, render });
  processAnalysisQueue();
}

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

function showVatWarningModal() {
  const modal = document.getElementById("vatWarningModal");
  if (!modal) {
    return;
  }
  modal.classList.add("is-visible");
  modal.setAttribute("aria-hidden", "false");
  document.body.classList.add("modal-open");
}

function hideVatWarningModal() {
  const modal = document.getElementById("vatWarningModal");
  if (!modal) {
    return;
  }
  modal.classList.remove("is-visible");
  modal.setAttribute("aria-hidden", "true");
  document.body.classList.remove("modal-open");
}

const currencyFormatter = new Intl.NumberFormat("es-ES", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

function formatCurrency(value) {
  let number = Number(value);
  if (!Number.isFinite(number)) {
    const parsed = parseNumberInput(value);
    if (parsed === null) {
      return "0,00 €";
    }
    number = parsed;
  }
  return `${currencyFormatter.format(number)} €`;
}

function formatPercent(value) {
  if (value === null || value === undefined || value === "") {
    return "-";
  }
  const number = Number(value);
  if (!Number.isFinite(number)) {
    return "-";
  }
  const digits = Number.isInteger(number) ? 0 : 2;
  const formatter = new Intl.NumberFormat("es-ES", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
  return `${formatter.format(number)}%`;
}

function parseNumberInput(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  let cleaned = String(value)
    .replace(/[%€]/g, "")
    .replace(/\s/g, "")
    .trim();
  if (cleaned.includes(",") && cleaned.includes(".")) {
    cleaned = cleaned.replace(/\./g, "").replace(",", ".");
  } else {
    cleaned = cleaned.replace(",", ".");
  }
  const numeric = Number(cleaned);
  return Number.isNaN(numeric) ? null : numeric;
}

function roundAmount(value) {
  return Math.round(Number(value) * 100) / 100;
}

function formatAmountInput(value) {
  if (value === null || value === undefined) {
    return "";
  }
  const parsed =
    typeof value === "string" ? parseNumberInput(value) : Number(value);
  if (parsed === null || Number.isNaN(parsed)) {
    return "";
  }
  return currencyFormatter.format(parsed);
}

function attachAmountInputBehavior(input) {
  if (!input) {
    return;
  }
  input.addEventListener("focus", () => {
    if (typeof input.select === "function") {
      input.select();
    }
  });
  input.addEventListener("blur", () => {
    const parsed = parseNumberInput(input.value);
    if (parsed === null) {
      input.value = "";
      return;
    }
    input.value = formatAmountInput(parsed);
  });
}

function getPnlInputValue(id) {
  const el = document.getElementById(id);
  if (!el) {
    return 0;
  }
  const parsed = parseNumberInput(el.value);
  return parsed === null ? 0 : parsed;
}

function getBalanceInputValue(id) {
  const el = document.getElementById(id);
  if (!el) {
    return 0;
  }
  const parsed = parseNumberInput(el.value);
  return parsed === null ? 0 : parsed;
}

function setBalanceInputValue(id, value) {
  const el = document.getElementById(id);
  if (!el) {
    return;
  }
  el.value = formatAmountInput(value);
}

function getBalanceManualValue(id) {
  return Number(currentBalanceManualData?.[id]) || 0;
}

function getBalanceAutoValues(netResult = null) {
  const autoValues = {
    bsAssetReceivables: getBalanceReceivablesEstimate(),
    bsAssetCash: getBalanceCashEstimate(),
    bsAssetCurrentTax: getCurrentTaxAssetEstimate(),
    bsEquityResult: netResult !== null ? netResult : getBalanceInputValue("bsEquityResult"),
    bsLiabShortTermDebt: getBalanceShortTermDebtEstimate(),
    bsLiabPayables: getBalancePayablesEstimate(),
    bsLiabOtherCurrent: getOutstandingTaxLiabilitiesEstimate(),
  };
  return autoValues;
}

function setBalanceFieldDisplayValues(netResult = null) {
  const autoValues = getBalanceAutoValues(netResult);
  balanceInputIds.forEach((id) => {
    const totalValue = getBalanceManualValue(id) + (Number(autoValues[id]) || 0);
    setBalanceInputValue(id, totalValue);
  });
}

function updateBalanceTotals() {
  if (!bsTotalAssets || !bsTotalLiabilities) {
    return;
  }
  const assetNonCurrent =
    getBalanceInputValue("bsAssetIntangible") +
    getBalanceInputValue("bsAssetTangible") +
    getBalanceInputValue("bsAssetInvestmentProperty") +
    getBalanceInputValue("bsAssetLongTermInv") +
    getBalanceInputValue("bsAssetDeferredTax");
  const assetCurrent =
    getBalanceInputValue("bsAssetInventory") +
    getBalanceInputValue("bsAssetReceivables") +
    getBalanceInputValue("bsAssetShortTermInv") +
    getBalanceInputValue("bsAssetCash") +
    getBalanceInputValue("bsAssetCurrentTax");
  const equityTotal =
    getBalanceInputValue("bsEquityCapital") +
    getBalanceInputValue("bsEquityReserves") +
    getBalanceInputValue("bsEquityResult") +
    getBalanceInputValue("bsEquityGrants");
  const liabNonCurrent =
    getBalanceInputValue("bsLiabLongTermDebt") +
    getBalanceInputValue("bsLiabLongTermOther");
  const liabCurrent =
    getBalanceInputValue("bsLiabShortTermDebt") +
    getBalanceInputValue("bsLiabPayables") +
    getBalanceInputValue("bsLiabOtherCurrent");
  const totalAssets = assetNonCurrent + assetCurrent;
  const totalLiabilities = equityTotal + liabNonCurrent + liabCurrent;

  if (bsAssetNonCurrentTotal) {
    bsAssetNonCurrentTotal.textContent = formatCurrency(assetNonCurrent);
  }
  if (bsAssetCurrentTotal) {
    bsAssetCurrentTotal.textContent = formatCurrency(assetCurrent);
  }
  if (bsEquityTotal) {
    bsEquityTotal.textContent = formatCurrency(equityTotal);
  }
  if (bsLiabNonCurrentTotal) {
    bsLiabNonCurrentTotal.textContent = formatCurrency(liabNonCurrent);
  }
  if (bsLiabCurrentTotal) {
    bsLiabCurrentTotal.textContent = formatCurrency(liabCurrent);
  }
  bsTotalAssets.textContent = formatCurrency(totalAssets);
  bsTotalLiabilities.textContent = formatCurrency(totalLiabilities);
}

function syncBalanceStatementFields(netResult) {
  setBalanceFieldDisplayValues(netResult);
  renderBalanceMissingInfo(netResult);
  updateBalanceTotals();
}

function normalizePaymentDateList(value) {
  if (Array.isArray(value)) {
    return value.filter(Boolean).map((item) => String(item));
  }
  if (!value) {
    return [];
  }
  if (typeof value === "string") {
    try {
      const parsed = JSON.parse(value);
      if (Array.isArray(parsed)) {
        return parsed.filter(Boolean).map((item) => String(item));
      }
    } catch (error) {
      return [value];
    }
    return [];
  }
  return [];
}

function splitScheduledAmounts(totalAmount, paymentDates) {
  const total = Number(totalAmount) || 0;
  const dates = normalizePaymentDateList(paymentDates);
  if (!dates.length) {
    return [];
  }
  const splitAmount = roundAmount(total / dates.length);
  const amounts = new Array(dates.length).fill(splitAmount);
  if (dates.length > 1) {
    amounts[dates.length - 1] = roundAmount(total - splitAmount * (dates.length - 1));
  }
  return amounts;
}

function getInvoicePaymentDates(invoice) {
  const paymentDates = normalizePaymentDateList(invoice?.payment_dates);
  if (paymentDates.length) {
    return paymentDates;
  }
  if (invoice?.payment_date) {
    return [invoice.payment_date];
  }
  if (invoice?.invoice_date) {
    return [computePaymentDate(invoice.invoice_date, invoice.payment_date)];
  }
  return [];
}

function getInvoiceCashAmount(invoice) {
  if (!invoice) {
    return 0;
  }
  return Math.max(
    (Number(invoice.total_amount) || 0) - (Number(invoice.withholding_amount) || 0),
    0
  );
}

function getNoInvoicePaymentDates(expense) {
  const paymentDates = normalizePaymentDateList(expense?.payment_dates);
  if (paymentDates.length) {
    return paymentDates;
  }
  if (expense?.payment_date) {
    return [expense.payment_date];
  }
  if (expense?.expense_date) {
    return [expense.expense_date];
  }
  return [];
}

function getNoInvoiceCashAmount(expense) {
  if (!expense) {
    return 0;
  }
  if (expense.expense_type === "amortizacion") {
    return 0;
  }
  if (expense.expense_type === "nomina") {
    return Number(expense.payroll_net_amount) || Math.max((Number(expense.amount) || 0) - (Number(expense.withholding_amount) || 0), 0);
  }
  return Math.max((Number(expense.amount) || 0) - (Number(expense.withholding_amount) || 0), 0);
}

function sumCompletedScheduledAmount(totalAmount, paymentDates, completedDates) {
  const dates = normalizePaymentDateList(paymentDates);
  if (!dates.length) {
    return 0;
  }
  const completed = new Set(normalizePaymentDateList(completedDates));
  const amounts = splitScheduledAmounts(totalAmount, dates);
  return dates.reduce((sum, paymentDate, index) => {
    if (!completed.has(paymentDate)) {
      return sum;
    }
    return sum + (Number(amounts[index]) || 0);
  }, 0);
}

function getIncomeInvoicePaymentDates(invoice) {
  const paymentDates = normalizePaymentDateList(invoice?.payment_dates);
  if (paymentDates.length) {
    return paymentDates;
  }
  if (invoice?.payment_date) {
    return [invoice.payment_date];
  }
  if (invoice?.invoice_date) {
    return [computePaymentDate(invoice.invoice_date, invoice.payment_date)];
  }
  return [];
}

function getOutstandingTaxLiabilitiesEstimate() {
  return roundAmount(
    (currentPayments?.items || []).reduce((sum, item) => {
      if (item?.type !== "tax_obligation" || item?.calendar_impact !== "payment") {
        return sum;
      }
      return sum + (Number(item.amount) || 0);
    }, 0)
  );
}

function getCurrentTaxAssetEstimate() {
  return roundAmount(
    (currentPayments?.items || []).reduce((sum, item) => {
      if (item?.type !== "tax_obligation") {
        return sum;
      }
      if (item.calendar_impact === "credit" || item.calendar_impact === "refund") {
        return sum + (Number(item.amount) || 0);
      }
      return sum;
    }, 0)
  );
}

function getBalanceShortTermDebtEstimate() {
  return roundAmount(
    (currentLoanInstallments || []).reduce((sum, installment) => {
      return (
        sum +
        sumOutstandingScheduledAmount(
          Number(installment.total_amount) || 0,
          normalizePaymentDateList([installment.payment_date]),
          installment.payment_completed_dates
        )
      );
    }, 0)
  );
}

function sumOutstandingScheduledAmount(totalAmount, paymentDates, completedDates) {
  const dates = normalizePaymentDateList(paymentDates);
  if (!dates.length) {
    return 0;
  }
  const completed = new Set(normalizePaymentDateList(completedDates));
  const amounts = splitScheduledAmounts(totalAmount, dates);
  return dates.reduce((sum, paymentDate, index) => {
    if (completed.has(paymentDate)) {
      return sum;
    }
    return sum + (Number(amounts[index]) || 0);
  }, 0);
}

function getBalancePayablesEstimate() {
  const invoicePayables = (currentInvoices || []).reduce((sum, invoice) => {
    return (
      sum +
      sumOutstandingScheduledAmount(
        getInvoiceCashAmount(invoice),
        getInvoicePaymentDates(invoice),
        invoice.payment_completed_dates
      )
    );
  }, 0);

  const noInvoicePayables = (currentNoInvoiceExpenses || []).reduce((sum, expense) => {
    if (expense.expense_type === "prestamo" || expense.expense_type === "amortizacion") {
      return sum;
    }
    return (
      sum +
      sumOutstandingScheduledAmount(
        getNoInvoiceCashAmount(expense),
        getNoInvoicePaymentDates(expense),
        expense.payment_completed_dates
      )
    );
  }, 0);

  return roundAmount(invoicePayables + noInvoicePayables);
}

function getBalanceReceivablesEstimate() {
  return roundAmount(
    (currentIncomeInvoices || []).reduce((sum, invoice) => {
      return (
        sum +
        sumOutstandingScheduledAmount(
          Number(invoice.total_amount) || 0,
          getIncomeInvoicePaymentDates(invoice),
          invoice.payment_completed_dates
        )
      );
    }, 0)
  );
}

function getBalanceCashEstimate() {
  const incomeInvoicesTreasury = (currentIncomeInvoices || []).reduce(
    (sum, invoice) =>
      sum +
      sumCompletedScheduledAmount(
        Number(invoice.total_amount) || 0,
        getIncomeInvoicePaymentDates(invoice),
        invoice.payment_completed_dates
      ),
    0
  );
  const invoiceOutflows = (currentInvoices || []).reduce(
    (sum, invoice) =>
      sum +
      sumCompletedScheduledAmount(
        getInvoiceCashAmount(invoice),
        getInvoicePaymentDates(invoice),
        invoice.payment_completed_dates
      ),
    0
  );
  const noInvoiceOutflows = (currentNoInvoiceExpenses || []).reduce(
    (sum, expense) =>
      sum +
      sumCompletedScheduledAmount(
        getNoInvoiceCashAmount(expense),
        getNoInvoicePaymentDates(expense),
        expense.payment_completed_dates
      ),
    0
  );
  const loanOutflows = (currentLoanInstallments || []).reduce(
    (sum, installment) =>
      sum +
      sumCompletedScheduledAmount(
        Number(installment.total_amount) || 0,
        normalizePaymentDateList([installment.payment_date]),
        installment.payment_completed_dates
      ),
    0
  );
  return roundAmount(
    incomeInvoicesTreasury - invoiceOutflows - noInvoiceOutflows - loanOutflows
  );
}

function getCurrentBalanceManualPayload(netResult = null) {
  const autoValues = getBalanceAutoValues(netResult);
  return balanceInputIds.reduce((acc, id) => {
    const enteredValue = getBalanceInputValue(id);
    const manualValue = roundAmount(enteredValue - (Number(autoValues[id]) || 0));
    if (Math.abs(manualValue) >= 0.005) {
      acc[id] = manualValue;
    }
    return acc;
  }, {});
}

function loadBalanceManualDataFromCompany() {
  const company = getSelectedCompany();
  if (!company || !company.balance_manual_data) {
    currentBalanceManualData = {};
    if (balanceSaveStatus) {
      balanceSaveStatus.textContent = "";
    }
    return;
  }
  const rawValue = company.balance_manual_data;
  if (typeof rawValue === "object" && rawValue !== null) {
    currentBalanceManualData = rawValue;
    if (balanceSaveStatus) {
      balanceSaveStatus.textContent = "";
    }
    return;
  }
  try {
    const parsed = JSON.parse(rawValue);
    currentBalanceManualData = parsed && typeof parsed === "object" ? parsed : {};
  } catch (error) {
    currentBalanceManualData = {};
  }
  if (balanceSaveStatus) {
    balanceSaveStatus.textContent = "";
  }
}

function renderBalanceMissingInfo(netResult = null) {
  if (!balanceMissingInfo) {
    return;
  }
  const autoValues = getBalanceAutoValues(netResult);
  const manualPayload = getCurrentBalanceManualPayload(netResult);
  const missing = [];

  balanceAlwaysManualFieldIds.forEach((id) => {
    const totalValue = (Number(autoValues[id]) || 0) + (Number(manualPayload[id]) || 0);
    if (Math.abs(totalValue) < 0.005) {
      missing.push(balanceFieldLabels[id]);
    }
  });

  if ((currentLoanInstallments || []).length) {
    const debtTotal =
      getBalanceInputValue("bsLiabShortTermDebt") + getBalanceInputValue("bsLiabLongTermDebt");
    if (Math.abs(debtTotal) < 0.005) {
      missing.push("deudas a corto o largo plazo vinculadas a financiación");
    }
  }

  if ((currentBillingEntries || []).length) {
    missing.push(
      "cobros reales de la facturación manual, que hoy impacta en resultado pero no en tesorería ni clientes hasta que se registre su seguimiento de cobro"
    );
  }

  if (!missing.length) {
    balanceMissingInfo.hidden = true;
    balanceMissingInfo.textContent = "";
    return;
  }

  balanceMissingInfo.hidden = false;
  balanceMissingInfo.textContent =
    `Información pendiente para que el balance sea completo: ${missing.join(", ")}. ` +
    "Ledged seguirá calculando automáticamente las partidas que conoce y conservará estos saldos manuales por empresa.";
}

function saveBalanceManualData() {
  const companyId = getSelectedCompanyId();
  if (!companyId) {
    alert("Selecciona una empresa antes de guardar el balance.");
    return;
  }
  const netResult = parseNumberInput(document.getElementById("pnlNet")?.textContent || "") || 0;
  const payload = getCurrentBalanceManualPayload(netResult);
  if (balanceSaveBtn) {
    balanceSaveBtn.disabled = true;
  }
  if (balanceSaveStatus) {
    balanceSaveStatus.textContent = "Guardando...";
  }
  fetch(`/api/companies/${companyId}/balance-manual-data`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      balance_manual_data: payload,
    }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["No se pudo guardar el balance."]).join("\n"));
        return;
      }
      currentBalanceManualData = data.balance_manual_data || {};
      const company = getSelectedCompany();
      if (company) {
        company.balance_manual_data = currentBalanceManualData;
      }
      renderBalanceMissingInfo(netResult);
      if (balanceSaveStatus) {
        balanceSaveStatus.textContent = "Saldos manuales guardados.";
      }
    })
    .catch(() => {
      alert("No se pudo guardar el balance.");
    })
    .finally(() => {
      if (balanceSaveBtn) {
        balanceSaveBtn.disabled = false;
      }
    });
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

function bindBalanceInputs() {
  balanceInputIds.forEach((id) => {
    const el = document.getElementById(id);
    if (!el) {
      return;
    }
    el.addEventListener("input", () => {
      const autoValues = getBalanceAutoValues();
      const manualValue = roundAmount(
        getBalanceInputValue(id) - (Number(autoValues[id]) || 0)
      );
      if (Math.abs(manualValue) < 0.005) {
        delete currentBalanceManualData[id];
      } else {
        currentBalanceManualData[id] = manualValue;
      }
      renderBalanceMissingInfo();
      updateBalanceTotals();
    });
  });
  if (balanceSaveBtn) {
    balanceSaveBtn.addEventListener("click", saveBalanceManualData);
  }
  updateBalanceTotals();
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

  if (source !== "base") {
    if (result.base !== null) {
      inputs.base.value = formatAmountInput(result.base);
    } else if (source === "total") {
      inputs.base.value = "";
    }
  }
  if (source !== "vatAmount" && result.vatAmount !== null) {
    inputs.vatAmount.value = formatAmountInput(result.vatAmount);
  } else if (source !== "vatAmount") {
    inputs.vatAmount.value = "";
  }
  if (source !== "total" && result.total !== null) {
    inputs.total.value = formatAmountInput(result.total);
  } else if (source !== "total") {
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

  if (source !== "base" && result.base !== null) {
    billingBaseInput.value = formatAmountInput(result.base);
  } else if (source !== "base" && source === "total") {
    billingBaseInput.value = "";
  }
  if (source !== "total" && result.total !== null) {
    billingTotalInput.value = formatAmountInput(result.total);
  } else if (source !== "total" && source === "base") {
    billingTotalInput.value = "";
  }
  if (source !== "vatAmount" && result.vatAmount !== null) {
    billingVatAmountInput.value = formatAmountInput(result.vatAmount);
  } else if (source !== "vatAmount") {
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

function addDaysToISO(dateStr, days) {
  if (!dateStr) {
    return "";
  }
  const base = new Date(`${dateStr}T00:00:00`);
  if (Number.isNaN(base.getTime())) {
    return "";
  }
  base.setDate(base.getDate() + Number(days || 0));
  return base.toISOString().slice(0, 10);
}

function getEffectivePaymentDates(item) {
  const dates = Array.isArray(item.paymentDates)
    ? item.paymentDates.filter((d) => Boolean(d))
    : [];
  if (!dates.length && item.paymentDate) {
    dates.push(item.paymentDate);
  }
  if (!dates.length && item.date) {
    const fallback = isPayrollUploadItem(item)
      ? item.date
      : computePaymentDate(item.date, item.paymentDate);
    if (fallback) {
      dates.push(fallback);
    }
  }
  return dates;
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
const billingSearchInput = document.getElementById("billingSearchInput");
const invoicesTableBody = document.querySelector("#invoicesTable tbody");
const invoicesEmpty = document.getElementById("invoicesEmpty");
const invoiceSearchInput = document.getElementById("invoiceSearchInput");
const taxPeriodBadge = document.getElementById("taxPeriodBadge");
const noInvoiceDate = document.getElementById("noInvoiceDate");
const noInvoiceConcept = document.getElementById("noInvoiceConcept");
const noInvoiceAmount = document.getElementById("noInvoiceAmount");
const noInvoiceInterest = document.getElementById("noInvoiceInterest");
const noInvoiceInterestField = document.getElementById("noInvoiceInterestField");
const noInvoiceWithholding = document.getElementById("noInvoiceWithholding");
const noInvoiceWithholdingField = document.getElementById("noInvoiceWithholdingField");
const noInvoiceVatDeductible = document.getElementById("noInvoiceVatDeductible");
const noInvoiceVatSelect = document.getElementById("noInvoiceVatSelect");
const noInvoiceVatBase = document.getElementById("noInvoiceVatBase");
const noInvoiceVatAmount = document.getElementById("noInvoiceVatAmount");
const noInvoiceVatRateField = document.getElementById("noInvoiceVatRateField");
const noInvoiceVatBaseField = document.getElementById("noInvoiceVatBaseField");
const noInvoiceVatAmountField = document.getElementById("noInvoiceVatAmountField");
const noInvoiceType = document.getElementById("noInvoiceType");
const noInvoiceDeductible = document.getElementById("noInvoiceDeductible");
const noInvoiceSaveBtn = document.getElementById("noInvoiceSaveBtn");
const noInvoiceTableBody = document.querySelector("#noInvoiceTable tbody");
const noInvoiceEmpty = document.getElementById("noInvoiceEmpty");
const noInvoiceSearchInput = document.getElementById("noInvoiceSearchInput");
const loanConceptInput = document.getElementById("loanConceptInput");
const loanPaymentDateInput = document.getElementById("loanPaymentDateInput");
const loanTotalInput = document.getElementById("loanTotalInput");
const loanInterestInput = document.getElementById("loanInterestInput");
const loanPrincipalInput = document.getElementById("loanPrincipalInput");
const loanSaveBtn = document.getElementById("loanSaveBtn");
const loanPlanFile = document.getElementById("loanPlanFile");
const loanPlanDropZone = document.getElementById("loanPlanDropZone");
const loanPlanSelectBtn = document.getElementById("loanPlanSelectBtn");
const loanPlanFileName = document.getElementById("loanPlanFileName");
const loanImportBtn = document.getElementById("loanImportBtn");
const loanPlanPreviewBody = document.querySelector("#loanPlanPreviewTable tbody");
const loanPlanPreviewEmpty = document.getElementById("loanPlanPreviewEmpty");
const loanPlanSaveBtn = document.getElementById("loanPlanSaveBtn");
const loanTableBody = document.querySelector("#loanTable tbody");
const loanEmpty = document.getElementById("loanEmpty");
const loanSearchInput = document.getElementById("loanSearchInput");
const fileInput = document.getElementById("fileInput");
const folderInput = document.getElementById("folderInput");
const dropZone = document.getElementById("dropZone");
const uploadTableHeadCells = document.querySelectorAll("#uploadTable thead th");
const uploadTableBody = document.querySelector("#uploadTable tbody");
const emptyMessage = document.getElementById("emptyMessage");
const uploadBtn = document.getElementById("uploadBtn");
const expenseUploadTitle = document.getElementById("expenseUploadTitle");
const expenseUploadDescription = document.getElementById("expenseUploadDescription");
const expensesSavedInvoicesPanel = document.getElementById("expensesSavedInvoicesPanel");
const navLinks = document.querySelectorAll(".nav-link[data-section]");
const sections = document.querySelectorAll(".page-section");
const sectionTabs = document.querySelectorAll(".section-tab[data-parent]");
const sidebarToggle = document.getElementById("sidebarToggle");
const sidebarOverlay = document.getElementById("sidebarOverlay");
const exportPnlBtn = document.getElementById("exportPnlBtn");
const pnlEmailBtn = document.getElementById("pnlEmailBtn");
const pnlName = document.getElementById("pnlName");
const pnlTaxId = document.getElementById("pnlTaxId");
const balancePdfBtn = document.getElementById("balancePdfBtn");
const balanceEmailBtn = document.getElementById("balanceEmailBtn");
const balanceName = document.getElementById("balanceName");
const balanceTaxId = document.getElementById("balanceTaxId");
const balanceSaveBtn = document.getElementById("balanceSaveBtn");
const balanceSaveStatus = document.getElementById("balanceSaveStatus");
const balanceMissingInfo = document.getElementById("balanceMissingInfo");
const bsAssetIntangible = document.getElementById("bsAssetIntangible");
const bsAssetTangible = document.getElementById("bsAssetTangible");
const bsAssetInvestmentProperty = document.getElementById("bsAssetInvestmentProperty");
const bsAssetLongTermInv = document.getElementById("bsAssetLongTermInv");
const bsAssetDeferredTax = document.getElementById("bsAssetDeferredTax");
const bsAssetInventory = document.getElementById("bsAssetInventory");
const bsAssetReceivables = document.getElementById("bsAssetReceivables");
const bsAssetShortTermInv = document.getElementById("bsAssetShortTermInv");
const bsAssetCash = document.getElementById("bsAssetCash");
const bsAssetCurrentTax = document.getElementById("bsAssetCurrentTax");
const bsEquityCapital = document.getElementById("bsEquityCapital");
const bsEquityReserves = document.getElementById("bsEquityReserves");
const bsEquityResult = document.getElementById("bsEquityResult");
const bsEquityGrants = document.getElementById("bsEquityGrants");
const bsLiabLongTermDebt = document.getElementById("bsLiabLongTermDebt");
const bsLiabLongTermOther = document.getElementById("bsLiabLongTermOther");
const bsLiabShortTermDebt = document.getElementById("bsLiabShortTermDebt");
const bsLiabPayables = document.getElementById("bsLiabPayables");
const bsLiabOtherCurrent = document.getElementById("bsLiabOtherCurrent");
const bsAssetNonCurrentTotal = document.getElementById("bsAssetNonCurrentTotal");
const bsAssetCurrentTotal = document.getElementById("bsAssetCurrentTotal");
const bsEquityTotal = document.getElementById("bsEquityTotal");
const bsLiabNonCurrentTotal = document.getElementById("bsLiabNonCurrentTotal");
const bsLiabCurrentTotal = document.getElementById("bsLiabCurrentTotal");
const bsTotalAssets = document.getElementById("bsTotalAssets");
const bsTotalLiabilities = document.getElementById("bsTotalLiabilities");
const globalProcessing = document.getElementById("globalProcessing");
const globalProcessingText = document.getElementById("globalProcessingText");
const companyDisplayName = document.getElementById("companyDisplayName");
const companyLegalName = document.getElementById("companyLegalName");
const companyTaxId = document.getElementById("companyTaxId");
const companyType = document.getElementById("companyType");
const companyAssignedSelect = document.getElementById("companyAssignedSelect");
const companyEmail = document.getElementById("companyEmail");
const companyPhone = document.getElementById("companyPhone");
const companyVatRegime = document.getElementById("companyVatRegime");
const companyTaxPeriodicity = document.getElementById("companyTaxPeriodicity");
const companyFilesModel111 = document.getElementById("companyFilesModel111");
const companyFilesModel115 = document.getElementById("companyFilesModel115");
const companyFilesModel130 = document.getElementById("companyFilesModel130");
const companyFilesModel202 = document.getElementById("companyFilesModel202");
const companySaveBtn = document.getElementById("companySaveBtn");
const companiesTableBody = document.querySelector("#companiesTable tbody");
const companiesEmpty = document.getElementById("companiesEmpty");
const accountAgencyName = document.getElementById("accountAgencyName");
const accountEmail = document.getElementById("accountEmail");
const accountPhone = document.getElementById("accountPhone");
const accountCurrentPassword = document.getElementById("accountCurrentPassword");
const accountNewPassword = document.getElementById("accountNewPassword");
const accountSaveBtn = document.getElementById("accountSaveBtn");
const staffEmail = document.getElementById("staffEmail");
const staffSaveBtn = document.getElementById("staffSaveBtn");
const staffTableBody = document.querySelector("#staffTable tbody");
const staffEmpty = document.getElementById("staffEmpty");
const headerCompanyLabel = document.getElementById("headerCompanyLabel");
const headerPeriodLabel = document.getElementById("headerPeriodLabel");
const headerUserEmail = document.getElementById("headerUserEmail");
const incomeUploadBtn = document.getElementById("incomeUploadBtn");
const incomeDropZone = document.getElementById("incomeDropZone");
const incomeFileInput = document.getElementById("incomeFileInput");
const incomeFolderInput = document.getElementById("incomeFolderInput");
const incomeUploadTableBody = document.querySelector("#incomeUploadTable tbody");
const incomeEmptyMessage = document.getElementById("incomeEmptyMessage");
const incomeInvoicesTableBody = document.querySelector("#incomeInvoicesTable tbody");
const incomeInvoicesEmpty = document.getElementById("incomeInvoicesEmpty");
const incomeSearchInput = document.getElementById("incomeSearchInput");
const archiveSearchInput = document.getElementById("archiveSearchInput");
const archiveTypeFilter = document.getElementById("archiveTypeFilter");
const archiveGroupBy = document.getElementById("archiveGroupBy");
const archiveMonthFilter = document.getElementById("archiveMonthFilter");
const archiveStatusFilter = document.getElementById("archiveStatusFilter");
const archiveExportBtn = document.getElementById("archiveExportBtn");
const archiveTableBody = document.querySelector("#archiveTable tbody");
const archiveEmpty = document.getElementById("archiveEmpty");
const archiveSummaryCount = document.getElementById("archiveSummaryCount");
const archiveSummaryBase = document.getElementById("archiveSummaryBase");
const archiveSummaryVat = document.getElementById("archiveSummaryVat");
const archiveSummaryTotal = document.getElementById("archiveSummaryTotal");
const reportYearSelect = document.getElementById("reportYearSelect");
const reportQuarterSelect = document.getElementById("reportQuarterSelect");
const reportStartMonthSelect = document.getElementById("reportStartMonthSelect");
const reportEndMonthSelect = document.getElementById("reportEndMonthSelect");
const reportDownloadBtn = document.getElementById("reportDownloadBtn");
const reportEmailBtn = document.getElementById("reportEmailBtn");
const reportStatus = document.getElementById("reportStatus");
const integrationStartDate = document.getElementById("integrationStartDate");
const integrationEndDate = document.getElementById("integrationEndDate");
const integrationFormat = document.getElementById("integrationFormat");
const integrationStatus = document.getElementById("integrationStatus");
const integrationPurchasesCount = document.getElementById("integrationPurchasesCount");
const integrationSalesCount = document.getElementById("integrationSalesCount");
const integrationJournalCount = document.getElementById("integrationJournalCount");
const integrationDocumentsCount = document.getElementById("integrationDocumentsCount");
const integrationExportPurchasesBtn = document.getElementById("integrationExportPurchasesBtn");
const integrationExportSalesBtn = document.getElementById("integrationExportSalesBtn");
const integrationExportJournalBtn = document.getElementById("integrationExportJournalBtn");
const integrationExportPackageBtn = document.getElementById("integrationExportPackageBtn");
const fiscalModelsTableBody = document.getElementById("fiscalModelsTableBody");
const documentCenterPeriod = document.getElementById("documentCenterPeriod");
const documentCenterFiles = document.getElementById("documentCenterFiles");
const documentCenterUploadBtn = document.getElementById("documentCenterUploadBtn");
const documentCenterStatus = document.getElementById("documentCenterStatus");
const documentCenterBatchFilter = document.getElementById("documentCenterBatchFilter");
const documentCenterStatusFilter = document.getElementById("documentCenterStatusFilter");
const documentCenterBatchCards = document.getElementById("documentCenterBatchCards");
const documentCenterRegisterReadyBtn = document.getElementById("documentCenterRegisterReadyBtn");
const documentCenterTableBody = document.querySelector("#documentCenterTable tbody");
const documentCenterEmpty = document.getElementById("documentCenterEmpty");
const documentCenterDetail = document.getElementById("documentCenterDetail");
const documentCenterDetailEmpty = document.getElementById("documentCenterDetailEmpty");
const documentCenterDetailTitle = document.getElementById("documentCenterDetailTitle");
const documentCenterDetailMeta = document.getElementById("documentCenterDetailMeta");
const documentCenterOpenFile = document.getElementById("documentCenterOpenFile");
const documentCenterDocType = document.getElementById("documentCenterDocType");
const documentCenterCounterparty = document.getElementById("documentCenterCounterparty");
const documentCenterTaxId = document.getElementById("documentCenterTaxId");
const documentCenterInvoiceNumber = document.getElementById("documentCenterInvoiceNumber");
const documentCenterInvoiceDate = document.getElementById("documentCenterInvoiceDate");
const documentCenterPaymentDate = document.getElementById("documentCenterPaymentDate");
const documentCenterBaseAmount = document.getElementById("documentCenterBaseAmount");
const documentCenterVatAmount = document.getElementById("documentCenterVatAmount");
const documentCenterTotalAmount = document.getElementById("documentCenterTotalAmount");
const documentCenterConcept = document.getElementById("documentCenterConcept");
const documentCenterPayrollFields = document.getElementById("documentCenterPayrollFields");
const documentCenterPayrollEmployee = document.getElementById("documentCenterPayrollEmployee");
const documentCenterPayrollPeriod = document.getElementById("documentCenterPayrollPeriod");
const documentCenterPayrollGross = document.getElementById("documentCenterPayrollGross");
const documentCenterPayrollNet = document.getElementById("documentCenterPayrollNet");
const documentCenterPayrollDeductions = document.getElementById("documentCenterPayrollDeductions");
const documentCenterPayrollEmployerCost = document.getElementById("documentCenterPayrollEmployerCost");
const documentCenterTaxModelFields = document.getElementById("documentCenterTaxModelFields");
const documentCenterTaxModelName = document.getElementById("documentCenterTaxModelName");
const documentCenterTaxModelStatus = document.getElementById("documentCenterTaxModelStatus");
const documentCenterTaxModelPeriod = document.getElementById("documentCenterTaxModelPeriod");
const documentCenterTaxModelAmount = document.getElementById("documentCenterTaxModelAmount");
const documentCenterTaxModelOffsetAmount = document.getElementById("documentCenterTaxModelOffsetAmount");
const documentCenterTaxModelRefundAmount = document.getElementById("documentCenterTaxModelRefundAmount");
const documentCenterSaveBtn = document.getElementById("documentCenterSaveBtn");
const documentCenterApproveBtn = document.getElementById("documentCenterApproveBtn");
const documentCenterRegisterBtn = document.getElementById("documentCenterRegisterBtn");
const documentCenterRejectBtn = document.getElementById("documentCenterRejectBtn");
const documentCenterExtractedJson = document.getElementById("documentCenterExtractedJson");
const documentCenterAuditLog = document.getElementById("documentCenterAuditLog");
const currentUserRole = document.body ? document.body.dataset.userRole : null;
const paymentCalendar = document.getElementById("paymentCalendar");
const paymentCalendarTitle = document.getElementById("paymentCalendarTitle");
const paymentPrevMonth = document.getElementById("paymentPrevMonth");
const paymentNextMonth = document.getElementById("paymentNextMonth");
const paymentMonthTotal = document.getElementById("paymentMonthTotal");
const paymentMonthEmpty = document.getElementById("paymentMonthEmpty");
const paymentDueAlert = document.getElementById("paymentDueAlert");
const paymentDueAlertSummary = document.getElementById("paymentDueAlertSummary");
const paymentDueAlertList = document.getElementById("paymentDueAlertList");
const paymentClearAllBtn = document.getElementById("paymentClearAllBtn");
const paymentDayTitle = document.getElementById("paymentDayTitle");
const paymentDayList = document.getElementById("paymentDayList");
const paymentDayTotal = document.getElementById("paymentDayTotal");
const expenseSectionTitle = document.getElementById("expenseSectionTitle");
const expenseSectionDescription = document.getElementById("expenseSectionDescription");
const loanPanel = document.getElementById("loanPanel");
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
const balanceInputIds = [
  "bsAssetIntangible",
  "bsAssetTangible",
  "bsAssetInvestmentProperty",
  "bsAssetLongTermInv",
  "bsAssetDeferredTax",
  "bsAssetInventory",
  "bsAssetReceivables",
  "bsAssetShortTermInv",
  "bsAssetCash",
  "bsAssetCurrentTax",
  "bsEquityCapital",
  "bsEquityReserves",
  "bsEquityResult",
  "bsEquityGrants",
  "bsLiabLongTermDebt",
  "bsLiabLongTermOther",
  "bsLiabShortTermDebt",
  "bsLiabPayables",
  "bsLiabOtherCurrent",
];
const balanceFieldLabels = {
  bsAssetIntangible: "Inmovilizado intangible",
  bsAssetTangible: "Inmovilizado material",
  bsAssetInvestmentProperty: "Inversiones inmobiliarias",
  bsAssetLongTermInv: "Inversiones financieras a largo plazo",
  bsAssetDeferredTax: "Activos por impuesto diferido",
  bsAssetInventory: "Existencias",
  bsAssetReceivables: "Deudores comerciales y otras cuentas a cobrar",
  bsAssetShortTermInv: "Inversiones financieras a corto plazo",
  bsAssetCash: "Efectivo y otros activos líquidos equivalentes",
  bsAssetCurrentTax: "Activos por impuesto corriente",
  bsEquityCapital: "Capital",
  bsEquityReserves: "Reservas",
  bsEquityResult: "Resultado del ejercicio",
  bsEquityGrants: "Subvenciones, donaciones y legados",
  bsLiabLongTermDebt: "Deudas a largo plazo",
  bsLiabLongTermOther: "Otras obligaciones a largo plazo",
  bsLiabShortTermDebt: "Deudas a corto plazo",
  bsLiabPayables: "Proveedores y otras cuentas a pagar",
  bsLiabOtherCurrent: "Otras obligaciones corrientes",
};
const balanceAlwaysManualFieldIds = [
  "bsEquityCapital",
  "bsEquityReserves",
];

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

function normalizeBreakdownLine(line, source = "auto") {
  const baseValue = parseNumberInput(line.base);
  const vatValue = parseNumberInput(line.vat_amount);
  const totalValue = parseNumberInput(line.total);
  const rateValue = normalizeVatRateValue(line.rate);

  if (baseValue === null && totalValue === null) {
    return null;
  }

  let base = baseValue;
  let vatAmount = vatValue;
  let total = totalValue;
  let rate = rateValue !== null ? Number(rateValue) : null;

  if (source === "base" || source === "rate") {
    if (base !== null && rate !== null) {
      vatAmount = roundAmount(base * (rate / 100));
      total = roundAmount(base + vatAmount);
    } else if (base !== null && vatAmount !== null) {
      total = roundAmount(base + vatAmount);
    } else if (base !== null && total !== null) {
      vatAmount = roundAmount(total - base);
    }
  } else if (source === "vat" || source === "vatAmount") {
    if (base !== null && vatAmount !== null) {
      total = roundAmount(base + vatAmount);
      if (base > 0) {
        rate = roundAmount((vatAmount / base) * 100);
      }
    } else if (total !== null && vatAmount !== null) {
      base = roundAmount(total - vatAmount);
      if (base > 0) {
        rate = roundAmount((vatAmount / base) * 100);
      }
    }
  } else if (source === "total") {
    if (base !== null && total !== null) {
      vatAmount = roundAmount(total - base);
      if (base > 0) {
        rate = roundAmount((vatAmount / base) * 100);
      }
    } else if (vatAmount !== null && total !== null) {
      base = roundAmount(total - vatAmount);
      if (base > 0) {
        rate = roundAmount((vatAmount / base) * 100);
      }
    }
  }

  if (base !== null && vatAmount === null && total !== null) {
    vatAmount = roundAmount(total - base);
  }
  if (base !== null && vatAmount !== null && total === null) {
    total = roundAmount(base + vatAmount);
  }
  if (base === null && total !== null && vatAmount !== null) {
    base = roundAmount(total - vatAmount);
  }
  if (rate === null && base !== null && vatAmount !== null && base > 0) {
    rate = roundAmount((vatAmount / base) * 100);
  }
  if (rate !== null && base !== null && vatAmount === null) {
    vatAmount = roundAmount(base * (rate / 100));
    if (total === null) {
      total = roundAmount(base + vatAmount);
    }
  }

  if (base === null || vatAmount === null || total === null) {
    return null;
  }

  return {
    rate,
    base: roundAmount(base),
    vat_amount: roundAmount(vatAmount),
    total: roundAmount(total),
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
  const rates = new Set(
    normalized.map((line) => line.rate).filter((rate) => rate !== null && rate !== undefined)
  );
  if (rates.size !== 1) {
    return null;
  }
  return String([...rates][0]);
}

function getVatDisplayFromInvoice(invoice) {
  const breakdown = parseVatBreakdown(invoice.vat_breakdown || invoice.vatBreakdown);
  const rates = getBreakdownRates(breakdown);
  const suffix = invoice?.vat_deductible === false ? " · no ded." : "";
  if (rates.length > 1) {
    return {
      label: `Mixto${suffix}`,
      title: `IVA: ${rates.map((rate) => formatPercent(rate).replace(" ", "")).join(" · ")}${invoice?.vat_deductible === false ? " · IVA no deducible" : ""}`,
    };
  }
  if (rates.length === 1) {
    return {
      label: `${formatPercent(rates[0])}${suffix}`,
      title: invoice?.vat_deductible === false ? "IVA no deducible" : "",
    };
  }
  return {
    label: `${formatPercent(invoice.vat_rate)}${suffix}`,
    title: invoice?.vat_deductible === false ? "IVA no deducible" : "",
  };
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

function normalizeSearchText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function recordMatchesSearch(parts, searchTerm) {
  const needle = normalizeSearchText(searchTerm);
  if (!needle) {
    return true;
  }
  const haystack = normalizeSearchText(
    parts
      .filter((part) => part !== null && part !== undefined)
      .join(" ")
  );
  return haystack.includes(needle);
}

function formatArchiveRecordType(type) {
  const labels = {
    expense_invoice: "Factura proveedor",
    payroll: "Personal",
    rent: "Alquiler",
    loan: "Financiación",
    adjustment: "Otro ajuste",
    income_invoice: "Factura emitida",
    manual_income: "Facturación manual",
  };
  return labels[type] || type;
}

function formatArchiveMonthLabel(dateValue, monthValue = null, yearValue = null) {
  if (dateValue) {
    const raw = String(dateValue);
    const year = raw.slice(0, 4);
    const month = Number(raw.slice(5, 7));
    if (year && month) {
      return `${monthNames[month - 1]} ${year}`;
    }
  }
  if (monthValue && yearValue) {
    return `${monthNames[Number(monthValue) - 1]} ${yearValue}`;
  }
  return "Sin periodo";
}

function getArchiveRecordStatus(paymentDates, completedDates) {
  const dates = normalizePaymentDateList(paymentDates);
  if (!dates.length) {
    return "untracked";
  }
  const completed = new Set(normalizePaymentDateList(completedDates));
  if (dates.every((value) => completed.has(value))) {
    return "paid";
  }
  const todayIso = new Date().toISOString().slice(0, 10);
  const pendingDates = dates.filter((value) => !completed.has(value));
  if (pendingDates.some((value) => value < todayIso)) {
    return "overdue";
  }
  return "pending";
}

function archiveStatusLabel(status) {
  const labels = {
    pending: "Pendiente",
    paid: "Pagado",
    overdue: "Atrasado",
    untracked: "Sin seguimiento",
  };
  return labels[status] || "Sin seguimiento";
}

function getArchiveGroupStatusMeta(records) {
  const counts = {
    pending: 0,
    paid: 0,
    overdue: 0,
    untracked: 0,
  };
  (records || []).forEach((record) => {
    const status = counts[record.status] !== undefined ? record.status : "untracked";
    counts[status] += 1;
  });
  const activeStatuses = Object.entries(counts).filter(([, count]) => count > 0);
  if (!activeStatuses.length) {
    return {
      key: "untracked",
      label: archiveStatusLabel("untracked"),
      breakdown: "",
    };
  }
  if (activeStatuses.length === 1) {
    const [status, count] = activeStatuses[0];
    return {
      key: status,
      label: archiveStatusLabel(status),
      breakdown: `${count} registro(s)`,
    };
  }
  if (counts.overdue > 0) {
    return {
      key: "overdue",
      label: "Mixto con atrasos",
      breakdown: `Atrasados: ${counts.overdue}, pendientes: ${counts.pending}, pagados: ${counts.paid}, sin seguimiento: ${counts.untracked}`,
    };
  }
  if (counts.pending > 0) {
    return {
      key: "pending",
      label: "Mixto pendiente",
      breakdown: `Pendientes: ${counts.pending}, pagados: ${counts.paid}, sin seguimiento: ${counts.untracked}`,
    };
  }
  if (counts.paid > 0) {
    return {
      key: "paid",
      label: "Mixto cerrado",
      breakdown: `Pagados: ${counts.paid}, sin seguimiento: ${counts.untracked}`,
    };
  }
  return {
    key: "untracked",
    label: "Mixto sin seguimiento",
    breakdown: `Sin seguimiento: ${counts.untracked}`,
  };
}

function buildArchiveCsv(records) {
  const rows = [
    ["Tipo", "Estado", "Tercero", "Concepto", "Fecha", "Mes", "Base", "IVA", "Total"],
    ...records.map((record) => [
      formatArchiveRecordType(record.type),
      archiveStatusLabel(record.status),
      record.counterparty || "",
      record.concept || "",
      record.date || "",
      record.monthLabel || "",
      (Number(record.base) || 0).toFixed(2),
      (Number(record.vat) || 0).toFixed(2),
      (Number(record.total) || 0).toFixed(2),
    ]),
  ];
  return rows
    .map((row) =>
      row
        .map((value) => `"${String(value ?? "").replace(/"/g, '""')}"`)
        .join(";")
    )
    .join("\n");
}

function exportArchiveCsv() {
  const records = getArchiveFilteredRecords();
  if (!records.length) {
    alert("No hay registros en el archivo para exportar.");
    return;
  }
  const csv = buildArchiveCsv(records);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const href = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const { month, year } = getSelectedMonthYear();
  link.href = href;
  link.download = `archivo_ledged_${year || "periodo"}_${month || "00"}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(href);
}

function syncArchiveMonthFilter() {
  if (!archiveMonthFilter) {
    return;
  }
  const selectedValue = archiveMonthFilter.value || "all";
  const records = buildArchiveRecords();
  const labels = Array.from(new Set(records.map((record) => record.monthLabel).filter(Boolean))).sort(
    (a, b) => a.localeCompare(b, "es")
  );
  archiveMonthFilter.innerHTML = '<option value="all">Todos</option>';
  labels.forEach((label) => {
    const option = document.createElement("option");
    option.value = label;
    option.textContent = label;
    archiveMonthFilter.appendChild(option);
  });
  archiveMonthFilter.value = labels.includes(selectedValue) ? selectedValue : "all";
}

function buildArchiveRecords() {
  const expenseInvoices = (currentInvoices || []).map((invoice) => ({
    type: "expense_invoice",
    counterparty: invoice.supplier || "Proveedor pendiente",
    concept: invoice.original_filename || invoice.supplier || "Factura proveedor",
    monthLabel: formatArchiveMonthLabel(invoice.invoice_date),
    date: invoice.invoice_date || "",
    base: Number(invoice.base_amount) || 0,
    vat: Number(invoice.vat_amount) || 0,
    total: Number(invoice.total_amount) || 0,
    searchKey: invoice.supplier || "",
    subview: "received",
    status: getArchiveRecordStatus(
      getInvoicePaymentDates(invoice),
      invoice.payment_completed_dates
    ),
  }));
  const noInvoiceRecords = (currentNoInvoiceExpenses || []).map((expense) => {
    let type = "adjustment";
    if (["nomina", "seguridad_social"].includes(expense.expense_type)) {
      type = "payroll";
    } else if (["alquiler_local", "alquiler_cabina"].includes(expense.expense_type)) {
      type = "rent";
    } else if (expense.expense_type === "prestamo") {
      type = "loan";
    }
    return {
      type,
      counterparty:
        expense.payroll_employee_name ||
        expense.concept ||
        formatNoInvoiceType(expense.expense_type),
      concept: expense.concept || formatNoInvoiceType(expense.expense_type),
      monthLabel: formatArchiveMonthLabel(expense.expense_date),
      date: expense.expense_date || "",
      base: Number(
        expense.base_amount_no_invoice ?? expense.base_amount ?? expense.amount
      ) || 0,
      vat: Number(expense.vat_amount_no_invoice ?? expense.vat_amount) || 0,
      total: Number(expense.amount) || 0,
      searchKey: expense.payroll_employee_name || expense.concept || "",
      subview:
        type === "payroll"
          ? "payroll"
          : type === "rent"
            ? "rent"
            : type === "loan"
              ? "financing"
              : "adjustments",
      status: getArchiveRecordStatus(
        getNoInvoicePaymentDates(expense),
        expense.payment_completed_dates
      ),
    };
  });
  const loanRecords = (currentLoanInstallments || []).map((installment) => ({
    type: "loan",
    counterparty: installment.bank_name || installment.concept || "Entidad financiera",
    concept: installment.concept || "Cuota préstamo",
    monthLabel: formatArchiveMonthLabel(installment.payment_date),
    date: installment.payment_date || "",
    base: Number(installment.principal_amount) || 0,
    vat: 0,
    total: Number(installment.total_amount) || 0,
    searchKey: installment.bank_name || installment.concept || "",
    subview: "financing",
    status: getArchiveRecordStatus([installment.payment_date], installment.payment_completed_dates),
  }));
  const incomeInvoiceRecords = (currentIncomeInvoices || []).map((invoice) => ({
    type: "income_invoice",
    counterparty: invoice.client || "Cliente pendiente",
    concept: invoice.original_filename || invoice.client || "Factura emitida",
    monthLabel: formatArchiveMonthLabel(invoice.invoice_date),
    date: invoice.invoice_date || "",
    base: Number(invoice.base_amount) || 0,
    vat: Number(invoice.vat_amount) || 0,
    total: Number(invoice.total_amount) || 0,
    searchKey: invoice.client || "",
    section: "income",
    status: getArchiveRecordStatus(
      getIncomeInvoicePaymentDates(invoice),
      invoice.payment_completed_dates
    ),
  }));
  const manualIncomeRecords = (currentBillingEntries || []).map((entry) => ({
    type: "manual_income",
    counterparty: entry.concept || "Facturación manual",
    concept: entry.concept || "Facturación manual",
    monthLabel: formatArchiveMonthLabel(entry.invoice_date, entry.month, entry.year),
    date: entry.invoice_date || `${entry.year}-${String(entry.month).padStart(2, "0")}-01`,
    base: Number(entry.base) || 0,
    vat: Number(entry.vatAmount) || 0,
    total: Number(entry.total) || 0,
    searchKey: entry.concept || "",
    section: "income",
    status: "untracked",
  }));
  currentArchiveRecords = [
    ...expenseInvoices,
    ...noInvoiceRecords,
    ...loanRecords,
    ...incomeInvoiceRecords,
    ...manualIncomeRecords,
  ];
  return currentArchiveRecords;
}

function getArchiveFilteredRecords() {
  const searchTerm = archiveSearchInput?.value || "";
  const selectedType = archiveTypeFilter?.value || "all";
  const selectedMonth = archiveMonthFilter?.value || "all";
  const selectedStatus = archiveStatusFilter?.value || "all";
  const source = buildArchiveRecords();
  return source.filter((record) => {
    if (selectedType !== "all" && record.type !== selectedType) {
      return false;
    }
    if (selectedMonth !== "all" && record.monthLabel !== selectedMonth) {
      return false;
    }
    if (selectedStatus !== "all" && record.status !== selectedStatus) {
      return false;
    }
    return recordMatchesSearch(
      [
        record.counterparty,
        record.concept,
        record.monthLabel,
        record.date,
        record.base,
        record.vat,
        record.total,
        formatArchiveRecordType(record.type),
      ],
      searchTerm
    );
  });
}

function applyArchiveDrilldown(group) {
  if (!group || !group.records?.length) {
    return;
  }
  const targetRecord = group.records[0];
  if (targetRecord.type === "income_invoice") {
    setActiveSection("income");
    if (incomeSearchInput) {
      incomeSearchInput.value = group.searchValue || targetRecord.searchKey || group.label;
      renderIncomeInvoices(currentIncomeInvoices || []);
    }
    return;
  }
  if (targetRecord.type === "manual_income") {
    setActiveSection("income");
    if (billingSearchInput) {
      billingSearchInput.value = group.searchValue || targetRecord.searchKey || group.label;
      renderBillingEntries(currentBillingEntries || []);
    }
    return;
  }
  setActiveSection("expenses");
  setExpenseSubview(targetRecord.subview || "adjustments");
  if (targetRecord.type === "expense_invoice") {
    if (invoiceSearchInput) {
      invoiceSearchInput.value = group.searchValue || targetRecord.searchKey || group.label;
      renderInvoices(currentInvoices || []);
    }
    return;
  }
  if (targetRecord.type === "loan") {
    if (loanSearchInput) {
      loanSearchInput.value = group.searchValue || targetRecord.searchKey || group.label;
      renderLoanInstallments(currentLoanInstallments || []);
    }
    return;
  }
  if (noInvoiceSearchInput) {
    noInvoiceSearchInput.value = group.searchValue || targetRecord.searchKey || group.label;
    renderNoInvoiceExpenses(currentNoInvoiceExpenses || []);
  }
}

function renderArchive() {
  if (!archiveTableBody || !archiveEmpty) {
    return;
  }
  syncArchiveMonthFilter();
  const filteredRecords = getArchiveFilteredRecords();
  const groupMode = archiveGroupBy?.value || "counterparty";
  const grouped = new Map();

  filteredRecords.forEach((record) => {
    const groupLabel =
      groupMode === "month"
        ? record.monthLabel
        : groupMode === "type"
          ? formatArchiveRecordType(record.type)
          : record.counterparty || "Sin tercero";
    const groupKey =
      groupMode === "type"
        ? record.type
        : normalizeSearchText(groupLabel);
    if (!grouped.has(groupKey)) {
      grouped.set(groupKey, {
        label: groupLabel,
        type: record.type,
        types: new Set(),
        months: new Set(),
        base: 0,
        vat: 0,
        total: 0,
        count: 0,
        records: [],
        searchValue:
          groupMode === "month"
            ? String(record.date || "").slice(0, 7)
            : groupMode === "type"
              ? ""
              : record.searchKey || groupLabel,
      });
    }
    const bucket = grouped.get(groupKey);
    bucket.types.add(formatArchiveRecordType(record.type));
    bucket.months.add(record.monthLabel);
    bucket.base += Number(record.base) || 0;
    bucket.vat += Number(record.vat) || 0;
    bucket.total += Number(record.total) || 0;
    bucket.count += 1;
    bucket.records.push(record);
  });

  const rows = Array.from(grouped.values()).sort((a, b) => a.label.localeCompare(b.label, "es"));
  archiveTableBody.innerHTML = "";
  archiveEmpty.style.display = rows.length ? "none" : "block";

  if (archiveSummaryCount) {
    archiveSummaryCount.textContent = String(filteredRecords.length);
  }
  if (archiveSummaryBase) {
    archiveSummaryBase.textContent = formatCurrency(
      filteredRecords.reduce((sum, item) => sum + (Number(item.base) || 0), 0)
    );
  }
  if (archiveSummaryVat) {
    archiveSummaryVat.textContent = formatCurrency(
      filteredRecords.reduce((sum, item) => sum + (Number(item.vat) || 0), 0)
    );
  }
  if (archiveSummaryTotal) {
    archiveSummaryTotal.textContent = formatCurrency(
      filteredRecords.reduce((sum, item) => sum + (Number(item.total) || 0), 0)
    );
  }

  rows.forEach((group) => {
    const tr = document.createElement("tr");

    const labelTd = document.createElement("td");
    labelTd.textContent = group.label;

    const typeTd = document.createElement("td");
    typeTd.textContent = Array.from(group.types).join(", ");

    const monthsTd = document.createElement("td");
    monthsTd.textContent = Array.from(group.months).join(", ");

    const statusTd = document.createElement("td");
    const statusMeta = getArchiveGroupStatusMeta(group.records);
    const statusPill = document.createElement("span");
    statusPill.className = `archive-status-pill ${statusMeta.key}`;
    statusPill.textContent = statusMeta.label;
    if (statusMeta.breakdown) {
      statusPill.title = statusMeta.breakdown;
    }
    statusTd.appendChild(statusPill);

    const countTd = document.createElement("td");
    countTd.textContent = String(group.count);

    const baseTd = document.createElement("td");
    baseTd.textContent = formatCurrency(group.base);

    const vatTd = document.createElement("td");
    vatTd.textContent = formatCurrency(group.vat);

    const totalTd = document.createElement("td");
    totalTd.textContent = formatCurrency(group.total);

    const actionsTd = document.createElement("td");
    const actionBtn = document.createElement("button");
    actionBtn.type = "button";
    actionBtn.className = "button ghost archive-link-btn";
    actionBtn.textContent = "Ver detalle";
    actionBtn.addEventListener("click", () => applyArchiveDrilldown(group));
    actionsTd.appendChild(actionBtn);

    tr.appendChild(labelTd);
    tr.appendChild(typeTd);
    tr.appendChild(statusTd);
    tr.appendChild(monthsTd);
    tr.appendChild(countTd);
    tr.appendChild(baseTd);
    tr.appendChild(vatTd);
    tr.appendChild(totalTd);
    tr.appendChild(actionsTd);
    archiveTableBody.appendChild(tr);
  });
}

function createNoInvoiceTypeSelect(selected, allowedTypes = Object.keys(noInvoiceTypeLabels)) {
  const select = document.createElement("select");
  allowedTypes.forEach((key) => {
    const option = document.createElement("option");
    option.value = key;
    option.textContent = noInvoiceTypeLabels[key];
    select.appendChild(option);
  });
  select.value = allowedTypes.includes(selected) ? selected : allowedTypes[0] || "otro";
  return select;
}

function getNoInvoiceDeductibleAmount(expense) {
  if (!expense) {
    return 0;
  }
  if (expense.expense_type === "prestamo") {
    return Number(expense.interest_amount) || 0;
  }
  if (expense.vat_deductible) {
    return (
      Number(expense.base_amount) ||
      Number(expense.base_amount_no_invoice) ||
      0
    );
  }
  if (!expense.deductible) {
    return 0;
  }
  return Number(expense.amount) || 0;
}

function getInvoiceDeductibleAmount(invoice) {
  if (!invoice || invoice.expense_category === "non_deductible") {
    return 0;
  }
  if (invoice.vat_deductible === false) {
    return Number(invoice.total_amount) || 0;
  }
  return Number(invoice.base_amount) || 0;
}

function getInvoiceVatDeductibleAmount(invoice) {
  if (!invoice || invoice.expense_category === "non_deductible" || invoice.vat_deductible === false) {
    return 0;
  }
  return Number(invoice.vat_amount) || 0;
}

function getNoInvoiceVatDeductibleAmount(expense) {
  if (!expense || !expense.vat_deductible) {
    return 0;
  }
  return (
    Number(expense.vat_amount) ||
    Number(expense.vat_amount_no_invoice) ||
    0
  );
}

function toggleLoanInterestField({
  typeValue,
  interestField,
  interestInput,
  deductibleSelect,
}) {
  if (!interestField || !interestInput || !deductibleSelect) {
    return;
  }
  const isLoan = typeValue === "prestamo";
  const canHide = interestField.id === "noInvoiceInterestField";
  if (canHide) {
    interestField.classList.toggle("is-hidden", !isLoan);
  } else {
    interestField.classList.remove("is-hidden");
  }
  interestInput.disabled = !isLoan;
  if (isLoan) {
    deductibleSelect.value = "false";
    deductibleSelect.disabled = true;
  } else {
    interestInput.value = "";
    deductibleSelect.disabled = false;
  }
}

function shouldShowWithholdingField(typeValue) {
  return ["nomina", "alquiler_local", "alquiler_cabina", "otro"].includes(typeValue);
}

function toggleNoInvoiceWithholdingField({ typeValue, field = noInvoiceWithholdingField, input = noInvoiceWithholding }) {
  if (!field || !input) {
    return;
  }
  const visible = shouldShowWithholdingField(typeValue);
  field.classList.toggle("is-hidden", !visible);
  input.disabled = !visible;
  if (!visible) {
    input.value = "";
  }
}

function toggleNoInvoiceVatFields({ typeValue, vatDeductibleValue }) {
  if (
    !noInvoiceVatRateField ||
    !noInvoiceVatBaseField ||
    !noInvoiceVatAmountField ||
    !noInvoiceVatDeductible ||
    !noInvoiceDeductible
  ) {
    return;
  }
  const isLoan = typeValue === "prestamo";
  const blocksVat = isLoan || ["nomina", "seguridad_social"].includes(typeValue);
  const forceDeductible = ["nomina", "seguridad_social"].includes(typeValue);
  if (blocksVat) {
    noInvoiceVatDeductible.value = "false";
    noInvoiceVatDeductible.disabled = true;
  } else {
    noInvoiceVatDeductible.disabled = false;
  }
  const isVatDeductible =
    (vatDeductibleValue || noInvoiceVatDeductible.value) === "true" && !blocksVat;

  noInvoiceVatRateField.classList.toggle("is-hidden", !isVatDeductible);
  noInvoiceVatBaseField.classList.toggle("is-hidden", !isVatDeductible);
  noInvoiceVatAmountField.classList.toggle("is-hidden", !isVatDeductible);

  if (isVatDeductible) {
    noInvoiceDeductible.value = "true";
    noInvoiceDeductible.disabled = true;
  } else {
    if (forceDeductible) {
      noInvoiceDeductible.value = "true";
      noInvoiceDeductible.disabled = true;
    } else {
      noInvoiceDeductible.disabled = isLoan;
    }
    if (!blocksVat) {
      noInvoiceVatBase.value = "";
      noInvoiceVatAmount.value = "";
    } else {
      noInvoiceVatBase.value = "";
      noInvoiceVatAmount.value = "";
    }
  }
}

function syncNoInvoiceVatCalculation() {
  if (
    !noInvoiceAmount ||
    !noInvoiceVatDeductible ||
    !noInvoiceVatSelect ||
    !noInvoiceVatBase ||
    !noInvoiceVatAmount
  ) {
    return;
  }
  if (noInvoiceVatDeductible.value !== "true") {
    noInvoiceVatBase.value = "";
    noInvoiceVatAmount.value = "";
    return;
  }
  const totalValue = parseNumberInput(noInvoiceAmount.value);
  if (totalValue === null) {
    noInvoiceVatBase.value = "";
    noInvoiceVatAmount.value = "";
    return;
  }
  const vatRateValue = resolveVatRateValue(noInvoiceVatSelect.value);
  const result = calculateVatFields({
    baseValue: null,
    totalValue,
    vatRateValue,
    source: "total",
  });
  noInvoiceVatBase.value =
    result.base !== null ? formatAmountInput(result.base) : "";
  noInvoiceVatAmount.value =
    result.vatAmount !== null ? formatAmountInput(result.vatAmount) : "";
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
  renderFiscalModelsSummary();
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
      loadBalanceManualDataFromCompany();
      renderCompaniesTable(companies);
      applyCompanyTaxModules();
      updatePnlSummary();
      updateHeaderContext();
      renderFiscalModelsSummary();
      return companies;
    });
}

function documentStatusLabel(status) {
  const labels = {
    ready_to_register: "Listo",
    needs_review: "Revisión",
    duplicate: "Duplicado",
    registered: "Registrado",
    rejected: "Rechazado",
  };
  return labels[status] || status || "-";
}

function documentTypeLabel(type) {
  const labels = {
    purchase_invoice: "Factura gasto",
    sales_invoice: "Factura emitida",
    receipt: "Ticket",
    payroll: "Nómina",
    social_security: "Seguridad Social",
    loan_document: "Préstamo",
    bank_statement: "Extracto",
    tax_model: "Modelo fiscal",
    unknown: "Sin clasificar",
  };
  return labels[type] || type || "-";
}

function taxFilingStatusLabel(status) {
  const labels = {
    estimated: "Estimado",
    a_ingresar: "Presentado · a ingresar",
    a_compensar: "Presentado · a compensar",
    sin_actividad: "Presentado · sin actividad",
    a_devolver: "Presentado · a devolver",
  };
  return labels[status] || status || "";
}

function toggleDocumentCenterSpecialFields(type) {
  const isPayroll = type === "payroll";
  const isTaxModel = type === "tax_model";
  if (documentCenterPayrollFields) {
    documentCenterPayrollFields.hidden = !isPayroll;
  }
  if (documentCenterTaxModelFields) {
    documentCenterTaxModelFields.hidden = !isTaxModel;
  }
}

function renderDocumentCenterBatches() {
  if (!documentCenterBatchFilter || !documentCenterBatchCards) {
    return;
  }
  documentCenterBatchFilter.innerHTML = '<option value="">Todos los lotes</option>';
  documentCenterBatchCards.innerHTML = "";
  currentDocumentBatches.forEach((batch) => {
    const option = document.createElement("option");
    option.value = String(batch.id);
    option.textContent = `${batch.period} · ${batch.totalDocuments} docs`;
    documentCenterBatchFilter.appendChild(option);

    const card = document.createElement("button");
    card.type = "button";
    card.className = `document-center-batch-card${
      String(selectedDocumentBatchId || "") === String(batch.id) ? " active" : ""
    }`;
    card.innerHTML = `
      <h3>${batch.period}</h3>
      <p>Total: ${batch.totalDocuments}</p>
      <p>Listos: ${batch.readyDocuments} · Revisión: ${batch.reviewDocuments} · Duplicados: ${batch.duplicateDocuments}</p>
      <span class="document-center-status-pill ${batch.status || ""}">${batch.status || "uploaded"}</span>
    `;
    card.addEventListener("click", () => {
      selectedDocumentBatchId = String(batch.id);
      documentCenterBatchFilter.value = selectedDocumentBatchId;
      loadDocumentCenterDocuments();
      renderDocumentCenterBatches();
    });
    documentCenterBatchCards.appendChild(card);
  });
  documentCenterBatchFilter.value = selectedDocumentBatchId || "";
}

function renderDocumentCenterDocuments() {
  if (!documentCenterTableBody || !documentCenterEmpty) {
    return;
  }
  documentCenterTableBody.innerHTML = "";
  if (!currentDocumentCenterDocuments.length) {
    documentCenterEmpty.style.display = "block";
    renderDocumentCenterDetail(null);
    return;
  }
  documentCenterEmpty.style.display = "none";
  currentDocumentCenterDocuments.forEach((doc) => {
    const tr = document.createElement("tr");
    tr.className = `document-center-row${
      String(selectedDocumentCenterDocumentId || "") === String(doc.id) ? " is-active" : ""
    }`;
    tr.innerHTML = `
      <td>${doc.originalFileName || "-"}</td>
      <td>${documentTypeLabel(doc.detectedDocumentType)}</td>
      <td>${doc.counterparty || "-"}</td>
      <td>${doc.totalAmount !== null && doc.totalAmount !== undefined ? formatCurrency(doc.totalAmount) : "-"}</td>
      <td>${doc.confidenceScore !== null && doc.confidenceScore !== undefined ? `${Math.round(doc.confidenceScore)}%` : "-"}</td>
      <td><span class="document-center-status-pill ${doc.validationStatus || ""}">${documentStatusLabel(doc.validationStatus)}</span></td>
    `;
    tr.addEventListener("click", () => {
      selectedDocumentCenterDocumentId = doc.id;
      renderDocumentCenterDocuments();
      renderDocumentCenterDetail(doc);
    });
    documentCenterTableBody.appendChild(tr);
  });
  const selected =
    currentDocumentCenterDocuments.find((doc) => String(doc.id) === String(selectedDocumentCenterDocumentId)) ||
    currentDocumentCenterDocuments[0];
  if (selected) {
    selectedDocumentCenterDocumentId = selected.id;
    renderDocumentCenterDetail(selected);
  }
}

function renderDocumentCenterAudit(auditLog) {
  if (!documentCenterAuditLog) {
    return;
  }
  documentCenterAuditLog.innerHTML = "";
  if (!Array.isArray(auditLog) || !auditLog.length) {
    documentCenterAuditLog.innerHTML = '<div class="document-center-audit-item"><strong>Sin eventos</strong><span>Todavía no hay trazabilidad registrada.</span></div>';
    return;
  }
  auditLog
    .slice()
    .reverse()
    .forEach((entry) => {
      const item = document.createElement("div");
      item.className = "document-center-audit-item";
      item.innerHTML = `
        <strong>${entry.action || "evento"}</strong>
        <span>${entry.timestamp || "-"}</span>
        <span>${entry.payload ? JSON.stringify(entry.payload) : ""}</span>
      `;
      documentCenterAuditLog.appendChild(item);
    });
}

function renderDocumentCenterDetail(doc) {
  if (!documentCenterDetail || !documentCenterDetailEmpty) {
    return;
  }
  if (!doc) {
    documentCenterDetail.hidden = true;
    documentCenterDetailEmpty.style.display = "block";
    toggleDocumentCenterSpecialFields("unknown");
    return;
  }
  documentCenterDetail.hidden = false;
  documentCenterDetailEmpty.style.display = "none";
  const effective = doc.effectiveData || {};
  if (documentCenterDetailTitle) {
    documentCenterDetailTitle.textContent = doc.originalFileName || `Documento #${doc.id}`;
  }
  if (documentCenterDetailMeta) {
    documentCenterDetailMeta.textContent = `${documentTypeLabel(doc.detectedDocumentType)} · ${documentStatusLabel(doc.validationStatus)} · Confianza ${doc.confidenceScore ? Math.round(doc.confidenceScore) : 0}%`;
  }
  if (documentCenterOpenFile) {
    documentCenterOpenFile.href = `/api/document-center/documents/${doc.id}/download?company_id=${getSelectedCompanyId()}`;
  }
  if (documentCenterDocType) {
    documentCenterDocType.value = doc.detectedDocumentType || "unknown";
  }
  toggleDocumentCenterSpecialFields(doc.detectedDocumentType || "unknown");
  if (documentCenterCounterparty) {
    documentCenterCounterparty.value =
      effective.provider_name || effective.client_name || effective.employee_name || effective.bank_name || "";
  }
  if (documentCenterTaxId) {
    documentCenterTaxId.value = effective.tax_id || "";
  }
  if (documentCenterInvoiceNumber) {
    documentCenterInvoiceNumber.value = effective.invoice_number || "";
  }
  if (documentCenterInvoiceDate) {
    documentCenterInvoiceDate.value = effective.invoice_date || "";
  }
  if (documentCenterPaymentDate) {
    documentCenterPaymentDate.value = effective.payment_date || "";
  }
  if (documentCenterBaseAmount) {
    documentCenterBaseAmount.value =
      effective.base_amount !== null && effective.base_amount !== undefined ? String(effective.base_amount) : "";
  }
  if (documentCenterVatAmount) {
    documentCenterVatAmount.value =
      effective.vat_amount !== null && effective.vat_amount !== undefined ? String(effective.vat_amount) : "";
  }
  if (documentCenterTotalAmount) {
    documentCenterTotalAmount.value =
      effective.total_amount !== null && effective.total_amount !== undefined
        ? String(effective.total_amount)
        : effective.amount !== null && effective.amount !== undefined
          ? String(effective.amount)
          : "";
  }
  if (documentCenterConcept) {
    documentCenterConcept.value = effective.concept || effective.filename || "";
  }
  if (documentCenterPayrollEmployee) {
    documentCenterPayrollEmployee.value = effective.employee_name || effective.provider_name || "";
  }
  if (documentCenterPayrollPeriod) {
    documentCenterPayrollPeriod.value = effective.payroll_period || "";
  }
  if (documentCenterPayrollGross) {
    documentCenterPayrollGross.value =
      effective.base_amount !== null && effective.base_amount !== undefined ? String(effective.base_amount) : "";
  }
  if (documentCenterPayrollNet) {
    documentCenterPayrollNet.value =
      effective.payroll_net_amount !== null && effective.payroll_net_amount !== undefined
        ? String(effective.payroll_net_amount)
        : effective.total_amount !== null && effective.total_amount !== undefined
          ? String(effective.total_amount)
          : "";
  }
  if (documentCenterPayrollDeductions) {
    documentCenterPayrollDeductions.value =
      effective.payroll_total_deductions_amount !== null &&
      effective.payroll_total_deductions_amount !== undefined
        ? String(effective.payroll_total_deductions_amount)
        : "";
  }
  if (documentCenterPayrollEmployerCost) {
    documentCenterPayrollEmployerCost.value =
      effective.payroll_employer_cost_amount !== null &&
      effective.payroll_employer_cost_amount !== undefined
        ? String(effective.payroll_employer_cost_amount)
        : "";
  }
  if (documentCenterTaxModelName) {
    documentCenterTaxModelName.value = effective.model_name || "";
  }
  if (documentCenterTaxModelStatus) {
    documentCenterTaxModelStatus.value = effective.filing_status || "";
  }
  if (documentCenterTaxModelPeriod) {
    documentCenterTaxModelPeriod.value = effective.tax_period || "";
  }
  if (documentCenterTaxModelAmount) {
    documentCenterTaxModelAmount.value =
      effective.amount !== null && effective.amount !== undefined
        ? String(effective.amount)
        : effective.total_amount !== null && effective.total_amount !== undefined
          ? String(effective.total_amount)
          : "";
  }
  if (documentCenterTaxModelOffsetAmount) {
    documentCenterTaxModelOffsetAmount.value =
      effective.offset_amount !== null && effective.offset_amount !== undefined
        ? String(effective.offset_amount)
        : "";
  }
  if (documentCenterTaxModelRefundAmount) {
    documentCenterTaxModelRefundAmount.value =
      effective.refund_amount !== null && effective.refund_amount !== undefined
        ? String(effective.refund_amount)
        : "";
  }
  if (documentCenterExtractedJson) {
    documentCenterExtractedJson.textContent = JSON.stringify(doc.originalExtractedData || {}, null, 2);
  }
  renderDocumentCenterAudit(doc.auditLog || []);
}

function loadDocumentCenterBatches() {
  if (!getSelectedCompanyId()) {
    currentDocumentBatches = [];
    renderDocumentCenterBatches();
    return Promise.resolve([]);
  }
  return fetch(withCompanyParam("/api/document-center/batches"))
    .then((res) => res.json())
    .then((data) => {
      currentDocumentBatches = data.batches || [];
      if (
        selectedDocumentBatchId &&
        !currentDocumentBatches.some((batch) => String(batch.id) === String(selectedDocumentBatchId))
      ) {
        selectedDocumentBatchId = "";
      }
      renderDocumentCenterBatches();
      return currentDocumentBatches;
    })
    .catch(() => {
      currentDocumentBatches = [];
      renderDocumentCenterBatches();
      return [];
    });
}

function loadDocumentCenterDocuments() {
  if (!getSelectedCompanyId()) {
    currentDocumentCenterDocuments = [];
    renderDocumentCenterDocuments();
    return Promise.resolve([]);
  }
  const params = new URLSearchParams();
  if (selectedDocumentBatchId) {
    params.set("batch_id", selectedDocumentBatchId);
  }
  if (documentCenterStatusFilter?.value) {
    params.set("validation_status", documentCenterStatusFilter.value);
  }
  const query = params.toString();
  const url = withCompanyParam(`/api/document-center/documents${query ? `?${query}` : ""}`);
  return fetch(url)
    .then((res) => res.json())
    .then((data) => {
      currentDocumentCenterDocuments = data.documents || [];
      if (
        selectedDocumentCenterDocumentId &&
        !currentDocumentCenterDocuments.some(
          (doc) => String(doc.id) === String(selectedDocumentCenterDocumentId)
        )
      ) {
        selectedDocumentCenterDocumentId = null;
      }
      renderDocumentCenterDocuments();
      return currentDocumentCenterDocuments;
    })
    .catch(() => {
      currentDocumentCenterDocuments = [];
      renderDocumentCenterDocuments();
      return [];
    });
}

function buildCurrentDocumentPayload() {
  const currentDoc = currentDocumentCenterDocuments.find(
    (doc) => String(doc.id) === String(selectedDocumentCenterDocumentId)
  );
  if (!currentDoc) {
    return null;
  }
  const source = { ...(currentDoc.effectiveData || {}) };
  const counterparty = documentCenterCounterparty?.value?.trim() || "";
  if (documentCenterDocType?.value === "sales_invoice") {
    source.client_name = counterparty;
    source.client = counterparty;
    delete source.provider_name;
    delete source.supplier;
  } else {
    source.provider_name = counterparty;
    source.supplier = counterparty;
    if (documentCenterDocType?.value === "payroll") {
      source.employee_name = counterparty;
    }
  }
  source.tax_id = documentCenterTaxId?.value?.trim() || null;
  source.invoice_number = documentCenterInvoiceNumber?.value?.trim() || null;
  source.invoice_date = documentCenterInvoiceDate?.value || null;
  source.payment_date = documentCenterPaymentDate?.value || null;
  source.base_amount = parseNumberInput(documentCenterBaseAmount?.value);
  source.vat_amount = parseNumberInput(documentCenterVatAmount?.value);
  source.total_amount = parseNumberInput(documentCenterTotalAmount?.value);
  source.amount = source.total_amount;
  source.concept = documentCenterConcept?.value?.trim() || null;
  const selectedType = documentCenterDocType?.value || currentDoc.detectedDocumentType;
  if (selectedType === "payroll") {
    source.employee_name = documentCenterPayrollEmployee?.value?.trim() || counterparty || null;
    source.provider_name = source.employee_name;
    source.supplier = source.employee_name;
    source.payroll_period = documentCenterPayrollPeriod?.value?.trim() || null;
    source.base_amount = parseNumberInput(documentCenterPayrollGross?.value);
    source.payroll_net_amount = parseNumberInput(documentCenterPayrollNet?.value);
    source.payroll_total_deductions_amount = parseNumberInput(
      documentCenterPayrollDeductions?.value
    );
    source.payroll_employer_cost_amount = parseNumberInput(
      documentCenterPayrollEmployerCost?.value
    );
    source.total_amount = source.payroll_net_amount;
    source.amount = source.payroll_net_amount;
    source.vat_amount = null;
    source.vat_rate = null;
    source.vat_deductible = false;
    source.concept =
      documentCenterConcept?.value?.trim() ||
      [ "Nómina", source.employee_name, source.payroll_period ].filter(Boolean).join(" - ");
  } else if (selectedType === "tax_model") {
    source.model_name = documentCenterTaxModelName?.value?.trim() || null;
    source.filing_status = documentCenterTaxModelStatus?.value || null;
    source.tax_period = documentCenterTaxModelPeriod?.value?.trim() || null;
    source.amount = parseNumberInput(documentCenterTaxModelAmount?.value);
    source.offset_amount = parseNumberInput(documentCenterTaxModelOffsetAmount?.value);
    source.refund_amount = parseNumberInput(documentCenterTaxModelRefundAmount?.value);
    if (source.filing_status === "a_compensar" && source.offset_amount !== null) {
      source.amount = source.offset_amount;
    }
    if (source.filing_status === "a_devolver" && source.refund_amount !== null) {
      source.amount = source.refund_amount;
    }
    if (source.filing_status === "sin_actividad") {
      source.amount = 0;
    }
    source.total_amount = source.amount;
    source.concept =
      documentCenterConcept?.value?.trim() ||
      [
        source.model_name ? `Modelo ${source.model_name}` : "Modelo fiscal",
        source.tax_period,
        source.filing_status ? source.filing_status.replaceAll("_", " ") : "",
      ]
        .filter(Boolean)
        .join(" - ");
  }
  return {
    detected_document_type: selectedType,
    corrected_data: source,
  };
}

function saveDocumentCenterDocument() {
  const payload = buildCurrentDocumentPayload();
  if (!payload || !selectedDocumentCenterDocumentId) {
    return Promise.resolve();
  }
  return fetch(withCompanyParam(`/api/document-center/documents/${selectedDocumentCenterDocumentId}`), {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        throw new Error((data.errors || ["No se pudo guardar el documento."]).join("\n"));
      }
      documentCenterStatus.textContent = "Corrección guardada.";
      return Promise.all([loadDocumentCenterBatches(), loadDocumentCenterDocuments()]);
    })
    .catch((error) => {
      alert(error.message || "No se pudo guardar el documento.");
      throw error;
    });
}

function changeDocumentCenterState(action, extraPayload = {}) {
  if (!selectedDocumentCenterDocumentId) {
    return Promise.resolve();
  }
  return fetch(
    withCompanyParam(`/api/document-center/documents/${selectedDocumentCenterDocumentId}/${action}`),
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(extraPayload),
    }
  )
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        throw new Error((data.errors || ["No se pudo actualizar el documento."]).join("\n"));
      }
      return Promise.all([loadDocumentCenterBatches(), loadDocumentCenterDocuments(), refreshAllData()]);
    })
    .catch((error) => {
      alert(error.message || "No se pudo actualizar el documento.");
    });
}

function registerReadyDocumentCenterBatch() {
  if (!selectedDocumentBatchId) {
    alert("Selecciona un lote antes de registrar.");
    return Promise.resolve();
  }
  return fetch(
    withCompanyParam(`/api/document-center/batches/${selectedDocumentBatchId}/register-ready`),
    { method: "POST" }
  )
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        throw new Error((data.errors || ["No se pudo registrar el lote."]).join("\n"));
      }
      const message = data.errors?.length
        ? `Registrados ${data.registered}. Revisa incidencias pendientes.`
        : `Registrados ${data.registered} documentos del lote.`;
      if (documentCenterStatus) {
        documentCenterStatus.textContent = message;
      }
      return Promise.all([loadDocumentCenterBatches(), loadDocumentCenterDocuments(), refreshAllData()]);
    })
    .catch((error) => {
      alert(error.message || "No se pudo registrar el lote.");
    });
}

function uploadDocumentCenterBatch() {
  if (!getSelectedCompanyId()) {
    alert("Selecciona una empresa antes de subir documentos.");
    return Promise.resolve();
  }
  const files = documentCenterFiles?.files;
  if (!files || !files.length) {
    alert("Selecciona al menos un archivo.");
    return Promise.resolve();
  }
  const formData = new FormData();
  Array.from(files).forEach((file) => formData.append("files", file));
  formData.append(
    "period",
    documentCenterPeriod?.value?.trim() || getPeriodLabel() || `${new Date().getFullYear()}`
  );
  documentCenterUploadBtn.disabled = true;
  if (documentCenterStatus) {
    documentCenterStatus.textContent = "Procesando lote documental...";
  }
  return fetch(withCompanyParam("/api/document-center/upload"), {
    method: "POST",
    body: formData,
  })
    .then(async (res) => {
      const rawBody = await res.text();
      let data = null;
      try {
        data = rawBody ? JSON.parse(rawBody) : null;
      } catch (error) {
        data = null;
      }
      if (!res.ok) {
        const message =
          data?.errors?.join("\n") ||
          "El servidor devolvió un error al procesar el lote documental.";
        throw new Error(message);
      }
      if (!data) {
        throw new Error("Respuesta inválida del servidor al procesar el lote documental.");
      }
      return data;
    })
    .then((data) => {
      if (!data.ok) {
        throw new Error((data.errors || ["No se pudo procesar el lote."]).join("\n"));
      }
      selectedDocumentBatchId = String(data.batch?.id || "");
      selectedDocumentCenterDocumentId = null;
      if (documentCenterFiles) {
        documentCenterFiles.value = "";
      }
      if (documentCenterStatus) {
        documentCenterStatus.textContent = `Lote procesado: ${data.count} documentos.`;
      }
      return Promise.all([loadDocumentCenterBatches(), loadDocumentCenterDocuments()]);
    })
    .catch((error) => {
      alert(error.message || "No se pudo procesar el lote.");
    })
    .finally(() => {
      documentCenterUploadBtn.disabled = false;
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
    const periodicityTd = document.createElement("td");
    periodicityTd.textContent =
      company.tax_periodicity === "monthly" ? "Mensual" : "Trimestral";
    const profileTd = document.createElement("td");
    profileTd.textContent = formatCompanyFiscalProfile(company);
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
    tr.appendChild(periodicityTd);
    tr.appendChild(profileTd);
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
  const periodicityTd = row.children[7];
  const profileTd = row.children[8];
  const actionsTd = row.children[9];

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
  const periodicitySelect = document.createElement("select");
  [
    { value: "quarterly", label: "Trimestral" },
    { value: "monthly", label: "Mensual" },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    periodicitySelect.appendChild(option);
  });
  periodicitySelect.value = company.tax_periodicity || "quarterly";
  const profileText = document.createElement("div");
  profileText.className = "inline-stack";
  const vatRegimeSelect = document.createElement("select");
  [
    { value: "general", label: "IVA general" },
    { value: "prorata", label: "IVA prorrata" },
    { value: "exempt", label: "Exento" },
  ].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.value;
    option.textContent = item.label;
    vatRegimeSelect.appendChild(option);
  });
  vatRegimeSelect.value = company.vat_regime || "general";
  const model111Select = document.createElement("select");
  const model115Select = document.createElement("select");
  const model130Select = document.createElement("select");
  const model202Select = document.createElement("select");
  [model111Select, model115Select, model130Select, model202Select].forEach((select) => {
    const noOption = document.createElement("option");
    noOption.value = "false";
    noOption.textContent = "No";
    const yesOption = document.createElement("option");
    yesOption.value = "true";
    yesOption.textContent = "Sí";
    select.appendChild(noOption);
    select.appendChild(yesOption);
  });
  model111Select.value = company.files_model_111 ? "true" : "false";
  model115Select.value = company.files_model_115 ? "true" : "false";
  model130Select.value = company.files_model_130 ? "true" : "false";
  model202Select.value = company.files_model_202 ? "true" : "false";
  [
    { label: "IVA", node: vatRegimeSelect },
    { label: "111", node: model111Select },
    { label: "115", node: model115Select },
    { label: "130", node: model130Select },
    { label: "202", node: model202Select },
  ].forEach((item) => {
    const row = document.createElement("label");
    row.className = "inline-chip";
    const span = document.createElement("span");
    span.textContent = item.label;
    row.appendChild(span);
    row.appendChild(item.node);
    profileText.appendChild(row);
  });
  const syncInlineFiscalProfile = () => {
    const isCompanyType = typeSelect.value === "company";
    model130Select.disabled = isCompanyType;
    model202Select.disabled = !isCompanyType;
    if (isCompanyType) {
      model130Select.value = "false";
    } else {
      model202Select.value = "false";
    }
  };
  typeSelect.addEventListener("change", syncInlineFiscalProfile);
  syncInlineFiscalProfile();

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
  periodicityTd.textContent = "";
  periodicityTd.appendChild(periodicitySelect);
  profileTd.textContent = "";
  profileTd.appendChild(profileText);

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
      tax_periodicity: periodicitySelect.value,
      vat_regime: vatRegimeSelect.value,
      files_model_303: vatRegimeSelect.value !== "exempt",
      files_model_111: model111Select.value === "true",
      files_model_115: model115Select.value === "true",
      files_model_130: model130Select.value === "true",
      files_model_202: model202Select.value === "true",
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
    !companyAssignedSelect ||
    !companyVatRegime ||
    !companyTaxPeriodicity ||
    !companyFilesModel111 ||
    !companyFilesModel115 ||
    !companyFilesModel130 ||
    !companyFilesModel202
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
      vat_regime: companyVatRegime.value,
      tax_periodicity: companyTaxPeriodicity.value,
      files_model_303: companyVatRegime.value !== "exempt",
      files_model_111: companyFilesModel111.value === "true",
      files_model_115: companyFilesModel115.value === "true",
      files_model_130: companyFilesModel130.value === "true",
      files_model_202: companyFilesModel202.value === "true",
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
      companyVatRegime.value = "general";
      companyTaxPeriodicity.value = "quarterly";
      companyFilesModel111.value = "false";
      companyFilesModel115.value = "false";
      companyFilesModel130.value = "false";
      companyFilesModel202.value = "false";
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
  if (!staffEmail || !staffSaveBtn) {
    return;
  }
  if (!staffEmail.value) {
    alert("El email es obligatorio.");
    return;
  }
  staffSaveBtn.disabled = true;
  fetch("/api/staff/invite", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      email: staffEmail.value,
    }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al enviar la invitación."]).join("\n"));
        return;
      }
      staffEmail.value = "";
      alert("Invitación enviada correctamente.");
    })
    .catch(() => {
      alert("No se pudo enviar la invitación.");
    })
    .finally(() => {
      staffSaveBtn.disabled = false;
    });
}

function saveAccount() {
  if (!accountEmail || !accountSaveBtn) {
    return;
  }
  const emailValue = accountEmail.value.trim();
  if (!emailValue) {
    alert("El email es obligatorio.");
    return;
  }
  accountSaveBtn.disabled = true;
  fetch("/api/account", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      agency_name: accountAgencyName ? accountAgencyName.value : "",
      email: emailValue,
      phone: accountPhone ? accountPhone.value : "",
      current_password: accountCurrentPassword ? accountCurrentPassword.value : "",
      new_password: accountNewPassword ? accountNewPassword.value : "",
    }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["No se pudo guardar la cuenta."]).join("\n"));
        return;
      }
      const account = data.account || {};
      if (accountEmail && account.email) {
        accountEmail.value = account.email;
      }
      if (headerUserEmail && account.email) {
        headerUserEmail.textContent = account.email;
      }
      if (accountAgencyName && typeof account.agency_name === "string") {
        accountAgencyName.value = account.agency_name;
      }
      if (accountPhone && typeof account.phone === "string") {
        accountPhone.value = account.phone;
      }
      if (accountCurrentPassword) {
        accountCurrentPassword.value = "";
      }
      if (accountNewPassword) {
        accountNewPassword.value = "";
      }
      alert("Datos de cuenta actualizados correctamente.");
    })
    .catch(() => {
      alert("No se pudo guardar la cuenta.");
    })
    .finally(() => {
      accountSaveBtn.disabled = false;
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
    if (company) {
      if (pnlName && !pnlName.value) {
        pnlName.value = company.legal_name || company.display_name || "";
      }
      if (pnlTaxId && !pnlTaxId.value) {
        pnlTaxId.value = company.tax_id || "";
      }
      if (balanceName && !balanceName.value) {
        balanceName.value = company.legal_name || company.display_name || "";
      }
      if (balanceTaxId && !balanceTaxId.value) {
        balanceTaxId.value = company.tax_id || "";
      }
    }
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

function getCompanyFiscalModels(company) {
  if (!company) {
    return [];
  }
  const models = [];
  if (company.files_model_303 !== false && company.vat_regime !== "exempt") {
    models.push("303");
    models.push("390");
  }
  if (company.files_model_111) {
    models.push("111");
    models.push("190");
  }
  if (company.files_model_115) {
    models.push("115");
    models.push("180");
  }
  if (company.company_type === "individual" && company.files_model_130) {
    models.push("130");
  }
  if (company.company_type === "company") {
    models.push("200");
    if (company.files_model_202) {
      models.push("202");
    }
  }
  return models;
}

function formatCompanyFiscalProfile(company) {
  const models = getCompanyFiscalModels(company);
  return models.length ? models.join(" · ") : "Perfil pendiente";
}

function getExpenseModeForSubview(subview) {
  if (subview === "rent") {
    return "rent";
  }
  if (subview === "payroll") {
    return "payroll";
  }
  if (subview === "financing") {
    return "financing";
  }
  return "adjustments";
}

function getExpenseUploadKind(subview = currentExpenseSubview) {
  if (subview === "rent") {
    return "rent";
  }
  if (subview === "payroll") {
    return "payroll";
  }
  if (subview === "financing") {
    return "financing";
  }
  if (subview === "adjustments" || subview === "other") {
    return "adjustments";
  }
  return "received";
}

function getNoInvoiceTypeForUploadKind(kind) {
  if (kind === "rent") {
    return "alquiler_local";
  }
  if (kind === "payroll") {
    return "nomina";
  }
  if (kind === "financing") {
    return "prestamo";
  }
  return "otro";
}

function getExpenseUploadConfig(kind = getExpenseUploadKind()) {
  if (kind === "rent") {
    return {
      title: "Subir documentos de alquiler",
      description:
        "Sube facturas o justificantes de alquiler. La IA priorizará arrendador, base, IVA, total y posibles retenciones.",
      buttonLabel: "Guardar alquileres",
    };
  }
  if (kind === "payroll") {
    return {
      title: "Subir nóminas o costes laborales",
      description:
        "Sube nóminas o documentos laborales. La IA priorizará fecha, emisor, importe y retenciones para reflejarlo como gasto de personal.",
      buttonLabel: "Guardar nóminas",
    };
  }
  if (kind === "financing") {
    return {
      title: "Subir financiación y justificantes bancarios",
      description:
        "Sube planes de amortización, recibos o justificantes bancarios. La IA priorizará vencimientos, cuota, interés y entidad financiera.",
      buttonLabel: "Guardar financiación",
    };
  }
  if (kind === "adjustments") {
    return {
      title: "Subir otros gastos y ajustes",
      description:
        "Sube justificantes o documentos de gasto no encajados en proveedor, alquiler o personal. La IA priorizará fecha, importe, base, IVA y naturaleza del ajuste.",
      buttonLabel: "Guardar ajustes",
    };
  }
  return {
    title: "Subir facturas de proveedor",
    description:
      "Arrastra archivos o selecciona una carpeta completa. La pestaña activa indica a la IA qué tipo de gasto debe priorizar.",
    buttonLabel: "Guardar facturas",
  };
}

function getAllowedNoInvoiceTypes(mode = currentExpenseMode) {
  return expenseModeTypes[mode] || expenseModeTypes.adjustments;
}

function populateNoInvoiceTypeSelect(select, allowedTypes, selectedValue) {
  if (!select) {
    return;
  }
  const previousValue = selectedValue || select.value;
  select.innerHTML = "";
  allowedTypes.forEach((key) => {
    const option = document.createElement("option");
    option.value = key;
    option.textContent = noInvoiceTypeLabels[key];
    select.appendChild(option);
  });
  if (allowedTypes.includes(previousValue)) {
    select.value = previousValue;
  } else if (allowedTypes.length) {
    select.value = allowedTypes[0];
  }
}

function updateExpenseUploadPanel(subview = currentExpenseSubview) {
  const config = getExpenseUploadConfig(getExpenseUploadKind(subview));
  if (expenseUploadTitle) {
    expenseUploadTitle.textContent = config.title;
  }
  if (expenseUploadDescription) {
    expenseUploadDescription.textContent = config.description;
  }
  if (uploadBtn) {
    uploadBtn.textContent = config.buttonLabel;
  }
  updateExpenseUploadTableHeaders(getExpenseUploadKind(subview));
}

function updateExpenseUploadTableHeaders(kind = getExpenseUploadKind()) {
  if (!uploadTableHeadCells || uploadTableHeadCells.length < 8) {
    return;
  }
  const labels =
    kind === "payroll"
      ? [
          "Archivo",
          "Fecha",
          "Empleado",
          "Bruto devengado (€)",
          "Total deducciones (€)",
          "Líquido (€)",
          "Periodo",
          "",
        ]
      : ["Archivo", "Fecha", "Proveedor", "Base imponible", "IVA", "IVA (€)", "Total (€)", ""];
  uploadTableHeadCells.forEach((cell, index) => {
    cell.textContent = labels[index] || "";
  });
}

function setExpenseMode(mode) {
  currentExpenseMode = mode;
  const allowedTypes = getAllowedNoInvoiceTypes(mode);
  if (expenseSectionTitle) {
    expenseSectionTitle.textContent =
      mode === "rent"
        ? "Alquileres y arrendamientos"
        : mode === "payroll"
          ? "Personal"
          : mode === "financing"
            ? "Financiación bancaria"
            : "Otros ajustes";
  }
  if (expenseSectionDescription) {
    expenseSectionDescription.textContent =
      mode === "rent"
        ? "Registra alquileres del local o de cabinas con retención, IVA y fechas de pago."
        : mode === "payroll"
          ? "Centraliza nóminas, Seguridad Social y otros costes laborales con fecha contable y vencimiento."
          : mode === "financing"
            ? "Registra préstamos, cuotas e intereses separando correctamente tesorería, gasto financiero y principal."
            : "Registra amortizaciones, kilometraje y otros ajustes manuales con coherencia contable y fiscal.";
  }
  if (noInvoiceType) {
    populateNoInvoiceTypeSelect(noInvoiceType, allowedTypes);
    if (mode === "rent") {
      noInvoiceType.value = "alquiler_local";
    } else if (mode === "payroll") {
      noInvoiceType.value = "nomina";
    } else if (mode === "financing") {
      noInvoiceType.value = "prestamo";
    } else if (!["amortizacion", "kilometraje", "otro"].includes(noInvoiceType.value)) {
      noInvoiceType.value = "otro";
    }
    toggleLoanInterestField({
      typeValue: noInvoiceType.value,
      interestField: noInvoiceInterestField,
      interestInput: noInvoiceInterest,
      deductibleSelect: noInvoiceDeductible,
    });
    toggleNoInvoiceWithholdingField({
      typeValue: noInvoiceType.value,
    });
    toggleNoInvoiceVatFields({
      typeValue: noInvoiceType.value,
      vatDeductibleValue: noInvoiceVatDeductible?.value,
    });
  }
  if (loanPanel) {
    loanPanel.style.display = mode === "financing" ? "" : "none";
  }
  renderNoInvoiceExpenses(currentNoInvoiceExpenses || []);
}

function setExpenseSubview(subview) {
  const normalizedSubview = subview === "other" ? "adjustments" : subview;
  currentExpenseSubview = normalizedSubview;
  localStorage.setItem("expensesSubview", normalizedSubview);
  document.querySelectorAll('.section-tab[data-parent="expenses"]').forEach((tab) => {
    tab.classList.toggle("active", tab.dataset.subsection === normalizedSubview);
  });
  const receivedSection = document.querySelector(
    '.page-section[data-group="expenses"][data-subsection="received"]'
  );
  const operationsSection = document.querySelector(
    '.page-section[data-group="expenses"][data-subsection="operations"]'
  );
  if (receivedSection) {
    receivedSection.classList.remove("sub-hidden");
  }
  if (operationsSection) {
    operationsSection.classList.toggle("sub-hidden", normalizedSubview === "received");
  }
  if (expensesSavedInvoicesPanel) {
    expensesSavedInvoicesPanel.classList.toggle("is-hidden", normalizedSubview !== "received");
  }
  updateExpenseUploadPanel(normalizedSubview);
  if (normalizedSubview !== "received") {
    setExpenseMode(getExpenseModeForSubview(normalizedSubview));
  }
}

function setGroupedSubview(parent, subsection) {
  if (parent === "reports") {
    currentReportsSubview = subsection;
  } else if (parent === "statements") {
    currentStatementsSubview = subsection;
  }
  localStorage.setItem(`${parent}Subview`, subsection);
  document.querySelectorAll(`.section-tab[data-parent="${parent}"]`).forEach((tab) => {
    tab.classList.toggle("active", tab.dataset.subsection === subsection);
  });
  document.querySelectorAll(`.page-section[data-group="${parent}"]`).forEach((section) => {
    section.classList.toggle("sub-hidden", section.dataset.subsection !== subsection);
  });
}

function ensureCompositeSectionState(sectionId) {
  if (sectionId === "expenses") {
    setExpenseSubview(localStorage.getItem("expensesSubview") || currentExpenseSubview || "received");
    return;
  }
  if (sectionId === "reports") {
    const storedSubview = localStorage.getItem("reportsSubview");
    currentReportsSubview =
      storedSubview && ["summary", "taxes"].includes(storedSubview)
        ? storedSubview
        : currentReportsSubview === "calendar"
          ? "summary"
          : currentReportsSubview || "summary";
    setGroupedSubview("reports", currentReportsSubview);
    return;
  }
  if (sectionId === "statements") {
    currentStatementsSubview =
      localStorage.getItem("statementsSubview") || currentStatementsSubview || "pnl";
    setGroupedSubview("statements", currentStatementsSubview);
  }
}

function renderFiscalModelsSummary() {
  if (!fiscalModelsTableBody) {
    return;
  }
  const company = getSelectedCompany();
  if (!company) {
    fiscalModelsTableBody.innerHTML = "";
    return;
  }
  const { month, year } = getSelectedMonthYear();
  const period = getSelectedPeriod();
  const companyId = getSelectedCompanyId();
  const suffix = companyId ? `&company_id=${companyId}` : "";
  fiscalModelsTableBody.innerHTML = "";
  const loadingRow = document.createElement("tr");
  const loadingCell = document.createElement("td");
  loadingCell.colSpan = 4;
  loadingCell.textContent = "Calculando modelos…";
  loadingRow.appendChild(loadingCell);
  fiscalModelsTableBody.appendChild(loadingRow);

  fetch(`/api/fiscal-models/summary?month=${month}&year=${year}&period=${period}${suffix}`)
    .then((res) => res.json())
    .then((data) => {
      fiscalModelsTableBody.innerHTML = "";
      const rows = data.rows || [];
      if (!data.ok || !rows.length) {
        const tr = document.createElement("tr");
        const td = document.createElement("td");
        td.colSpan = 4;
        td.textContent =
          (data.errors && data.errors.join(" · ")) ||
          "Configura el perfil fiscal de la empresa para ver sus modelos.";
        tr.appendChild(td);
        fiscalModelsTableBody.appendChild(tr);
        return;
      }
      rows.forEach((row) => {
        const tr = document.createElement("tr");
        const values = [
          row.model,
          row.periodicity,
          row.status,
          typeof row.amount === "number" ? formatCurrency(row.amount) : row.amount || "0,00 €",
        ];
        values.forEach((value) => {
          const td = document.createElement("td");
          td.textContent = value;
          tr.appendChild(td);
        });
        fiscalModelsTableBody.appendChild(tr);
      });
    })
    .catch(() => {
      fiscalModelsTableBody.innerHTML = "";
      const tr = document.createElement("tr");
      const td = document.createElement("td");
      td.colSpan = 4;
      td.textContent = "No se pudo calcular el resumen de modelos.";
      tr.appendChild(td);
      fiscalModelsTableBody.appendChild(tr);
    });
}

function syncCompanyFiscalProfileInputs() {
  if (
    !companyType ||
    !companyFilesModel130 ||
    !companyFilesModel202 ||
    !companyVatRegime
  ) {
    return;
  }
  const isCompany = companyType.value === "company";
  companyFilesModel130.disabled = isCompany;
  companyFilesModel202.disabled = !isCompany;
  if (isCompany) {
    companyFilesModel130.value = "false";
  } else {
    companyFilesModel202.value = "false";
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
  const expenseUploadKind = getExpenseUploadKind();
  Array.from(fileList).forEach((file) => {
    if (!isAllowedFile(file.name)) {
      return;
    }
    const item = {
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      file,
      originalFilename: file.name,
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
      expenseUploadKind,
      payrollPeriod: "",
      payrollDeductionsAmount: "",
      payrollEmployerCostAmount: "",
      withholdingAmount: "",
      analysisText: "",
      analysisPending: true,
      analysisQueued: true,
      analysisError: false,
      analysisErrorMessage: "",
      analysisStatus: "ok",
      touched: {
        date: false,
        supplier: false,
        base: false,
        payrollDeductionsAmount: false,
        vat: false,
        vatAmount: false,
        total: false,
      },
    };
    pendingFiles.push(item);
    enqueueAnalysisTask(item, analyzeInvoiceForItem, renderTable);
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
      analysisQueued: true,
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
    enqueueAnalysisTask(item, analyzeIncomeForItem, renderIncomeTable);
  });
  renderIncomeTable();
}

function appendPendingPaymentDatesRow(item) {
  const paymentRow = document.createElement("tr");
  paymentRow.className = "payment-dates-row";
  const paymentCell = document.createElement("td");
  paymentCell.colSpan = 8;
  const paymentWrap = document.createElement("div");
  paymentWrap.className = "payment-dates-wrap";
  const paymentLabel = document.createElement("span");
  paymentLabel.className = "field-label";
  paymentLabel.textContent = "Fechas de pago";
  paymentWrap.appendChild(paymentLabel);

  const paymentInputs = document.createElement("div");
  paymentInputs.className = "payment-dates-inputs";
  const dates = getEffectivePaymentDates(item);
  item.paymentDates = dates.slice();
  item.paymentDate = dates[0] || "";

  const renderPaymentInputs = () => {
    paymentInputs.innerHTML = "";
    (item.paymentDates || []).forEach((dateValue, idx) => {
      const dateInput = document.createElement("input");
      dateInput.type = "date";
      dateInput.value = dateValue || "";
      dateInput.disabled = item.analysisPending;
      dateInput.addEventListener("change", () => {
        item.paymentDates[idx] = dateInput.value;
        item.paymentDate = item.paymentDates[0] || "";
      });
      paymentInputs.appendChild(dateInput);
      if ((item.paymentDates || []).length > 1) {
        const removeBtn = document.createElement("button");
        removeBtn.type = "button";
        removeBtn.className = "button ghost small";
        removeBtn.textContent = "Quitar";
        removeBtn.disabled = item.analysisPending;
        removeBtn.addEventListener("click", () => {
          item.paymentDates.splice(idx, 1);
          item.paymentDate = item.paymentDates[0] || "";
          renderPaymentInputs();
        });
        paymentInputs.appendChild(removeBtn);
      }
    });
  };
  renderPaymentInputs();

  const quickActions = document.createElement("div");
  quickActions.className = "payment-quick-actions";
  [15, 30, 60, 90].forEach((days) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "button ghost small";
    btn.textContent = `${days} días`;
    btn.disabled = item.analysisPending;
    btn.addEventListener("click", () => {
      if (!item.date) {
        alert("Selecciona primero la fecha del documento.");
        return;
      }
      const newDate = addDaysToISO(item.date, days);
      item.paymentDates = newDate ? [newDate] : [];
      item.paymentDate = newDate || "";
      renderPaymentInputs();
    });
    quickActions.appendChild(btn);
  });

  const addDateBtn = document.createElement("button");
  addDateBtn.type = "button";
  addDateBtn.className = "button ghost small";
  addDateBtn.textContent = "Añadir fecha";
  addDateBtn.disabled = item.analysisPending;
  addDateBtn.addEventListener("click", () => {
    item.paymentDates = item.paymentDates || [];
    item.paymentDates.push("");
    renderPaymentInputs();
  });

  paymentWrap.appendChild(paymentInputs);
  paymentWrap.appendChild(quickActions);
  paymentWrap.appendChild(addDateBtn);
  paymentCell.appendChild(paymentWrap);
  paymentRow.appendChild(paymentCell);
  uploadTableBody.appendChild(paymentRow);
}

function appendPendingStatusRow(item) {
  if (!item.analysisPending && !item.analysisError) {
    return;
  }
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
    : item.analysisQueued
      ? "Documento en cola… Se analiza de una en una para evitar bloqueos."
      : "Analizando documento… Puede tardar hasta 1 minuto.";
  statusWrapper.appendChild(message);
  statusTd.appendChild(statusWrapper);
  statusRow.appendChild(statusTd);
  uploadTableBody.appendChild(statusRow);
}

function appendPayrollPendingRow(item) {
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
    if (!item.paymentDate) {
      item.paymentDate = item.date;
    }
  });
  dateTd.appendChild(dateInput);

  const employeeTd = document.createElement("td");
  const employeeInput = document.createElement("input");
  employeeInput.type = "text";
  employeeInput.placeholder = "Empleado";
  employeeInput.value = item.supplier;
  employeeInput.disabled = item.analysisPending;
  employeeInput.addEventListener("input", () => {
    item.supplier = employeeInput.value;
    item.touched.supplier = true;
  });
  employeeTd.appendChild(employeeInput);

  const grossTd = document.createElement("td");
  const grossInput = document.createElement("input");
  grossInput.type = "text";
  grossInput.placeholder = "0,00";
  grossInput.value = item.base;
  grossInput.disabled = item.analysisPending;
  attachAmountInputBehavior(grossInput);
  grossInput.addEventListener("input", () => {
    item.base = grossInput.value;
    item.touched.base = true;
    const netAmount = getPayrollNetAmount(item);
    const grossAmount = getPayrollGrossAmount(item);
    if (grossAmount !== null && netAmount !== null) {
      item.payrollDeductionsAmount = formatAmountInput(
        Math.max(grossAmount - netAmount, 0)
      );
      deductionsInput.value = item.payrollDeductionsAmount;
    }
  });
  grossTd.appendChild(grossInput);

  const deductionsTd = document.createElement("td");
  const deductionsInput = document.createElement("input");
  deductionsInput.type = "text";
  deductionsInput.placeholder = "0,00";
  deductionsInput.value = item.payrollDeductionsAmount || "";
  deductionsInput.disabled = item.analysisPending;
  attachAmountInputBehavior(deductionsInput);
  deductionsInput.addEventListener("input", () => {
    item.payrollDeductionsAmount = deductionsInput.value;
    item.touched.payrollDeductionsAmount = true;
    const grossAmount = getPayrollGrossAmount(item);
    const deductionsAmount = getPayrollDeductionsAmount(item);
    if (grossAmount !== null && deductionsAmount !== null) {
      item.total = formatAmountInput(Math.max(grossAmount - deductionsAmount, 0));
      netInput.value = item.total;
    }
  });
  deductionsTd.appendChild(deductionsInput);

  const netTd = document.createElement("td");
  const netInput = document.createElement("input");
  netInput.type = "text";
  netInput.placeholder = "0,00";
  netInput.value = item.total;
  netInput.disabled = item.analysisPending;
  attachAmountInputBehavior(netInput);
  netInput.addEventListener("input", () => {
    item.total = netInput.value;
    item.touched.total = true;
    const grossAmount = getPayrollGrossAmount(item);
    const netAmount = getPayrollNetAmount(item);
    if (grossAmount !== null && netAmount !== null) {
      item.payrollDeductionsAmount = formatAmountInput(
        Math.max(grossAmount - netAmount, 0)
      );
      deductionsInput.value = item.payrollDeductionsAmount;
    }
  });
  netTd.appendChild(netInput);

  const periodTd = document.createElement("td");
  const periodInput = document.createElement("input");
  periodInput.type = "text";
  periodInput.placeholder = "YYYY-MM";
  periodInput.value = item.payrollPeriod || "";
  periodInput.disabled = item.analysisPending;
  periodInput.addEventListener("input", () => {
    item.payrollPeriod = periodInput.value;
  });
  periodTd.appendChild(periodInput);

  const actionsTd = document.createElement("td");
  actionsTd.classList.add("row-actions");
  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.textContent = item.analysisPending ? "Cancelar" : "Quitar";
  removeBtn.addEventListener("click", () => {
    if (item.analysisPending) {
      abortPendingAnalysis(item);
    }
    const index = pendingFiles.findIndex((entry) => entry.id === item.id);
    if (index !== -1) {
      pendingFiles.splice(index, 1);
      renderTable();
    }
  });
  actionsTd.appendChild(removeBtn);

  tr.appendChild(nameTd);
  tr.appendChild(dateTd);
  tr.appendChild(employeeTd);
  tr.appendChild(grossTd);
  tr.appendChild(deductionsTd);
  tr.appendChild(netTd);
  tr.appendChild(periodTd);
  tr.appendChild(actionsTd);
  uploadTableBody.appendChild(tr);

  appendPendingPaymentDatesRow(item);
  appendPendingStatusRow(item);
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
    if (isPayrollUploadItem(item)) {
      appendPayrollPendingRow(item);
      return;
    }
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
    baseInput.type = "text";
    baseInput.min = "0";
    baseInput.placeholder = "0,00";
    baseInput.value = item.base;
    const breakdownActive =
      Array.isArray(item.vatBreakdown) && item.vatBreakdown.length > 0;
    baseInput.disabled = item.analysisPending;
    baseInput.readOnly = breakdownActive;
    attachAmountInputBehavior(baseInput);
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
    vatAmountInput.type = "text";
    vatAmountInput.min = "0";
    vatAmountInput.placeholder = "0,00";
    vatAmountInput.value = item.vatAmount;
    vatAmountInput.disabled = item.analysisPending;
    vatAmountInput.readOnly = breakdownActive;
    attachAmountInputBehavior(vatAmountInput);
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

    const withholdingTd = document.createElement("td");
    const withholdingInput = document.createElement("input");
    withholdingInput.type = "text";
    withholdingInput.min = "0";
    withholdingInput.placeholder = "0,00";
    withholdingInput.value = item.withholdingAmount || "";
    withholdingInput.disabled = item.analysisPending;
    attachAmountInputBehavior(withholdingInput);
    withholdingInput.addEventListener("input", () => {
      item.withholdingAmount = withholdingInput.value;
      item.touched.withholdingAmount = true;
    });
    withholdingTd.appendChild(withholdingInput);

    const totalTd = document.createElement("td");
    const totalInput = document.createElement("input");
    totalInput.type = "text";
    totalInput.min = "0";
    totalInput.placeholder = "0,00";
    totalInput.value = item.total;
    totalInput.disabled = item.analysisPending;
    totalInput.readOnly = breakdownActive;
    attachAmountInputBehavior(totalInput);
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
    removeBtn.textContent = item.analysisPending ? "Cancelar" : "Quitar";
    removeBtn.addEventListener("click", () => {
      if (item.analysisPending) {
        abortPendingAnalysis(item);
      }
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
    tr.appendChild(withholdingTd);
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
        const withholdingCell = document.createElement("td");
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
        baseLineInput.type = "text";
        baseLineInput.min = "0";
        baseLineInput.placeholder = "0,00";
        baseLineInput.value = line.base || "";
        baseLineInput.disabled = item.analysisPending;
        baseCell.appendChild(baseLineInput);
        attachAmountInputBehavior(baseLineInput);

        const vatLineInput = document.createElement("input");
        vatLineInput.type = "text";
        vatLineInput.min = "0";
        vatLineInput.readOnly = false;
        vatLineInput.value = line.vat_amount || "";
        vatLineInput.disabled = item.analysisPending;
        vatCell.appendChild(vatLineInput);
        attachAmountInputBehavior(vatLineInput);

        const totalLineInput = document.createElement("input");
        totalLineInput.type = "text";
        totalLineInput.min = "0";
        totalLineInput.readOnly = false;
        totalLineInput.value = line.total || "";
        totalLineInput.disabled = item.analysisPending;
        totalCell.appendChild(totalLineInput);
        attachAmountInputBehavior(totalLineInput);

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

        const syncLine = (source = "auto") => {
          line.rate = rateSelect.value;
          line.base = baseLineInput.value;
          line.vat_amount = vatLineInput.value;
          line.total = totalLineInput.value;
          const normalized = normalizeBreakdownLine(line, source);
          if (normalized) {
            if (source !== "base") {
              line.base = formatAmountInput(normalized.base);
              baseLineInput.value = line.base;
            }
            if (source !== "vatAmount") {
              line.vat_amount = formatAmountInput(normalized.vat_amount);
              vatLineInput.value = line.vat_amount;
            }
            if (source !== "total") {
              line.total = formatAmountInput(normalized.total);
              totalLineInput.value = line.total;
            }
          } else {
            if (source !== "vatAmount") {
              line.vat_amount = "";
              vatLineInput.value = "";
            }
            if (source !== "total") {
              line.total = "";
              totalLineInput.value = "";
            }
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
        rateSelect.addEventListener("change", () => syncLine("rate"));
        baseLineInput.addEventListener("input", () => syncLine("base"));
        vatLineInput.addEventListener("input", () => syncLine("vatAmount"));
        totalLineInput.addEventListener("input", () => syncLine("total"));
        syncLine();

        row.appendChild(baseCell);
        row.appendChild(rateCell);
        row.appendChild(vatCell);
        row.appendChild(withholdingCell);
        row.appendChild(totalCell);
        row.appendChild(actionsCell);
        uploadTableBody.appendChild(row);
      });
    }

    const paymentRow = document.createElement("tr");
    paymentRow.className = "payment-dates-row";
    const paymentCell = document.createElement("td");
    paymentCell.colSpan = 9;
    const paymentWrap = document.createElement("div");
    paymentWrap.className = "payment-dates-wrap";
    const paymentLabel = document.createElement("span");
    paymentLabel.className = "field-label";
    paymentLabel.textContent = "Fechas de pago";
    paymentWrap.appendChild(paymentLabel);

    const paymentInputs = document.createElement("div");
    paymentInputs.className = "payment-dates-inputs";
    const dates = getEffectivePaymentDates(item);
    item.paymentDates = dates.slice();
    item.paymentDate = dates[0] || "";

    const renderPaymentInputs = () => {
      paymentInputs.innerHTML = "";
      (item.paymentDates || []).forEach((dateValue, idx) => {
        const dateInput = document.createElement("input");
        dateInput.type = "date";
        dateInput.value = dateValue || "";
        dateInput.disabled = item.analysisPending;
        dateInput.addEventListener("change", () => {
          item.paymentDates[idx] = dateInput.value;
          item.paymentDate = item.paymentDates[0] || "";
        });
        paymentInputs.appendChild(dateInput);
        if ((item.paymentDates || []).length > 1) {
          const removeBtn = document.createElement("button");
          removeBtn.type = "button";
          removeBtn.className = "button ghost small";
          removeBtn.textContent = "Quitar";
          removeBtn.disabled = item.analysisPending;
          removeBtn.addEventListener("click", () => {
            item.paymentDates.splice(idx, 1);
            item.paymentDate = item.paymentDates[0] || "";
            renderPaymentInputs();
          });
          paymentInputs.appendChild(removeBtn);
        }
      });
    };
    renderPaymentInputs();

    const quickActions = document.createElement("div");
    quickActions.className = "payment-quick-actions";
    [15, 30, 60, 90].forEach((days) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "button ghost small";
      btn.textContent = `${days} días`;
      btn.disabled = item.analysisPending;
      btn.addEventListener("click", () => {
        if (!item.date) {
          alert("Selecciona primero la fecha de factura.");
          return;
        }
        const newDate = addDaysToISO(item.date, days);
        item.paymentDates = newDate ? [newDate] : [];
        item.paymentDate = newDate || "";
        renderPaymentInputs();
      });
      quickActions.appendChild(btn);
    });

    const addDateBtn = document.createElement("button");
    addDateBtn.type = "button";
    addDateBtn.className = "button ghost small";
    addDateBtn.textContent = "Añadir fecha";
    addDateBtn.disabled = item.analysisPending;
    addDateBtn.addEventListener("click", () => {
      item.paymentDates = item.paymentDates || [];
      item.paymentDates.push("");
      renderPaymentInputs();
    });

    paymentWrap.appendChild(paymentInputs);
    paymentWrap.appendChild(quickActions);
    paymentWrap.appendChild(addDateBtn);
    paymentCell.appendChild(paymentWrap);
    paymentRow.appendChild(paymentCell);
    uploadTableBody.appendChild(paymentRow);

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
      statusTd.colSpan = 9;
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
        : item.analysisQueued
          ? "Factura en cola… Se analiza de una en una para evitar bloqueos con escaneadas."
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
    const clientWarning = document.createElement("div");
    clientWarning.className = "field-warning";
    const updateClientWarning = () => {
      const value = clientInput.value.trim();
      if (!value) {
        clientWarning.textContent = "Cliente pendiente de completar.";
        clientWarning.style.display = "block";
        clientInput.classList.add("input-warning");
        return;
      }
      if (isSupplierSameAsCompany(value)) {
        clientWarning.textContent =
          "El cliente no puede ser la empresa activa.";
        clientWarning.style.display = "block";
        clientInput.classList.add("input-warning");
        return;
      }
      clientWarning.textContent = "";
      clientWarning.style.display = "none";
      clientInput.classList.remove("input-warning");
    };
    clientInput.addEventListener("input", () => {
      item.client = clientInput.value;
      item.touched.client = true;
      updateClientWarning();
    });
    clientTd.appendChild(clientInput);
    clientTd.appendChild(clientWarning);

    const baseTd = document.createElement("td");
    const baseInput = document.createElement("input");
    baseInput.type = "text";
    baseInput.min = "0";
    baseInput.placeholder = "0,00";
    baseInput.value = item.base;
    const breakdownActive =
      Array.isArray(item.vatBreakdown) && item.vatBreakdown.length > 0;
    baseInput.disabled = item.analysisPending;
    baseInput.readOnly = breakdownActive;
    attachAmountInputBehavior(baseInput);
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
    vatAmountInput.type = "text";
    vatAmountInput.min = "0";
    vatAmountInput.placeholder = "0,00";
    vatAmountInput.value = item.vatAmount;
    vatAmountInput.disabled = item.analysisPending;
    vatAmountInput.readOnly = breakdownActive;
    attachAmountInputBehavior(vatAmountInput);
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
    totalInput.type = "text";
    totalInput.min = "0";
    totalInput.placeholder = "0,00";
    totalInput.value = item.total;
    totalInput.disabled = item.analysisPending;
    totalInput.readOnly = breakdownActive;
    attachAmountInputBehavior(totalInput);
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
    removeBtn.textContent = item.analysisPending ? "Cancelar" : "Quitar";
    removeBtn.addEventListener("click", () => {
      if (item.analysisPending) {
        abortPendingAnalysis(item);
      }
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
        baseLineInput.type = "text";
        baseLineInput.min = "0";
        baseLineInput.placeholder = "0,00";
        baseLineInput.value = line.base || "";
        baseLineInput.disabled = item.analysisPending;
        baseCell.appendChild(baseLineInput);
        attachAmountInputBehavior(baseLineInput);

        const vatLineInput = document.createElement("input");
        vatLineInput.type = "text";
        vatLineInput.min = "0";
        vatLineInput.readOnly = false;
        vatLineInput.value = line.vat_amount || "";
        vatLineInput.disabled = item.analysisPending;
        vatCell.appendChild(vatLineInput);
        attachAmountInputBehavior(vatLineInput);

        const totalLineInput = document.createElement("input");
        totalLineInput.type = "text";
        totalLineInput.min = "0";
        totalLineInput.readOnly = false;
        totalLineInput.value = line.total || "";
        totalLineInput.disabled = item.analysisPending;
        totalCell.appendChild(totalLineInput);
        attachAmountInputBehavior(totalLineInput);

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

        const syncLine = (source = "auto") => {
          line.rate = rateSelect.value;
          line.base = baseLineInput.value;
          line.vat_amount = vatLineInput.value;
          line.total = totalLineInput.value;
          const normalized = normalizeBreakdownLine(line, source);
          if (normalized) {
            if (source !== "base") {
              line.base = formatAmountInput(normalized.base);
              baseLineInput.value = line.base;
            }
            if (source !== "vatAmount") {
              line.vat_amount = formatAmountInput(normalized.vat_amount);
              vatLineInput.value = line.vat_amount;
            }
            if (source !== "total") {
              line.total = formatAmountInput(normalized.total);
              totalLineInput.value = line.total;
            }
          } else {
            if (source !== "vatAmount") {
              line.vat_amount = "";
              vatLineInput.value = "";
            }
            if (source !== "total") {
              line.total = "";
              totalLineInput.value = "";
            }
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
        rateSelect.addEventListener("change", () => syncLine("rate"));
        baseLineInput.addEventListener("input", () => syncLine("base"));
        vatLineInput.addEventListener("input", () => syncLine("vatAmount"));
        totalLineInput.addEventListener("input", () => syncLine("total"));
        syncLine();

        row.appendChild(baseCell);
        row.appendChild(rateCell);
        row.appendChild(vatCell);
        row.appendChild(totalCell);
        row.appendChild(actionsCell);
        incomeUploadTableBody.appendChild(row);
      });
    }

    const paymentRow = document.createElement("tr");
    paymentRow.className = "payment-dates-row";
    const paymentCell = document.createElement("td");
    paymentCell.colSpan = 8;
    const paymentWrap = document.createElement("div");
    paymentWrap.className = "payment-dates-wrap";
    const paymentLabel = document.createElement("span");
    paymentLabel.className = "field-label";
    paymentLabel.textContent = "Fechas de cobro";
    paymentWrap.appendChild(paymentLabel);

    const paymentInputs = document.createElement("div");
    paymentInputs.className = "payment-dates-inputs";
    const dates = getEffectivePaymentDates(item);
    item.paymentDates = dates.slice();
    item.paymentDate = dates[0] || "";

    const renderPaymentInputs = () => {
      paymentInputs.innerHTML = "";
      (item.paymentDates || []).forEach((dateValue, idx) => {
        const dateInput = document.createElement("input");
        dateInput.type = "date";
        dateInput.value = dateValue || "";
        dateInput.disabled = item.analysisPending;
        dateInput.addEventListener("change", () => {
          item.paymentDates[idx] = dateInput.value;
          item.paymentDate = item.paymentDates[0] || "";
        });
        paymentInputs.appendChild(dateInput);
        if ((item.paymentDates || []).length > 1) {
          const removeBtn = document.createElement("button");
          removeBtn.type = "button";
          removeBtn.className = "button ghost small";
          removeBtn.textContent = "Quitar";
          removeBtn.disabled = item.analysisPending;
          removeBtn.addEventListener("click", () => {
            item.paymentDates.splice(idx, 1);
            item.paymentDate = item.paymentDates[0] || "";
            renderPaymentInputs();
          });
          paymentInputs.appendChild(removeBtn);
        }
      });
    };
    renderPaymentInputs();

    const quickActions = document.createElement("div");
    quickActions.className = "payment-quick-actions";
    [15, 30, 60, 90].forEach((days) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "button ghost small";
      btn.textContent = `${days} días`;
      btn.disabled = item.analysisPending;
      btn.addEventListener("click", () => {
        if (!item.date) {
          alert("Selecciona primero la fecha de factura.");
          return;
        }
        const newDate = addDaysToISO(item.date, days);
        item.paymentDates = newDate ? [newDate] : [];
        item.paymentDate = newDate || "";
        renderPaymentInputs();
      });
      quickActions.appendChild(btn);
    });

    const addDateBtn = document.createElement("button");
    addDateBtn.type = "button";
    addDateBtn.className = "button ghost small";
    addDateBtn.textContent = "Añadir fecha";
    addDateBtn.disabled = item.analysisPending;
    addDateBtn.addEventListener("click", () => {
      item.paymentDates = item.paymentDates || [];
      item.paymentDates.push("");
      renderPaymentInputs();
    });

    paymentWrap.appendChild(paymentInputs);
    paymentWrap.appendChild(quickActions);
    paymentWrap.appendChild(addDateBtn);
    paymentCell.appendChild(paymentWrap);
    paymentRow.appendChild(paymentCell);
    incomeUploadTableBody.appendChild(paymentRow);

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
        : item.analysisQueued
          ? "Factura en cola… Se analiza de una en una para evitar bloqueos con escaneadas."
          : "Analizando factura… Las facturas escaneadas pueden tardar hasta 1 minuto.";
      statusWrapper.appendChild(message);
      statusTd.appendChild(statusWrapper);
      statusRow.appendChild(statusTd);
      incomeUploadTableBody.appendChild(statusRow);
    }
  });
}

function analyzeIncomeForItem(item) {
  if (item._analysisCancelled || !isPendingUploadItemPresent(item)) {
    return Promise.resolve();
  }
  const formData = new FormData();
  formData.append("file", item.file);
  formData.append("document_type", "income");
  const companyId = getSelectedCompanyId();
  if (companyId) {
    formData.append("company_id", companyId);
  }
  const controller = new AbortController();
  item._analysisController = controller;

  return fetch("/api/analyze-invoice", {
    method: "POST",
    body: formData,
    signal: controller.signal,
  })
    .then((res) => res.json())
    .then((data) => {
      if (item._analysisCancelled || !isPendingUploadItemPresent(item)) {
        return;
      }
      if (!data.ok) {
        item.analysisPending = false;
        item.analysisQueued = false;
        item.analysisError = true;
        item.analysisErrorMessage = ANALYSIS_ERROR_MESSAGE;
        renderIncomeTable();
        return;
      }
      const extracted = data.extracted || {};
      item.analysisText = extracted.analysis_text || "";
      item.analysisStatus = extracted.analysis_status || "ok";
      const extractedBreakdown = parseVatBreakdown(
        extracted.vat_breakdown || extracted.vatBreakdown
      );
      if (extractedBreakdown.length) {
        item.vatBreakdown = extractedBreakdown;
        item.vatBreakdownOpen = extractedBreakdown.length > 1;
      }
      if (extracted.breakdown_warning) {
        item.analysisWarning = VAT_WARNING_MESSAGE;
        if (!vatWarningDismissedIds.has(item.id)) {
          showVatWarningModal();
          vatWarningDismissedIds.add(item.id);
        }
      }
      if (extracted.analysis_status === "low_quality_scan") {
        item.analysisPending = false;
        item.analysisQueued = false;
        item.analysisError = true;
        item.analysisErrorMessage = LOW_QUALITY_SCAN_MESSAGE;
        if (!lowQualityDismissedIds.has(item.id)) {
          showLowQualityModal();
          lowQualityDismissedIds.add(item.id);
        }
        renderIncomeTable();
        return;
      }
      if (extracted.analysis_status === "timeout") {
        item.analysisPending = false;
        item.analysisQueued = false;
        item.analysisError = true;
        item.analysisErrorMessage = TIMEOUT_MESSAGE;
        renderIncomeTable();
        return;
      }
      if (extracted.is_rectificativa) {
        item.isRectificativa = true;
      }
      const detectedClient = extracted.client_name || extracted.client;

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
      item.analysisQueued = false;
      const hasExtractedValue = [
        detectedClient,
        extracted.invoice_date,
        extracted.base_amount,
        extracted.vat_rate,
        extracted.vat_amount,
        extracted.total_amount,
        extracted.vat_breakdown,
        extracted.totals,
      ].some((value) => value !== null && value !== undefined && value !== "");
      item.analysisError = !hasExtractedValue && !item.analysisText;
      item.analysisErrorMessage = item.analysisError ? ANALYSIS_ERROR_MESSAGE : "";
      renderIncomeTable();
    })
    .catch((error) => {
      if (error && error.name === "AbortError") {
        return;
      }
      if (item._analysisCancelled || !isPendingUploadItemPresent(item)) {
        return;
      }
      item.analysisPending = false;
      item.analysisQueued = false;
      item.analysisError = true;
      item.analysisErrorMessage = ANALYSIS_ERROR_MESSAGE;
      renderIncomeTable();
    })
    .finally(() => {
      if (item._analysisTimeoutId) {
        clearTimeout(item._analysisTimeoutId);
        item._analysisTimeoutId = null;
      }
      item._analysisController = null;
    });
}

function validateIncomePending() {
  const errors = [];
  pendingIncomeFiles.forEach((item) => {
    if (item.analysisPending) {
      errors.push(`Análisis en proceso: ${item.file.name}`);
    }
    if (!item.client.trim()) {
      errors.push(`Cliente obligatorio: ${item.file.name}`);
    }
    if (item.client.trim() && isSupplierSameAsCompany(item.client)) {
      errors.push(`El cliente no puede ser la empresa activa: ${item.file.name}`);
    }
    const baseValue = parseNumberInput(item.base);
    const totalValue = parseNumberInput(item.total);
    if (baseValue === null && totalValue === null) {
      errors.push(`Base imponible o total obligatorio: ${item.file.name}`);
    }
    const isRectificativa =
      item.isRectificativa ||
      (baseValue !== null && baseValue < 0) ||
      (totalValue !== null && totalValue < 0);
    if (baseValue !== null && baseValue < 0 && !isRectificativa) {
      errors.push(`Base imponible inválida: ${item.file.name}`);
    }
    if (totalValue !== null && totalValue < 0 && !isRectificativa) {
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
        originalFilename: item.originalFilename,
        date: item.date,
        paymentDate: computePaymentDate(item.date, item.paymentDate),
        paymentDates: item.paymentDates || [],
        analysisStatus: item.analysisStatus || "ok",
        isRectificativa: item.isRectificativa || false,
        client: item.client.trim(),
        base: breakdownTotals ? breakdownTotals.base : normalized.base || item.base,
        vat: breakdownPayload.length
          ? getPrimaryVatRateFromBreakdown(breakdownPayload)
          : resolveVatRateValue(item.vat),
        vatAmount: breakdownTotals ? breakdownTotals.vatAmount : normalized.vatAmount || item.vatAmount,
        total: breakdownTotals ? breakdownTotals.total : normalized.total || item.total,
        vatBreakdown: breakdownPayload,
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
    if (!item.supplier.trim()) {
      errors.push(`Proveedor obligatorio: ${item.file.name}`);
    }
    if (isSupplierSameAsCompany(item.supplier)) {
      errors.push(`El proveedor no puede ser la empresa activa: ${item.file.name}`);
    }
    const baseValue = parseNumberInput(item.base);
    const totalValue = parseNumberInput(item.total);
    const withholdingValue = parseNumberInput(item.withholdingAmount);
    if (baseValue === null && totalValue === null) {
      errors.push(`Base imponible o total obligatorio: ${item.file.name}`);
    }
    if (withholdingValue !== null && withholdingValue < 0) {
      errors.push(`Retención inválida: ${item.file.name}`);
    }
    if (withholdingValue !== null && totalValue !== null && withholdingValue > totalValue) {
      errors.push(`La retención no puede superar el total: ${item.file.name}`);
    }
    const isRectificativa =
      item.isRectificativa ||
      (baseValue !== null && baseValue < 0) ||
      (totalValue !== null && totalValue < 0);
    if (baseValue !== null && baseValue < 0 && !isRectificativa) {
      errors.push(`Base imponible inválida: ${item.file.name}`);
    }
    if (totalValue !== null && totalValue < 0 && !isRectificativa) {
      errors.push(`Total inválido: ${item.file.name}`);
    }
    if (!item.date) {
      errors.push(`Fecha obligatoria: ${item.file.name}`);
    }
  });
  return errors;
}

function validatePendingOperationItems(items) {
  const errors = [];
  items.forEach((item) => {
    if (item.analysisPending) {
      errors.push(`Análisis en proceso: ${item.file.name}`);
    }
    if (!item.date) {
      errors.push(`Fecha obligatoria: ${item.file.name}`);
    }
    if (isPayrollUploadItem(item)) {
      if (getPayrollGrossAmount(item) === null && getPayrollNetAmount(item) === null) {
        errors.push(`Bruto o líquido obligatorio: ${item.file.name}`);
      }
      return;
    }
    const normalized = normalizeInvoiceAmounts(item);
    const breakdownPayload = buildVatBreakdownPayload(item.vatBreakdown || []);
    const breakdownTotals = breakdownPayload.length
      ? summarizeVatBreakdown(item.vatBreakdown || [])
      : null;
    const amountValue = parseNumberInput(
      breakdownTotals ? breakdownTotals.total : normalized.total || item.total
    );
    const baseValue = parseNumberInput(
      breakdownTotals ? breakdownTotals.base : normalized.base || item.base
    );
    if (amountValue === null && baseValue === null) {
      errors.push(`Importe inválido: ${item.file.name}`);
    }
  });
  return errors;
}

function isPayrollUploadItem(item) {
  return (item?.expenseUploadKind || getExpenseUploadKind()) === "payroll";
}

function getPayrollGrossAmount(item) {
  const explicitGross = parseNumberInput(item.base);
  if (explicitGross !== null) {
    return explicitGross;
  }
  const netAmount = parseNumberInput(item.total);
  const deductionsAmount = parseNumberInput(item.payrollDeductionsAmount);
  if (netAmount !== null && deductionsAmount !== null) {
    return roundAmount(netAmount + deductionsAmount);
  }
  return null;
}

function getPayrollNetAmount(item) {
  const netAmount = parseNumberInput(item.total);
  if (netAmount !== null) {
    return netAmount;
  }
  const grossAmount = parseNumberInput(item.base);
  const deductionsAmount = parseNumberInput(item.payrollDeductionsAmount);
  if (grossAmount !== null && deductionsAmount !== null) {
    return roundAmount(Math.max(grossAmount - deductionsAmount, 0));
  }
  return null;
}

function getPayrollDeductionsAmount(item) {
  const deductionsAmount = parseNumberInput(item.payrollDeductionsAmount);
  if (deductionsAmount !== null) {
    return deductionsAmount;
  }
  const grossAmount = parseNumberInput(item.base);
  const netAmount = parseNumberInput(item.total);
  if (grossAmount !== null && netAmount !== null) {
    return roundAmount(Math.max(grossAmount - netAmount, 0));
  }
  return null;
}

function getPayrollConcept(item) {
  const employeeName = (item.supplier || "").trim();
  const periodLabel = (item.payrollPeriod || "").trim();
  const parts = ["Nómina"];
  if (employeeName) {
    parts.push(employeeName);
  }
  if (periodLabel) {
    parts.push(periodLabel);
  }
  return parts.join(" - ");
}

function getExpenseOperationPayload(item) {
  if (isPayrollUploadItem(item)) {
    const grossAmount = getPayrollGrossAmount(item);
    const netAmount = getPayrollNetAmount(item);
    const deductionsAmount = getPayrollDeductionsAmount(item);
    return {
      expense_date: item.date,
      payment_date: item.paymentDate || item.date || "",
      payment_dates: item.paymentDates && item.paymentDates.length ? item.paymentDates : [item.date],
      concept: getPayrollConcept(item),
      amount: grossAmount,
      expense_type: "nomina",
      interest_amount: null,
      vat_deductible: false,
      vat_rate: null,
      vat_amount: null,
      base_amount: grossAmount,
      withholding_amount: parseNumberInput(item.withholdingAmount) || 0,
      payroll_employee_name: (item.supplier || "").trim(),
      payroll_period: (item.payrollPeriod || "").trim(),
      payroll_net_amount: netAmount,
      payroll_total_deductions_amount: deductionsAmount,
      payroll_employer_cost_amount:
        parseNumberInput(item.payrollEmployerCostAmount) || null,
      deductible: true,
    };
  }
  const normalized = normalizeInvoiceAmounts(item);
  const breakdownPayload = buildVatBreakdownPayload(item.vatBreakdown || []);
  const breakdownTotals = breakdownPayload.length
    ? summarizeVatBreakdown(item.vatBreakdown || [])
    : null;
  const baseValue = parseNumberInput(
    breakdownTotals ? breakdownTotals.base : normalized.base || item.base
  );
  const vatAmountValue = parseNumberInput(
    breakdownTotals ? breakdownTotals.vatAmount : normalized.vatAmount || item.vatAmount
  );
  const totalValue = parseNumberInput(
    breakdownTotals ? breakdownTotals.total : normalized.total || item.total
  );
  const vatRateValue = breakdownPayload.length
    ? getPrimaryVatRateFromBreakdown(breakdownPayload)
    : resolveVatRateValue(item.vat);
  const kind = item.expenseUploadKind || getExpenseUploadKind();
  const expenseType = getNoInvoiceTypeForUploadKind(kind);
  const vatDeductible =
    expenseType !== "nomina" &&
    expenseType !== "seguridad_social" &&
    expenseType !== "prestamo" &&
    vatRateValue !== null &&
    vatAmountValue !== null &&
    vatAmountValue > 0;
  const amount =
    totalValue !== null
      ? totalValue
      : baseValue !== null && vatAmountValue !== null
        ? roundAmount(baseValue + vatAmountValue)
        : baseValue;
  const concept =
    (item.supplier || "").trim() || item.file.name.replace(/\.[^.]+$/, "");
  return {
    expense_date: item.date,
    payment_date: item.paymentDate || "",
    payment_dates: item.paymentDates || [],
    concept,
    amount,
    expense_type: expenseType,
    interest_amount: expenseType === "prestamo" ? amount : null,
    vat_deductible: vatDeductible,
    vat_rate: vatDeductible ? vatRateValue : null,
    vat_amount: vatDeductible ? vatAmountValue : null,
    base_amount: baseValue,
    withholding_amount: parseNumberInput(item.withholdingAmount) || 0,
    deductible: expenseType !== "prestamo",
  };
}

function analyzeInvoiceForItem(item) {
  if (item._analysisCancelled || !isPendingUploadItemPresent(item)) {
    return Promise.resolve();
  }
  const formData = new FormData();
  formData.append("file", item.file);
  const expenseUploadKind = item.expenseUploadKind || getExpenseUploadKind();
  const expenseDocumentType =
    expenseUploadKind === "received"
      ? "expense"
      : expenseUploadKind === "rent"
        ? "expense_rent"
        : expenseUploadKind === "payroll"
          ? "expense_payroll"
          : "expense_other";
  formData.append("document_type", expenseDocumentType);
  const companyId = getSelectedCompanyId();
  if (companyId) {
    formData.append("company_id", companyId);
  }
  const controller = new AbortController();
  item._analysisController = controller;

  return fetch("/api/analyze-invoice", {
    method: "POST",
    body: formData,
    signal: controller.signal,
  })
    .then((res) => res.json())
    .then((data) => {
      if (item._analysisCancelled || !isPendingUploadItemPresent(item)) {
        return;
      }
      if (!data.ok) {
        item.analysisPending = false;
        item.analysisQueued = false;
        item.analysisError = true;
        item.analysisErrorMessage = ANALYSIS_ERROR_MESSAGE;
        renderTable();
        return;
      }
      const extracted = data.extracted || {};
      item.analysisText = extracted.analysis_text || "";
      item.analysisStatus = extracted.analysis_status || "ok";
      const extractedBreakdown = parseVatBreakdown(
        extracted.vat_breakdown || extracted.vatBreakdown
      );
      if (extractedBreakdown.length) {
        item.vatBreakdown = extractedBreakdown;
        item.vatBreakdownOpen = extractedBreakdown.length > 1;
      }
      if (extracted.breakdown_warning) {
        item.analysisWarning = VAT_WARNING_MESSAGE;
        if (!vatWarningDismissedIds.has(item.id)) {
          showVatWarningModal();
          vatWarningDismissedIds.add(item.id);
        }
      }
      if (extracted.analysis_status === "low_quality_scan") {
        item.analysisPending = false;
        item.analysisQueued = false;
        item.analysisError = true;
        item.analysisErrorMessage = LOW_QUALITY_SCAN_MESSAGE;
        if (!lowQualityDismissedIds.has(item.id)) {
          showLowQualityModal();
          lowQualityDismissedIds.add(item.id);
        }
        renderTable();
        return;
      }
      if (extracted.analysis_status === "timeout") {
        item.analysisPending = false;
        item.analysisQueued = false;
        item.analysisError = true;
        item.analysisErrorMessage = TIMEOUT_MESSAGE;
        renderTable();
        return;
      }
      if (extracted.is_rectificativa) {
        item.isRectificativa = true;
      }
      if (extracted.withholding_amount !== null && extracted.withholding_amount !== undefined) {
        item.withholdingAmount = String(extracted.withholding_amount);
      }

      if (isPayrollUploadItem(item)) {
        if (!item.touched.supplier && extracted.employee_name) {
          item.supplier = extracted.employee_name;
        }
        if (!item.touched.base && extracted.base_amount !== null && extracted.base_amount !== undefined) {
          item.base = String(extracted.base_amount);
        }
        if (
          !item.touched.payrollDeductionsAmount &&
          extracted.payroll_total_deductions_amount !== null &&
          extracted.payroll_total_deductions_amount !== undefined
        ) {
          item.payrollDeductionsAmount = String(extracted.payroll_total_deductions_amount);
        }
        if (
          !item.touched.total &&
          extracted.payroll_net_amount !== null &&
          extracted.payroll_net_amount !== undefined
        ) {
          item.total = String(extracted.payroll_net_amount);
        }
        if (!item.payrollPeriod && extracted.payroll_period) {
          item.payrollPeriod = extracted.payroll_period;
        }
        if (
          extracted.payroll_employer_cost_amount !== null &&
          extracted.payroll_employer_cost_amount !== undefined
        ) {
          item.payrollEmployerCostAmount = String(extracted.payroll_employer_cost_amount);
        }
      } else if (!item.touched.supplier && extracted.provider_name) {
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
        item.paymentDate = isPayrollUploadItem(item)
          ? primaryDate || item.date
          : computePaymentDate(item.date, primaryDate);
      }
      if (
        !isPayrollUploadItem(item) &&
        !item.touched.base &&
        extracted.base_amount !== null &&
        extracted.base_amount !== undefined
      ) {
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
      if (
        !isPayrollUploadItem(item) &&
        !item.touched.total &&
        extracted.total_amount !== null &&
        extracted.total_amount !== undefined
      ) {
        item.total = String(extracted.total_amount);
      }

      if (!isPayrollUploadItem(item)) {
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
      } else if (!item.payrollDeductionsAmount) {
        const grossAmount = getPayrollGrossAmount(item);
        const netAmount = getPayrollNetAmount(item);
        if (grossAmount !== null && netAmount !== null) {
          item.payrollDeductionsAmount = String(
            roundAmount(Math.max(grossAmount - netAmount, 0))
          );
        }
      }

      item.analysisPending = false;
      item.analysisQueued = false;
      const hasExtractedValue = [
        extracted.employee_name,
        extracted.provider_name,
        extracted.invoice_date,
        extracted.base_amount,
        extracted.payroll_total_deductions_amount,
        extracted.payroll_net_amount,
        extracted.vat_rate,
        extracted.vat_amount,
        extracted.total_amount,
        extracted.vat_breakdown,
      ].some((value) => value !== null && value !== undefined && value !== "");
      item.analysisError = !hasExtractedValue && !item.analysisText;
      item.analysisErrorMessage = item.analysisError ? ANALYSIS_ERROR_MESSAGE : "";
      renderTable();
    })
    .catch((error) => {
      if (error && error.name === "AbortError") {
        return;
      }
      if (item._analysisCancelled || !isPendingUploadItemPresent(item)) {
        return;
      }
      item.analysisPending = false;
      item.analysisQueued = false;
      item.analysisError = true;
      item.analysisErrorMessage = ANALYSIS_ERROR_MESSAGE;
      renderTable();
    })
    .finally(() => {
      if (item._analysisTimeoutId) {
        clearTimeout(item._analysisTimeoutId);
        item._analysisTimeoutId = null;
      }
      item._analysisController = null;
    });
}

function uploadPending() {
  const invoiceItems = pendingFiles.filter(
    (item) => (item.expenseUploadKind || "received") === "received"
  );
  const operationItems = pendingFiles.filter(
    (item) => (item.expenseUploadKind || "received") !== "received"
  );

  if (!pendingFiles.length) {
    alert("No hay gastos para subir.");
    return;
  }
  if (!getSelectedCompanyId()) {
    alert("Selecciona una empresa antes de subir gastos.");
    return;
  }

  const errors = [
    ...validatePending(invoiceItems),
    ...validatePendingOperationItems(operationItems),
  ];
  if (errors.length) {
    alert(errors.slice(0, 3).join("\n"));
    return;
  }

  uploadBtn.disabled = true;
  const uploadedIds = [];
  const groupedErrors = [];

  const payload = {
    companyId: getSelectedCompanyId(),
    entries: invoiceItems.map((item) => {
      const normalized = normalizeInvoiceAmounts(item);
      const breakdownPayload = buildVatBreakdownPayload(item.vatBreakdown || []);
      const breakdownTotals = breakdownPayload.length
        ? summarizeVatBreakdown(item.vatBreakdown || [])
        : null;
      return {
        originalFilename: item.originalFilename,
        date: item.date,
        paymentDate: computePaymentDate(item.date, item.paymentDate),
        paymentDates: item.paymentDates || [],
        analysisStatus: item.analysisStatus || "ok",
        companyId: getSelectedCompanyId(),
        supplier: item.supplier.trim(),
        isRectificativa: item.isRectificativa || false,
        base: breakdownTotals ? breakdownTotals.base : normalized.base || item.base,
        vat: breakdownPayload.length
          ? getPrimaryVatRateFromBreakdown(breakdownPayload)
          : resolveVatRateValue(item.vat),
        vatAmount: breakdownTotals ? breakdownTotals.vatAmount : normalized.vatAmount || item.vatAmount,
        total: breakdownTotals ? breakdownTotals.total : normalized.total || item.total,
        withholding_amount: parseNumberInput(item.withholdingAmount) || 0,
        vatBreakdown: breakdownPayload,
      };
    }),
  };
  const invoicePromise = invoiceItems.length
    ? fetch("/api/upload", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(payload),
      })
        .then((res) => res.json())
        .then((data) => {
          if (!data.ok) {
            groupedErrors.push(...(data.errors || ["Error al subir facturas."]));
            return;
          }
          uploadedIds.push(...invoiceItems.map((item) => item.id));
          if (data.errors && data.errors.length) {
            groupedErrors.push(...data.errors);
          }
        })
        .catch(() => {
          groupedErrors.push("No se pudieron subir las facturas.");
        })
    : Promise.resolve();

  const operationsPromise = operationItems.length
    ? Promise.all(
        operationItems.map((item) =>
          fetch(withCompanyParam("/api/expenses/no-invoice"), {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              ...getExpenseOperationPayload(item),
              company_id: getSelectedCompanyId(),
            }),
          })
            .then((res) => res.json())
            .then((data) => {
              if (!data.ok) {
                groupedErrors.push(...(data.errors || [`Error al guardar ${item.file.name}.`]));
                return;
              }
              uploadedIds.push(item.id);
            })
            .catch(() => {
              groupedErrors.push(`No se pudo guardar ${item.file.name}.`);
            })
        )
      )
    : Promise.resolve();

  Promise.all([invoicePromise, operationsPromise])
    .then(() => {
      if (uploadedIds.length) {
        const uploadedSet = new Set(uploadedIds);
        for (let index = pendingFiles.length - 1; index >= 0; index -= 1) {
          if (uploadedSet.has(pendingFiles[index].id)) {
            pendingFiles.splice(index, 1);
          }
        }
      }
      renderTable();
      if (groupedErrors.length) {
        alert(groupedErrors.slice(0, 5).join("\n"));
      }
      return loadYears().then(() => {
        refreshAllData();
      });
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
  updateDashboardTotals();
}

function updateDashboardTotals() {
  const selectedMetrics = currentFinancialMetrics?.selected || null;
  const expenseInvoiceTotal = (currentInvoices || []).reduce(
    (sum, invoice) => sum + (Number(invoice.total_amount) || 0),
    0
  );
  const summaryInvoiceTotal = Number(currentSummary?.totalSpent || 0);
  const noInvoiceTotal = (currentNoInvoiceExpenses || []).reduce(
    (sum, expense) => sum + (Number(expense.amount) || 0),
    0
  );
  const invoiceGrossTotal =
    expenseInvoiceTotal > 0 || !summaryInvoiceTotal
      ? expenseInvoiceTotal
      : summaryInvoiceTotal;
  const expenseTotalGross = invoiceGrossTotal + noInvoiceTotal;

  const invoiceVatSupported = (currentInvoices || []).reduce(
    (sum, invoice) => sum + (Number(invoice.vat_amount) || 0),
    0
  );
  const noInvoiceVatSupported = (currentNoInvoiceExpenses || []).reduce(
    (sum, expense) => sum + getNoInvoiceVatDeductibleAmount(expense),
    0
  );
  const summaryVatTotals = currentSummary?.vatTotals || null;
  const summaryVatTotal =
    summaryVatTotals
      ? (Number(summaryVatTotals["0"]) || 0) +
        (Number(summaryVatTotals["4"]) || 0) +
        (Number(summaryVatTotals["10"]) || 0) +
        (Number(summaryVatTotals["21"]) || 0)
      : 0;
  const invoiceVatSupportedTotal =
    invoiceVatSupported > 0 || !summaryVatTotal
      ? invoiceVatSupported
      : Math.max(summaryVatTotal - noInvoiceVatSupported, 0);
  const vatSupported = invoiceVatSupportedTotal + noInvoiceVatSupported;

  const invoiceVatDeductible = (currentInvoices || []).reduce((sum, invoice) => {
    return sum + getInvoiceVatDeductibleAmount(invoice);
  }, 0);
  const noInvoiceVatDeductible = (currentNoInvoiceExpenses || []).reduce(
    (sum, expense) => sum + getNoInvoiceVatDeductibleAmount(expense),
    0
  );
  const invoiceVatDeductibleTotal =
    invoiceVatDeductible > 0 || !summaryVatTotal
      ? invoiceVatDeductible
      : Math.max(summaryVatTotal - noInvoiceVatDeductible, 0);
  const vatDeductible = invoiceVatDeductibleTotal + noInvoiceVatDeductible;

  const baseTotals = currentBillingSummary?.baseTotals || {};
  const vatTotals = currentBillingSummary?.vatTotals || {};
  const billingBase = (Number(baseTotals["0"]) || 0) +
    (Number(baseTotals["4"]) || 0) +
    (Number(baseTotals["10"]) || 0) +
    (Number(baseTotals["21"]) || 0);
  const billingVat = (Number(vatTotals["0"]) || 0) +
    (Number(vatTotals["4"]) || 0) +
    (Number(vatTotals["10"]) || 0) +
    (Number(vatTotals["21"]) || 0);

  const incomeInvoicesGross = (currentIncomeInvoices || []).reduce(
    (sum, invoice) => sum + (Number(invoice.total_amount) || 0),
    0
  );
  const incomeInvoicesVat = (currentIncomeInvoices || []).reduce(
    (sum, invoice) => sum + (Number(invoice.vat_amount) || 0),
    0
  );

  expenseGrossTotal = roundAmount(
    selectedMetrics ? Number(selectedMetrics.expense_gross_total) || 0 : expenseTotalGross
  );
  expenseVatSupportedTotal = roundAmount(
    selectedMetrics ? Number(selectedMetrics.expense_vat) || 0 : vatSupported
  );
  expenseVatDeductibleTotal = roundAmount(
    selectedMetrics ? Number(selectedMetrics.expense_vat) || 0 : vatDeductible
  );
  incomeGrossTotal = roundAmount(
    selectedMetrics ? Number(selectedMetrics.income_gross_total) || 0 : billingBase + billingVat + incomeInvoicesGross
  );
  incomeVatOutputTotal = roundAmount(
    selectedMetrics ? Number(selectedMetrics.income_vat) || 0 : billingVat + incomeInvoicesVat
  );
  expenseVatTotal = expenseVatDeductibleTotal;

  const expenseTotalEl = document.getElementById("expenseTotalGross");
  if (expenseTotalEl) {
    expenseTotalEl.textContent = formatCurrency(expenseGrossTotal);
  }
  const expenseVatSupportedEl = document.getElementById("expenseVatSupported");
  if (expenseVatSupportedEl) {
    expenseVatSupportedEl.textContent = formatCurrency(expenseVatSupportedTotal);
  }
  const expenseVatDeductibleEl = document.getElementById("expenseVatDeductible");
  if (expenseVatDeductibleEl) {
    expenseVatDeductibleEl.textContent = formatCurrency(expenseVatDeductibleTotal);
  }
  const incomeTotalEl = document.getElementById("incomeTotalGross");
  if (incomeTotalEl) {
    incomeTotalEl.textContent = formatCurrency(incomeGrossTotal);
  }
  const incomeVatOutputEl = document.getElementById("incomeVatOutput");
  if (incomeVatOutputEl) {
    incomeVatOutputEl.textContent = formatCurrency(incomeVatOutputTotal);
  }
  const vatResultLabelEl = document.getElementById("incomeVatResultLabel");
  const vatResultValueEl = document.getElementById("incomeVatResult");
  const vatResult = incomeVatOutputTotal - expenseVatDeductibleTotal;
  if (vatResultLabelEl) {
    vatResultLabelEl.textContent = vatResult >= 0 ? "IVA a pagar" : "IVA a devolver";
  }
  if (vatResultValueEl) {
    vatResultValueEl.textContent = formatCurrency(Math.abs(vatResult));
  }

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
    const incomeMap = {};
    months.forEach((month) => {
      expensesMap[month] = 0;
      incomeMap[month] = 0;
    });
    currentInvoices.forEach((invoice) => {
      const month = Number(String(invoice.invoice_date || "").slice(5, 7));
      if (!expensesMap[month]) {
        expensesMap[month] = 0;
      }
      expensesMap[month] += getInvoiceDeductibleAmount(invoice);
    });
    currentNoInvoiceExpenses.forEach((expense) => {
      if (!expense.deductible) {
        return;
      }
      const month = Number(String(expense.expense_date || "").slice(5, 7));
      if (!expensesMap[month]) {
        expensesMap[month] = 0;
      }
      expensesMap[month] += getNoInvoiceDeductibleAmount(expense);
    });
    currentLoanInstallments.forEach((installment) => {
      const month = Number(String(installment.payment_date || "").slice(5, 7));
      if (!expensesMap[month]) {
        expensesMap[month] = 0;
      }
      expensesMap[month] += Number(installment.interest_amount) || 0;
    });
    currentIncomeInvoices.forEach((invoice) => {
      const month = Number(String(invoice.invoice_date || "").slice(5, 7));
      if (!incomeMap[month]) {
        incomeMap[month] = 0;
      }
      incomeMap[month] += Number(invoice.base_amount) || 0;
    });

    labels = months.map((month) => monthNames[month - 1]);
    values = months.map((month) => {
      const income = currentBillingSummary.monthlyTotals.find(
        (entry) => entry.month === month
      )?.total || 0;
      const incomeTotal = income + (incomeMap[month] || 0);
      const expenses = expensesMap[month] || 0;
      return incomeTotal - expenses;
    });
  } else {
    const incomeInvoicesBase = currentIncomeInvoices.reduce(
      (sum, invoice) => sum + (Number(invoice.base_amount) || 0),
      0
    );
    const netValue = billingBaseTotal + incomeInvoicesBase - currentDeductibleExpenses;
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

function buildFiscalVatSummaryData(data) {
  const baseTotals = data.baseTotals || {};
  const vatTotals = data.vatTotals || {};
  const mergedBaseTotals = {
    "0": Number(baseTotals["0"]) || 0,
    "4": Number(baseTotals["4"]) || 0,
    "10": Number(baseTotals["10"]) || 0,
    "21": Number(baseTotals["21"]) || 0,
  };
  const mergedVatTotals = {
    "0": Number(vatTotals["0"]) || 0,
    "4": Number(vatTotals["4"]) || 0,
    "10": Number(vatTotals["10"]) || 0,
    "21": Number(vatTotals["21"]) || 0,
  };

  (currentIncomeInvoices || []).forEach((invoice) => {
    const breakdown = buildVatBreakdownPayload(
      parseVatBreakdown(invoice.vat_breakdown || invoice.vatBreakdown)
    );
    if (breakdown.length) {
      breakdown.forEach((line) => {
        const rateKey = String(Math.round(Number(line.rate) || 0));
        if (!(rateKey in mergedBaseTotals)) {
          return;
        }
        mergedBaseTotals[rateKey] += Number(line.base) || 0;
        mergedVatTotals[rateKey] += Number(line.vat_amount) || 0;
      });
      return;
    }

    const rateKey = resolveVatRateValue(invoice.vat_rate, "21");
    if (!(rateKey in mergedBaseTotals)) {
      return;
    }
    mergedBaseTotals[rateKey] += Number(invoice.base_amount) || 0;
    mergedVatTotals[rateKey] += Number(invoice.vat_amount) || 0;
  });

  return {
    baseTotals: {
      "0": Number(mergedBaseTotals["0"].toFixed(2)),
      "4": Number(mergedBaseTotals["4"].toFixed(2)),
      "10": Number(mergedBaseTotals["10"].toFixed(2)),
      "21": Number(mergedBaseTotals["21"].toFixed(2)),
    },
    vatTotals: {
      "0": Number(mergedVatTotals["0"].toFixed(2)),
      "4": Number(mergedVatTotals["4"].toFixed(2)),
      "10": Number(mergedVatTotals["10"].toFixed(2)),
      "21": Number(mergedVatTotals["21"].toFixed(2)),
    },
  };
}

function updateBillingSummary(data) {
  const merged = buildFiscalVatSummaryData(data);
  const baseTotals = merged.baseTotals;
  const vatTotals = merged.vatTotals;

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
  billingVatTotal =
    (Number(vatTotals["0"]) || 0) +
    (Number(vatTotals["4"]) || 0) +
    (Number(vatTotals["10"]) || 0) +
    (Number(vatTotals["21"]) || 0);
  updateDashboardTotals();
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
    paymentMonthTotal.textContent = `Total previsto del mes: ${formatCurrency(monthTotal)}`;
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
    const dayItems = itemsByDay[day] || [];
    if (dayItems.some((item) => item.status === "overdue")) {
      cell.classList.add("has-overdue-payments");
    } else if (dayItems.some((item) => item.status === "due_today")) {
      cell.classList.add("has-due-today-payments");
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
  renderPaymentDueAlert(data);
  renderPaymentDayDetails(selectedPaymentDay);
}

function getPaymentActionLabel(item) {
  return item.type === "income" ? "Cobrado" : "Pagado";
}

function getPaymentUndoActionLabel(item) {
  return item.type === "income" ? "No cobrado" : "No pagado";
}

function getAlertPendingPayments(data) {
  if (!Array.isArray(data?.items)) {
    return [];
  }
  return data.items.filter(
    (item) =>
      ["expense", "no_invoice", "loan_installment"].includes(item.type) &&
      (item.status === "due_today" || item.status === "overdue")
  );
}

function executePaymentStatusUpdate(item, payload, options = {}) {
  const { silent = false, skipRefresh = false } = options;
  const handleSuccess = () => {
    if (skipRefresh) {
      return;
    }
    refreshPayments();
    if (item.type === "loan_installment") {
      refreshLoanInstallments();
    }
  };
  const handleFailure = (data) => {
    if (!silent) {
      alert((data?.errors || ["Error al actualizar."]).join("\n"));
    }
  };
  const handleCatch = (message) => {
    if (!silent) {
      alert(message);
    }
  };

  if (item.type === "loan_installment") {
    return fetch(withCompanyParam(`/api/loan-installments/${item.id}`), {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    })
      .then((res) => res.json())
      .then((data) => {
        if (!data.ok) {
          handleFailure(data);
          return false;
        }
        handleSuccess();
        return true;
      })
      .catch(() => {
        handleCatch("No se pudo marcar el pago como realizado.");
        return false;
      });
  }
  if (item.type === "no_invoice") {
    return fetch(withCompanyParam(`/api/expenses/no-invoice/${item.id}`), {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    })
      .then((res) => res.json())
      .then((data) => {
        if (!data.ok) {
          handleFailure(data);
          return false;
        }
        handleSuccess();
        return true;
      })
      .catch(() => {
        handleCatch("No se pudo marcar el pago como realizado.");
        return false;
      });
  }
  if (item.type === "expense") {
    return fetch(withCompanyParam(`/api/invoices/${item.id}`), {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    })
      .then((res) => res.json())
      .then((data) => {
        if (!data.ok) {
          handleFailure(data);
          return false;
        }
        handleSuccess();
        return true;
      })
      .catch(() => {
        handleCatch("No se pudo marcar el pago como realizado.");
        return false;
      });
  }
  if (item.type === "income") {
    return fetch(withCompanyParam(`/api/income-invoices/${item.id}`), {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    })
      .then((res) => res.json())
      .then((data) => {
        if (!data.ok) {
          handleFailure(data);
          return false;
        }
        handleSuccess();
        return true;
      })
      .catch(() => {
        handleCatch("No se pudo marcar el cobro como realizado.");
        return false;
      });
  }
  return Promise.resolve(false);
}

function renderPaymentDueAlert(data) {
  if (!paymentDueAlert || !paymentDueAlertSummary || !paymentDueAlertList) {
    return;
  }
  const alertPendingPayments = getAlertPendingPayments(data);
  const todayPending = Array.isArray(data?.todayPending) ? data.todayPending : [];
  const overduePendingCount = Number(data?.overduePendingCount || 0);
  if (!todayPending.length && overduePendingCount <= 0) {
    paymentDueAlert.hidden = true;
    paymentDueAlertList.innerHTML = "";
    paymentDueAlertSummary.textContent = "";
    return;
  }

  paymentDueAlert.hidden = false;
  if (paymentClearAllBtn) {
    paymentClearAllBtn.style.display = alertPendingPayments.length ? "inline-flex" : "none";
  }
  paymentDueAlertList.innerHTML = "";
  if (todayPending.length) {
    todayPending.forEach((item) => {
      const row = document.createElement("div");
      row.className = "payment-alert-item";
      row.textContent = `${item.counterparty || item.concept || "Pago"} · ${
        item.payment_date
      } · ${formatCurrency(item.amount)}`;
      paymentDueAlertList.appendChild(row);
    });
  } else if (overduePendingCount > 0) {
    const row = document.createElement("div");
    row.className = "payment-alert-item";
    row.textContent = "No vencen pagos hoy, pero tienes pagos atrasados pendientes.";
    paymentDueAlertList.appendChild(row);
  }
  const summaryParts = [];
  if (todayPending.length) {
    summaryParts.push(
      `${todayPending.length} pago${todayPending.length === 1 ? "" : "s"} vencen hoy`
    );
  }
  if (overduePendingCount > 0) {
    summaryParts.push(
      `${overduePendingCount} pago${overduePendingCount === 1 ? "" : "s"} atrasado${
        overduePendingCount === 1 ? "" : "s"
      }`
    );
  }
  paymentDueAlertSummary.textContent = summaryParts.join(" · ");
}

function markPaymentsAsPaid(items) {
  const pendingItems = Array.isArray(items) ? items.filter(Boolean) : [];
  if (!pendingItems.length) {
    return Promise.resolve();
  }
  const updates = pendingItems.map((item) => {
    const payload = {
      payment_only: true,
      mark_paid: true,
      payment_date: item.payment_date,
      payment_reference_date: item.payment_date,
    };
    if (item.type === "no_invoice") {
      payload.expense_date = item.invoice_date;
    } else if (item.type === "expense") {
      payload.invoice_date = item.invoice_date;
    }
    return executePaymentStatusUpdate(item, payload, {
      silent: true,
      skipRefresh: true,
    });
  });

  return Promise.all(updates).then((results) => {
    if (results.every(Boolean)) {
      refreshPayments();
      if (pendingItems.some((item) => item.type === "loan_installment")) {
        refreshLoanInstallments();
      }
      return;
    }
    alert("No se pudieron marcar todos los pagos como realizados.");
    refreshPayments();
  });
}

function markPaymentAsPaid(item) {
  const payload = {
    payment_only: true,
    mark_paid: true,
    payment_date: item.payment_date,
    payment_reference_date: item.payment_date,
  };
  if (item.type === "no_invoice") {
    payload.expense_date = item.invoice_date;
  } else if (item.type === "expense" || item.type === "income") {
    payload.invoice_date = item.invoice_date;
  }
  return executePaymentStatusUpdate(item, payload);
}

function markPaymentAsUnpaid(item) {
  const payload = {
    payment_only: true,
    mark_unpaid: true,
    payment_date: item.payment_date,
    payment_reference_date: item.payment_date,
  };
  if (item.type === "no_invoice") {
    payload.expense_date = item.invoice_date;
  } else if (item.type === "expense" || item.type === "income") {
    payload.invoice_date = item.invoice_date;
  }
  return executePaymentStatusUpdate(item, payload);
}

function updatePaymentDateFromCalendar(item, newDate) {
  if (!newDate) {
    alert("Selecciona una fecha válida.");
    return Promise.resolve();
  }
  if (item.type === "loan_installment") {
    const payload = {
      payment_date: newDate,
      payment_reference_date: item.payment_date,
      payment_only: true,
    };
    const url = withCompanyParam(`/api/loan-installments/${item.id}`);
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
        refreshLoanInstallments();
      })
      .catch(() => {
        alert("No se pudo actualizar la fecha de vencimiento.");
      });
  }
  if (item.type === "no_invoice") {
    const payload = {
      expense_date: item.invoice_date,
      payment_date: newDate,
      payment_reference_date: item.payment_date,
      payment_only: true,
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
    payment_reference_date: item.payment_date,
    payment_dates: updatedDates,
    payment_only: true,
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
  paymentDayTitle.textContent = `Movimientos del ${day} ${monthLabel}`;

  let total = 0;
  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "payment-day-item";
    if (item.status) {
      row.classList.add(`is-${item.status}`);
    }
    const supplier = document.createElement("span");
    supplier.className = "payment-day-supplier";
    let label = "Proveedor";
    if (item.type === "income") {
      label = "Cliente";
    } else if (item.type === "no_invoice") {
      label = "Concepto";
    } else if (item.type === "loan_installment") {
      label = "Préstamo";
    } else if (item.type === "tax_obligation") {
      label = "Organismo";
    }
    supplier.textContent = `${label}: ${item.counterparty || "-"}`;
    const concept = document.createElement("span");
    concept.className = "payment-day-concept";
    if (item.type === "tax_obligation") {
      const filingLabel = taxFilingStatusLabel(item.filing_status);
      concept.textContent = filingLabel
        ? `${item.concept || "Modelo fiscal"} · ${filingLabel}`
        : item.concept || "Modelo fiscal";
    } else {
      concept.textContent = item.concept || "Factura";
    }
    const dateLabel = document.createElement("span");
    dateLabel.textContent = item.payment_date;
    if (item.status === "overdue") {
      dateLabel.className = "payment-status-label payment-status-overdue";
      dateLabel.textContent = `Atrasado · ${item.payment_date}`;
    } else if (item.status === "due_today") {
      dateLabel.className = "payment-status-label payment-status-due-today";
      dateLabel.textContent = `Vence hoy · ${item.payment_date}`;
    } else if (item.status === "pending") {
      dateLabel.className = "payment-status-label payment-status-pending";
      dateLabel.textContent = `Pendiente · ${item.payment_date}`;
    } else if (item.status === "paid") {
      dateLabel.className = "payment-status-label payment-status-paid";
      dateLabel.textContent = `Pagado · ${item.payment_date}`;
    }
    const amount = document.createElement("span");
    amount.className = "payment-day-amount";
    const amountLabel =
      item.type === "income"
        ? "Cobro"
        : item.type === "no_invoice"
        ? Number(item.withholding_amount || 0) > 0
          ? "Pago neto"
          : "Pago operativo"
        : item.type === "loan_installment"
        ? "Cuota préstamo"
        : item.type === "tax_obligation"
        ? item.calendar_impact === "credit"
          ? "Crédito fiscal"
          : item.calendar_impact === "refund"
          ? "Devolución solicitada"
          : item.calendar_impact === "informational"
          ? "Expediente sin cuota"
          : `Modelo ${item.tax_model || ""}`.trim()
        : "Gasto";
    amount.textContent =
      item.type === "tax_obligation" && item.calendar_impact === "informational"
        ? amountLabel
        : `${formatCurrency(item.amount)} (${amountLabel})`;
    const editContainer = document.createElement("div");
    editContainer.className = "payment-edit";
    const actionsGroup = document.createElement("div");
    actionsGroup.className = "payment-day-actions";
    let editBtn = null;
    if (item.type !== "tax_obligation") {
      editBtn = document.createElement("button");
      editBtn.type = "button";
      editBtn.className = "button ghost small";
      editBtn.textContent = "Editar fecha";
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
    }
    let paidBtn = null;
    if (["expense", "no_invoice", "loan_installment"].includes(item.type)) {
      paidBtn = document.createElement("button");
      paidBtn.type = "button";
      paidBtn.className =
        item.status === "paid" ? "button ghost small" : "button primary small";
      paidBtn.textContent =
        item.status === "paid"
          ? getPaymentUndoActionLabel(item)
          : getPaymentActionLabel(item);
      paidBtn.addEventListener("click", () => {
        if (item.status === "paid") {
          markPaymentAsUnpaid(item);
          return;
        }
        markPaymentAsPaid(item);
      });
    }
    row.appendChild(supplier);
    row.appendChild(concept);
    row.appendChild(dateLabel);
    row.appendChild(amount);
    if (paidBtn) {
      actionsGroup.appendChild(paidBtn);
    }
    if (editBtn) {
      actionsGroup.appendChild(editBtn);
    }
    if (actionsGroup.childElementCount) {
      row.appendChild(actionsGroup);
    }
    row.appendChild(editContainer);
    paymentDayList.appendChild(row);
    total += Number(item.amount || 0);
  });
  paymentDayTotal.textContent = `Total previsto del día: ${formatCurrency(total)}`;
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
    refreshLoanInstallments(),
    refreshAnnualTaxData(),
    loadAccountingIntegrationSummary(),
  ]).then(() => {
    updateDashboardTotals();
    updateDashboardEmptyState();
    renderArchive();
    const netResult =
      parseNumberInput(document.getElementById("pnlNet")?.textContent || "") || 0;
    syncBalanceStatementFields(netResult);
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
  const hasLoans =
    Array.isArray(currentLoanInstallments) && currentLoanInstallments.length > 0;
  const hasPayments =
    Array.isArray(currentPayments?.items) && currentPayments.items.length > 0;
  const hasData = hasExpenses || hasBilling || hasNoInvoice || hasLoans || hasPayments;
  emptyNode.style.display = hasData ? "none" : "block";
}

function updatePeriodBadge() {
  if (!taxPeriodBadge) {
    return;
  }
  const label = getPeriodLabel();
  taxPeriodBadge.textContent = label ? `Periodo: ${label}` : "";
}

function fetchFinancialMetrics() {
  const companyId = getSelectedCompanyId();
  const { month, year } = getSelectedMonthYear();
  const period = getSelectedPeriod();
  if (!companyId || !month || !year) {
    return Promise.resolve(null);
  }
  return fetch(
    `/api/financial-metrics?month=${month}&year=${year}&period=${period}&company_id=${companyId}`
  )
    .then((res) => res.json())
    .then((data) => (data?.ok ? data : null));
}

// IRPF e IS pasan a venir del backend para usar el mismo motor que calendario y modelos.
function refreshAnnualTaxData() {
  const { year } = getSelectedMonthYear();
  if (!year) {
    return Promise.resolve();
  }
  return fetchFinancialMetrics().then((data) => {
    currentFinancialMetrics = data;
    const fullYear = data?.fullYear || {};
    annualBillingBaseTotal = Number(fullYear.income_base) || 0;
    annualLoanInterestTotal = Number(fullYear.loan_interest) || 0;
    annualDeductibleExpenses = Number(fullYear.expense_base) || 0;
    annualTaxEstimateTotal =
      getSelectedCompanyType() === "individual"
        ? Number(fullYear.model_130_estimate) || 0
        : Number(fullYear.corporate_tax_estimate) || 0;
    pnlDataReady = Boolean(data?.selected && data?.fullYear);
    updateTaxSummary();
    updatePnlSummary();
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
      todayPending: data.todayPending || [],
      overduePendingCount: data.overduePendingCount || 0,
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
  const filteredInvoices = invoices.filter((invoice) =>
    recordMatchesSearch(
      [
        invoice.invoice_date,
        invoice.supplier,
        invoice.base_amount,
        invoice.vat_amount,
        invoice.total_amount,
        invoice.withholding_amount,
        formatExpenseCategory(invoice.expense_category || "with_invoice"),
      ],
      invoiceSearchInput?.value || ""
    )
  );
  if (!filteredInvoices.length) {
    invoicesEmpty.style.display = "block";
    invoicesEmpty.textContent = invoices.length
      ? "No hay facturas que coincidan con la búsqueda."
      : "No hay facturas para este período.";
    renderArchive();
    updateTaxSummary();
    updateDashboardTotals();
    updateSupplierSuggestions();
    return;
  }
  invoicesEmpty.style.display = "none";

  filteredInvoices.forEach((invoice) => {
    const tr = document.createElement("tr");
    tr.dataset.id = invoice.id;

    const dateTd = document.createElement("td");
    dateTd.textContent = invoice.invoice_date || invoice.payment_date || "";

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

    const withholdingTd = document.createElement("td");
    const withholdingValue = Number(invoice.withholding_amount || 0);
    withholdingTd.textContent =
      withholdingValue > 0 ? formatCurrency(withholdingValue) : "-";

    const totalTd = document.createElement("td");
    totalTd.textContent = formatCurrency(invoice.total_amount);

    const categoryTd = document.createElement("td");
    categoryTd.textContent = formatExpenseCategory(
      invoice.expense_category || "with_invoice"
    );
    if (invoice.vat_deductible === false && invoice.expense_category !== "non_deductible") {
      categoryTd.title = "Gasto deducible con IVA no deducible";
    }

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
    tr.appendChild(withholdingTd);
    tr.appendChild(totalTd);
    tr.appendChild(categoryTd);
    tr.appendChild(actionsTd);
    invoicesTableBody.appendChild(tr);
  });

  renderArchive();
  updateTaxSummary();
  updateDashboardTotals();
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
  const filteredInvoices = invoices.filter((invoice) =>
    recordMatchesSearch(
      [
        invoice.invoice_date,
        invoice.client,
        invoice.base_amount,
        invoice.vat_amount,
        invoice.total_amount,
      ],
      incomeSearchInput?.value || ""
    )
  );
  if (!filteredInvoices.length) {
    incomeInvoicesEmpty.style.display = "block";
    incomeInvoicesEmpty.textContent = invoices.length
      ? "No hay facturas emitidas que coincidan con la búsqueda."
      : "No hay facturas emitidas para este período.";
    if (currentBillingSummary) {
      updateBillingSummary(currentBillingSummary);
    }
    renderArchive();
    updateBillingChart();
    updateDashboardTotals();
    return;
  }
  incomeInvoicesEmpty.style.display = "none";

  filteredInvoices.forEach((invoice) => {
    const tr = document.createElement("tr");
    tr.dataset.id = invoice.id;

    const dateTd = document.createElement("td");
    dateTd.textContent = invoice.invoice_date || invoice.payment_date || "";

    const clientTd = document.createElement("td");
    clientTd.textContent = invoice.client || "Cliente pendiente";

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
  renderArchive();
  if (currentBillingSummary) {
    updateBillingSummary(currentBillingSummary);
  }
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
  dateInput.value = invoice.invoice_date || invoice.payment_date || "";

  const clientInput = document.createElement("input");
  clientInput.type = "text";
  clientInput.value = invoice.client;

  const baseInput = document.createElement("input");
  baseInput.type = "text";
  baseInput.min = "0";
  baseInput.value = formatAmountInput(invoice.base_amount);

  const vatSelect = createVatSelect(invoice.vat_rate);

  const vatAmountInput = document.createElement("input");
  vatAmountInput.type = "text";
  vatAmountInput.min = "0";
  vatAmountInput.value = formatAmountInput(invoice.vat_amount || 0);

  const totalInput = document.createElement("input");
  totalInput.type = "text";
  totalInput.min = "0";
  totalInput.value = formatAmountInput(invoice.total_amount);

  const calcInputs = {
    base: baseInput,
    vat: vatSelect,
    vatAmount: vatAmountInput,
    total: totalInput,
  };
  attachAmountInputBehavior(baseInput);
  attachAmountInputBehavior(vatAmountInput);
  attachAmountInputBehavior(totalInput);
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
  const withholdingTd = row.children[5];
  const totalTd = row.children[6];
  const categoryTd = row.children[7];
  const actionsTd = row.children[8];

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
  baseInput.type = "text";
  baseInput.min = "0";
  baseInput.value = formatAmountInput(invoice.base_amount);

  const vatSelect = createVatSelect(invoice.vat_rate);

  const vatAmountInput = document.createElement("input");
  vatAmountInput.type = "text";
  vatAmountInput.min = "0";
  vatAmountInput.value = formatAmountInput(invoice.vat_amount || 0);

  const totalInput = document.createElement("input");
  totalInput.type = "text";
  totalInput.min = "0";
  totalInput.value = formatAmountInput(invoice.total_amount);

  const withholdingInput = document.createElement("input");
  withholdingInput.type = "text";
  withholdingInput.min = "0";
  withholdingInput.value = formatAmountInput(invoice.withholding_amount || 0);

  const calcInputs = {
    base: baseInput,
    vat: vatSelect,
    vatAmount: vatAmountInput,
    total: totalInput,
  };
  attachAmountInputBehavior(baseInput);
  attachAmountInputBehavior(vatAmountInput);
  attachAmountInputBehavior(totalInput);
  attachAmountInputBehavior(withholdingInput);
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
  const vatDeductibleSelect = createDeductibleSelect(invoice.vat_deductible !== false);
  const vatDeductibleWrap = document.createElement("div");
  vatDeductibleWrap.className = "inline-stack";
  const vatDeductibleLabel = document.createElement("span");
  vatDeductibleLabel.className = "field-warning";
  vatDeductibleLabel.textContent = "IVA deducible";
  vatDeductibleWrap.appendChild(vatDeductibleLabel);
  vatDeductibleWrap.appendChild(vatDeductibleSelect);
  const syncInvoiceDeductibleControls = () => {
    if (categorySelect.value === "non_deductible") {
      vatDeductibleSelect.value = "false";
      vatDeductibleSelect.disabled = true;
      return;
    }
    vatDeductibleSelect.disabled = false;
  };
  categorySelect.addEventListener("change", syncInvoiceDeductibleControls);
  syncInvoiceDeductibleControls();

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
  withholdingTd.textContent = "";
  withholdingTd.appendChild(withholdingInput);
  totalTd.textContent = "";
  totalTd.appendChild(totalInput);
  categoryTd.textContent = "";
  categoryTd.appendChild(categorySelect);
  categoryTd.appendChild(vatDeductibleWrap);
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
      withholding_amount: withholdingInput.value,
      total_amount: totalInput.value,
      expense_category: categorySelect.value,
      vat_deductible: vatDeductibleSelect.value === "true",
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

function getFilteredNoInvoiceExpenses(expenses) {
  if (currentExpenseMode === "rent") {
    return expenses.filter((expense) =>
      ["alquiler_local", "alquiler_cabina"].includes(expense.expense_type)
    );
  }
  if (currentExpenseMode === "payroll") {
    return expenses.filter((expense) =>
      ["nomina", "seguridad_social"].includes(expense.expense_type)
    );
  }
  if (currentExpenseMode === "financing") {
    return expenses.filter((expense) => expense.expense_type === "prestamo");
  }
  return expenses.filter(
    (expense) => ["amortizacion", "kilometraje", "otro"].includes(expense.expense_type)
  );
}

function renderNoInvoiceExpenses(expenses) {
  noInvoiceTableBody.innerHTML = "";
  currentNoInvoiceExpenses = expenses;
  const filteredExpenses = getFilteredNoInvoiceExpenses(expenses).filter((expense) =>
    recordMatchesSearch(
      [
        expense.expense_date,
        expense.concept,
        expense.payroll_employee_name,
        expense.payroll_period,
        expense.amount,
        expense.interest_amount,
        expense.withholding_amount,
        formatNoInvoiceType(expense.expense_type),
      ],
      noInvoiceSearchInput?.value || ""
    )
  );
  if (!filteredExpenses.length) {
    noInvoiceEmpty.style.display = "block";
    const emptyByMode =
      currentExpenseMode === "rent"
        ? "No hay gastos de alquiler en este período."
        : currentExpenseMode === "payroll"
          ? "No hay gastos de personal en este período."
          : currentExpenseMode === "financing"
            ? "No hay movimientos de financiación en este período."
            : "No hay otros ajustes en este período.";
    noInvoiceEmpty.textContent =
      expenses.length && (noInvoiceSearchInput?.value || "").trim()
        ? "No hay resultados para ese filtro."
        : emptyByMode;
    renderArchive();
    updateTaxSummary();
    updateDashboardTotals();
    return;
  }
  noInvoiceEmpty.style.display = "none";

  filteredExpenses.forEach((expense) => {
    const tr = document.createElement("tr");
    tr.dataset.id = expense.id;

    const dateTd = document.createElement("td");
    dateTd.textContent = expense.expense_date;

    const conceptTd = document.createElement("td");
    conceptTd.textContent = expense.concept;

    const amountTd = document.createElement("td");
    amountTd.textContent = formatCurrency(expense.amount);

    const vatRateTd = document.createElement("td");
    const vatAmountTd = document.createElement("td");
    if (expense.vat_deductible) {
      const rateValue =
        expense.vat_rate ?? expense.vat_rate_no_invoice ?? null;
      const vatValue =
        expense.vat_amount ?? expense.vat_amount_no_invoice ?? 0;
      vatRateTd.textContent = formatPercent(rateValue);
      vatAmountTd.textContent = formatCurrency(vatValue);
    } else {
      vatRateTd.textContent = "-";
      vatAmountTd.textContent = "-";
    }

    const interestTd = document.createElement("td");
    const interestValue = Number(expense.interest_amount || 0);
    interestTd.textContent =
      expense.expense_type === "prestamo" ? formatCurrency(interestValue) : "-";

    const withholdingTd = document.createElement("td");
    const withholdingValue = Number(expense.withholding_amount || 0);
    withholdingTd.textContent =
      withholdingValue > 0 ? formatCurrency(withholdingValue) : "-";

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
    tr.appendChild(vatRateTd);
    tr.appendChild(vatAmountTd);
    tr.appendChild(interestTd);
    tr.appendChild(withholdingTd);
    tr.appendChild(typeTd);
    tr.appendChild(deductibleTd);
    tr.appendChild(actionsTd);
    noInvoiceTableBody.appendChild(tr);
  });

  renderArchive();
  updateTaxSummary();
  updateDashboardTotals();
}

function syncLoanPrincipal() {
  if (!loanTotalInput || !loanInterestInput || !loanPrincipalInput) {
    return;
  }
  const totalValue = parseNumberInput(loanTotalInput.value) || 0;
  const interestValue = parseNumberInput(loanInterestInput.value) || 0;
  const principalValue = Math.max(totalValue - interestValue, 0);
  loanPrincipalInput.value = formatAmountInput(principalValue);
}

function fetchLoanInstallments(month, year) {
  const companyId = getSelectedCompanyId();
  const suffix = companyId ? `&company_id=${companyId}` : "";
  return fetch(`/api/loan-installments?month=${month}&year=${year}${suffix}`)
    .then((res) => res.json())
    .then((data) => data.installments || []);
}

function refreshLoanInstallments() {
  if (!loanTableBody) {
    return Promise.resolve();
  }
  const { month, year } = getSelectedMonthYear();
  if (!month || !year) {
    return Promise.resolve();
  }
  const months = getPeriodMonths();
  return Promise.all(months.map((targetMonth) => fetchLoanInstallments(targetMonth, year)))
    .then((installmentsByMonth) => {
      const installments = installmentsByMonth.flat();
      installments.sort((a, b) => b.payment_date.localeCompare(a.payment_date));
      renderLoanInstallments(installments);
    });
}

function resetLoanPlanPreview() {
  loanPlanDraft = [];
  if (loanPlanPreviewBody) {
    loanPlanPreviewBody.innerHTML = "";
  }
  if (loanPlanPreviewEmpty) {
    loanPlanPreviewEmpty.style.display = "block";
  }
}

function renderLoanInstallments(installments) {
  if (!loanTableBody || !loanEmpty) {
    return;
  }
  loanTableBody.innerHTML = "";
  currentLoanInstallments = installments;
  const filteredInstallments = installments.filter((installment) =>
    recordMatchesSearch(
      [
        installment.payment_date,
        installment.bank_name,
        installment.concept,
        installment.total_amount,
        installment.interest_amount,
        installment.principal_amount,
      ],
      loanSearchInput?.value || ""
    )
  );
  if (!filteredInstallments.length) {
    loanEmpty.style.display = "block";
    loanEmpty.textContent = installments.length
      ? "No hay cuotas que coincidan con la búsqueda."
      : "No hay cuotas registradas en este período.";
    renderArchive();
    updateTaxSummary();
    updateDashboardTotals();
    return;
  }
  loanEmpty.style.display = "none";

  filteredInstallments.forEach((installment) => {
    const tr = document.createElement("tr");
    tr.dataset.id = installment.id;

    const dateTd = document.createElement("td");
    dateTd.textContent = installment.payment_date;

    const bankTd = document.createElement("td");
    bankTd.textContent = installment.bank_name || "-";

    const conceptTd = document.createElement("td");
    conceptTd.textContent = installment.concept;

    const totalTd = document.createElement("td");
    totalTd.textContent = formatCurrency(installment.total_amount);

    const interestTd = document.createElement("td");
    interestTd.textContent = formatCurrency(installment.interest_amount);

    const principalTd = document.createElement("td");
    principalTd.textContent = formatCurrency(installment.principal_amount);

    const actionsTd = document.createElement("td");
    actionsTd.classList.add("billing-actions");

    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "button ghost";
    editBtn.textContent = "Editar";
    editBtn.addEventListener("click", () => {
      enterLoanEditMode(tr, installment);
    });

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "button danger";
    deleteBtn.textContent = "Eliminar";
    deleteBtn.addEventListener("click", () => {
      if (!confirm("¿Seguro que deseas eliminar esta cuota?")) {
        return;
      }
      deleteLoanInstallment(installment.id);
    });

    actionsTd.appendChild(editBtn);
    actionsTd.appendChild(deleteBtn);

    tr.appendChild(dateTd);
    tr.appendChild(bankTd);
    tr.appendChild(conceptTd);
    tr.appendChild(totalTd);
    tr.appendChild(interestTd);
    tr.appendChild(principalTd);
    tr.appendChild(actionsTd);
    loanTableBody.appendChild(tr);
  });

  renderArchive();
  updateTaxSummary();
  updateDashboardTotals();
}

function enterLoanEditMode(row, installment) {
  row.innerHTML = "";

  const dateTd = document.createElement("td");
  const dateInput = document.createElement("input");
  dateInput.type = "date";
  dateInput.value = installment.payment_date;
  dateTd.appendChild(dateInput);

  const bankTd = document.createElement("td");
  const bankInput = document.createElement("input");
  bankInput.type = "text";
  bankInput.value = installment.bank_name || "";
  bankTd.appendChild(bankInput);

  const conceptTd = document.createElement("td");
  const conceptInput = document.createElement("input");
  conceptInput.type = "text";
  conceptInput.value = installment.concept || "";
  conceptTd.appendChild(conceptInput);

  const totalTd = document.createElement("td");
  const totalInput = document.createElement("input");
  totalInput.type = "text";
  totalInput.min = "0";
  totalInput.value = formatAmountInput(installment.total_amount);
  totalTd.appendChild(totalInput);

  const interestTd = document.createElement("td");
  const interestInput = document.createElement("input");
  interestInput.type = "text";
  interestInput.min = "0";
  interestInput.value = formatAmountInput(installment.interest_amount);
  interestTd.appendChild(interestInput);

  const principalTd = document.createElement("td");
  const principalInput = document.createElement("input");
  principalInput.type = "text";
  principalInput.readOnly = true;
  const updatePrincipal = () => {
    const totalValue = parseNumberInput(totalInput.value) || 0;
    const interestValue = parseNumberInput(interestInput.value) || 0;
    principalInput.value = formatAmountInput(Math.max(totalValue - interestValue, 0));
  };
  totalInput.addEventListener("input", updatePrincipal);
  interestInput.addEventListener("input", updatePrincipal);
  updatePrincipal();
  principalTd.appendChild(principalInput);

  const actionsTd = document.createElement("td");
  const saveBtn = document.createElement("button");
  saveBtn.type = "button";
  saveBtn.className = "button primary";
  saveBtn.textContent = "Guardar";
  saveBtn.addEventListener("click", () => {
    updateLoanInstallment(installment.id, {
      payment_date: dateInput.value,
      bank_name: bankInput.value,
      concept: conceptInput.value,
      total_amount: parseNumberInput(totalInput.value),
      interest_amount: parseNumberInput(interestInput.value),
    });
  });
  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "button ghost";
  cancelBtn.textContent = "Cancelar";
  cancelBtn.addEventListener("click", () => {
    refreshLoanInstallments();
  });
  actionsTd.appendChild(saveBtn);
  actionsTd.appendChild(cancelBtn);

  row.appendChild(dateTd);
  row.appendChild(bankTd);
  row.appendChild(conceptTd);
  row.appendChild(totalTd);
  row.appendChild(interestTd);
  row.appendChild(principalTd);
  row.appendChild(actionsTd);
}

function saveLoanInstallment() {
  if (!loanConceptInput || !loanPaymentDateInput || !loanTotalInput || !loanInterestInput) {
    return;
  }
  const payload = {
    concept: loanConceptInput.value.trim(),
    payment_date: loanPaymentDateInput.value,
    total_amount: parseNumberInput(loanTotalInput.value),
    interest_amount: parseNumberInput(loanInterestInput.value),
    company_id: getSelectedCompanyId(),
  };
  fetch(withCompanyParam("/api/loan-installments"), {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al guardar la cuota."]).join("\n"));
        return;
      }
      loanTotalInput.value = "";
      loanInterestInput.value = "";
      loanPrincipalInput.value = "";
      refreshLoanInstallments();
      refreshPayments();
    })
    .catch(() => {
      alert("No se pudo guardar la cuota.");
    });
}

function updateLoanInstallment(id, payload) {
  fetch(withCompanyParam(`/api/loan-installments/${id}`), {
    method: "PUT",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ ...payload, company_id: getSelectedCompanyId() }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al actualizar la cuota."]).join("\n"));
        return;
      }
      refreshLoanInstallments();
      refreshPayments();
    })
    .catch(() => {
      alert("No se pudo actualizar la cuota.");
    });
}

function deleteLoanInstallment(id) {
  fetch(withCompanyParam(`/api/loan-installments/${id}`), {
    method: "DELETE",
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al eliminar la cuota."]).join("\n"));
        return;
      }
      refreshLoanInstallments();
      refreshPayments();
    })
    .catch(() => {
      alert("No se pudo eliminar la cuota.");
    });
}

function importLoanPlan() {
  const file =
    (loanPlanFile && loanPlanFile.files && loanPlanFile.files[0]) ||
    selectedLoanPlanFile;
  if (!file) {
    alert("Selecciona un archivo PDF o Excel.");
    return;
  }
  const formData = new FormData();
  formData.append("file", file);
  fetch(withCompanyParam("/api/loan-installments/import?preview=1"), {
    method: "POST",
    body: formData,
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["No se pudo importar el plan."]).join("\n"));
        return;
      }
      loanPlanDraft = Array.isArray(data.installments) ? data.installments : [];
      renderLoanPlanPreview();
    })
    .catch(() => {
      alert("No se pudo importar el plan.");
    });
}

function renderLoanPlanPreview() {
  if (!loanPlanPreviewBody) {
    return;
  }
  loanPlanPreviewBody.innerHTML = "";

  if (!loanPlanDraft.length) {
    if (loanPlanPreviewEmpty) {
      loanPlanPreviewEmpty.style.display = "block";
    }
    return;
  }

  if (loanPlanPreviewEmpty) {
    loanPlanPreviewEmpty.style.display = "none";
  }

  loanPlanDraft.forEach((item, index) => {
    const row = document.createElement("tr");

    const dateTd = document.createElement("td");
    const dateInput = document.createElement("input");
    dateInput.type = "date";
    dateInput.value = item.payment_date || "";
    dateInput.addEventListener("input", () => {
      loanPlanDraft[index].payment_date = dateInput.value;
    });
    dateTd.appendChild(dateInput);

    const bankTd = document.createElement("td");
    const bankInput = document.createElement("input");
    bankInput.type = "text";
    bankInput.value = item.bank_name || "";
    bankInput.addEventListener("input", () => {
      loanPlanDraft[index].bank_name = bankInput.value;
    });
    bankTd.appendChild(bankInput);

    const conceptTd = document.createElement("td");
    const conceptInput = document.createElement("input");
    conceptInput.type = "text";
    conceptInput.value = item.concept || "Plan de amortización";
    conceptInput.addEventListener("input", () => {
      loanPlanDraft[index].concept = conceptInput.value;
    });
    conceptTd.appendChild(conceptInput);

    const totalTd = document.createElement("td");
    const totalInput = document.createElement("input");
    totalInput.type = "text";
    totalInput.min = "0";
    totalInput.value = formatAmountInput(item.total_amount);
    attachAmountInputBehavior(totalInput);
    totalInput.addEventListener("input", () => {
      loanPlanDraft[index].total_amount = parseNumberInput(totalInput.value);
    });
    totalTd.appendChild(totalInput);

    const interestTd = document.createElement("td");
    const interestInput = document.createElement("input");
    interestInput.type = "text";
    interestInput.min = "0";
    interestInput.value = formatAmountInput(item.interest_amount);
    attachAmountInputBehavior(interestInput);
    interestInput.addEventListener("input", () => {
      loanPlanDraft[index].interest_amount = parseNumberInput(interestInput.value);
    });
    interestTd.appendChild(interestInput);

    const principalTd = document.createElement("td");
    const principalInput = document.createElement("input");
    principalInput.type = "text";
    principalInput.min = "0";
    principalInput.value = formatAmountInput(item.principal_amount);
    attachAmountInputBehavior(principalInput);
    principalInput.addEventListener("input", () => {
      loanPlanDraft[index].principal_amount = parseNumberInput(principalInput.value);
    });
    principalTd.appendChild(principalInput);

    const actionsTd = document.createElement("td");
    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "button danger";
    removeBtn.textContent = "Quitar";
    removeBtn.addEventListener("click", () => {
      loanPlanDraft.splice(index, 1);
      renderLoanPlanPreview();
    });
    actionsTd.appendChild(removeBtn);

    row.appendChild(dateTd);
    row.appendChild(bankTd);
    row.appendChild(conceptTd);
    row.appendChild(totalTd);
    row.appendChild(interestTd);
    row.appendChild(principalTd);
    row.appendChild(actionsTd);

    loanPlanPreviewBody.appendChild(row);
  });
}

function saveLoanPlanDraft() {
  if (!loanPlanDraft.length) {
    alert("No hay entradas para guardar.");
    return;
  }
  const cleaned = loanPlanDraft
    .map((item) => ({
      payment_date: item.payment_date,
      bank_name: (item.bank_name || "").trim(),
      concept: (item.concept || "Plan de amortización").trim(),
      total_amount: Number(item.total_amount),
      interest_amount: Number(item.interest_amount || 0),
      principal_amount: Number(item.principal_amount || 0),
    }))
    .filter((item) => item.payment_date && !Number.isNaN(item.total_amount));

  if (!cleaned.length) {
    alert("Revisa las fechas y los importes antes de guardar.");
    return;
  }

  fetch(withCompanyParam("/api/loan-installments/batch"), {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ installments: cleaned, company_id: getSelectedCompanyId() }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["No se pudo guardar el plan."]).join("\n"));
        return;
      }
      loanPlanDraft = [];
      renderLoanPlanPreview();
      if (loanPlanFile) {
        loanPlanFile.value = "";
      }
      selectedLoanPlanFile = null;
      if (loanPlanFileName) {
        loanPlanFileName.textContent = "Ningún archivo seleccionado.";
      }
      refreshLoanInstallments();
      refreshPayments();
    })
    .catch(() => {
      alert("No se pudo guardar el plan.");
    });
}

function enterNoInvoiceEditMode(row, expense) {
  const dateTd = row.children[0];
  const conceptTd = row.children[1];
  const amountTd = row.children[2];
  const vatRateTd = row.children[3];
  const vatAmountTd = row.children[4];
  const interestTd = row.children[5];
  const withholdingTd = row.children[6];
  const typeTd = row.children[7];
  const deductibleTd = row.children[8];
  const actionsTd = row.children[9];

  const dateInput = document.createElement("input");
  dateInput.type = "date";
  dateInput.value = expense.expense_date;

  const conceptInput = document.createElement("input");
  conceptInput.type = "text";
  conceptInput.value = expense.concept;

  const amountInput = document.createElement("input");
  amountInput.type = "text";
  amountInput.min = "0";
  amountInput.value = expense.amount;

  const vatRateSelect = createVatSelect(
    expense.vat_deductible ? expense.vat_rate : ""
  );
  if (!expense.vat_deductible) {
    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    emptyOption.textContent = "-";
    vatRateSelect.insertBefore(emptyOption, vatRateSelect.firstChild);
    vatRateSelect.value = "";
  }
  const vatAmountInput = document.createElement("input");
  vatAmountInput.type = "text";
  vatAmountInput.min = "0";
  vatAmountInput.value =
    expense.vat_deductible && expense.vat_amount
      ? Number(expense.vat_amount).toFixed(2)
      : "";
  vatAmountInput.readOnly = true;

  const interestInput = document.createElement("input");
  interestInput.type = "text";
  interestInput.min = "0";
  interestInput.value =
    expense.expense_type === "prestamo" ? Number(expense.interest_amount || 0) : "";

  const withholdingInput = document.createElement("input");
  withholdingInput.type = "text";
  withholdingInput.min = "0";
  withholdingInput.value = formatAmountInput(expense.withholding_amount || 0);
  attachAmountInputBehavior(withholdingInput);

  const typeSelect = createNoInvoiceTypeSelect(
    expense.expense_type,
    getAllowedNoInvoiceTypes(currentExpenseMode)
  );
  const deductibleSelect = createDeductibleSelect(expense.deductible);

  dateTd.textContent = "";
  dateTd.appendChild(dateInput);
  conceptTd.textContent = "";
  conceptTd.appendChild(conceptInput);
  amountTd.textContent = "";
  amountTd.appendChild(amountInput);
  vatRateTd.textContent = "";
  vatRateTd.appendChild(vatRateSelect);
  vatAmountTd.textContent = "";
  vatAmountTd.appendChild(vatAmountInput);
  interestTd.textContent = "";
  interestTd.appendChild(interestInput);
  withholdingTd.textContent = "";
  withholdingTd.appendChild(withholdingInput);
  typeTd.textContent = "";
  typeTd.appendChild(typeSelect);
  deductibleTd.textContent = "";
  deductibleTd.appendChild(deductibleSelect);

  toggleLoanInterestField({
    typeValue: typeSelect.value,
    interestField: interestTd,
    interestInput,
    deductibleSelect,
  });
  toggleNoInvoiceWithholdingField({
    typeValue: typeSelect.value,
    field: withholdingTd,
    input: withholdingInput,
  });
  const applyNoInvoiceVatEdit = () => {
    const vatBlocked = ["prestamo", "nomina", "seguridad_social"].includes(typeSelect.value);
    vatRateSelect.disabled = vatBlocked;
    if (vatBlocked) {
      vatRateSelect.value = "";
      vatAmountInput.value = "";
      return;
    }
    if (!vatRateSelect.value) {
      vatAmountInput.value = "";
      return;
    }
    const computed = calculateVatFields({
      baseValue: null,
      totalValue: parseNumberInput(amountInput.value),
      vatRateValue: resolveVatRateValue(vatRateSelect.value),
      source: "total",
    });
    vatAmountInput.value =
      computed.vatAmount !== null ? formatAmountInput(computed.vatAmount) : "";
  };
  applyNoInvoiceVatEdit();
  typeSelect.addEventListener("change", () => {
    toggleLoanInterestField({
      typeValue: typeSelect.value,
      interestField: interestTd,
      interestInput,
      deductibleSelect,
    });
    toggleNoInvoiceWithholdingField({
      typeValue: typeSelect.value,
      field: withholdingTd,
      input: withholdingInput,
    });
    applyNoInvoiceVatEdit();
  });
  amountInput.addEventListener("input", applyNoInvoiceVatEdit);
  vatRateSelect.addEventListener("change", applyNoInvoiceVatEdit);

  actionsTd.innerHTML = "";
  const saveBtn = document.createElement("button");
  saveBtn.type = "button";
  saveBtn.className = "button primary";
  saveBtn.textContent = "Guardar";
  saveBtn.addEventListener("click", () => {
    const vatRateValue = vatRateSelect.value;
    const hasVat = !!vatRateValue;
    const calculated = calculateVatFields({
      baseValue: null,
      totalValue: parseNumberInput(amountInput.value),
      vatRateValue: resolveVatRateValue(vatRateValue),
      source: "total",
    });
    updateNoInvoiceExpense(expense.id, {
      expense_date: dateInput.value,
      concept: conceptInput.value,
      amount: amountInput.value,
      vat_deductible: hasVat && typeSelect.value !== "prestamo",
      vat_rate: hasVat ? vatRateValue : null,
      vat_amount: calculated.vatAmount,
      base_amount: calculated.base,
      interest_amount: interestInput.value,
      withholding_amount: withholdingInput.value,
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
  const vatDeductibleValue = noInvoiceVatDeductible?.value === "true";
  const interestValue = noInvoiceInterest ? noInvoiceInterest.value : "";
  const withholdingValue = noInvoiceWithholding ? noInvoiceWithholding.value : "";
  const amountNumeric = parseNumberInput(amountValue);
  const withholdingNumeric = parseNumberInput(withholdingValue) || 0;

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
  if (typeValue === "prestamo") {
    const interestNumeric = parseNumberInput(interestValue) || 0;
    if (interestNumeric < 0) {
      alert("Interés inválido.");
      return;
    }
    if (amountNumeric !== null && interestNumeric > amountNumeric) {
      alert("El interés no puede superar el importe.");
      return;
    }
  }
  if (withholdingNumeric < 0 || (amountNumeric !== null && withholdingNumeric > amountNumeric)) {
    alert("La retención no puede superar el importe.");
    return;
  }

  let vatRateValue = null;
  let vatAmountValue = null;
  let baseAmountValue = null;
  if (vatDeductibleValue && typeValue !== "prestamo") {
    vatRateValue = resolveVatRateValue(noInvoiceVatSelect?.value);
    const computed = calculateVatFields({
      baseValue: null,
      totalValue: parseNumberInput(amountValue),
      vatRateValue,
      source: "total",
    });
    baseAmountValue = computed.base;
    vatAmountValue = computed.vatAmount;
    if (vatAmountValue === null) {
      alert("No se pudo calcular el IVA deducible.");
      return;
    }
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
      vat_deductible: vatDeductibleValue && typeValue !== "prestamo",
      vat_rate: vatRateValue,
      vat_amount: vatAmountValue,
      base_amount: baseAmountValue,
      interest_amount: interestValue,
      withholding_amount: withholdingValue,
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
      if (noInvoiceInterest) {
        noInvoiceInterest.value = "";
      }
      if (noInvoiceWithholding) {
        noInvoiceWithholding.value = "";
      }
      if (noInvoiceVatDeductible) {
        noInvoiceVatDeductible.value = "false";
      }
      if (noInvoiceVatSelect) {
        noInvoiceVatSelect.value = "21";
      }
      if (noInvoiceVatBase) {
        noInvoiceVatBase.value = "";
      }
      if (noInvoiceVatAmount) {
        noInvoiceVatAmount.value = "";
      }
      toggleNoInvoiceVatFields({
        typeValue: noInvoiceType?.value,
        vatDeductibleValue: noInvoiceVatDeductible?.value,
      });
      Promise.all([refreshNoInvoiceExpenses(), refreshPayments()]);
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
      Promise.all([refreshNoInvoiceExpenses(), refreshPayments()]);
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
      Promise.all([refreshNoInvoiceExpenses(), refreshPayments()]);
    })
    .catch(() => {
      alert("No se pudo eliminar el gasto.");
    });
}

function renderBillingEntries(entries) {
  billingEntriesBody.innerHTML = "";
  currentBillingEntries = entries;
  const filteredEntries = entries.filter((entry) =>
    recordMatchesSearch(
      [
        entry.concept,
        entry.invoice_date,
        entry.month ? monthNames[Number(entry.month) - 1] : "",
        entry.year,
        entry.base,
        entry.vat,
        entry.vatAmount,
        entry.total,
      ],
      billingSearchInput?.value || ""
    )
  );
  if (!filteredEntries.length) {
    billingEntriesEmpty.style.display = "block";
    billingEntriesEmpty.textContent = entries.length
      ? "No hay resultados para ese filtro."
      : "No hay entradas de facturación.";
    renderArchive();
    updateBillingChart();
    return;
  }
  billingEntriesEmpty.style.display = "none";

  const showMonth = getSelectedPeriod() === "quarterly";
  filteredEntries.forEach((entry) => {
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
    vatTd.textContent = formatPercent(entry.vat);

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
  renderArchive();
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
  baseInput.type = "text";
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
  const outputTotal = Number(incomeVatOutputTotal || billingVatTotal || 0);
  const inputTotal = Number(expenseVatDeductibleTotal || expenseVatTotal || 0);
  const result = outputTotal - inputTotal;
  document.getElementById("vatOutputTotal").textContent = formatCurrency(
    outputTotal
  );
  document.getElementById("vatInputTotal").textContent = formatCurrency(
    inputTotal
  );
  document.getElementById("vatResultLabel").textContent =
    result >= 0 ? "IVA A PAGAR" : "IVA A DEVOLVER";
  document.getElementById("vatResultValue").textContent = formatCurrency(
    Math.abs(result)
  );
}

function updateTaxSummary() {
  const selectedMetrics = currentFinancialMetrics?.selected || null;
  const fullYearMetrics = currentFinancialMetrics?.fullYear || null;
  const deductibleInvoices = selectedMetrics
    ? Number(selectedMetrics.invoice_expenses) || 0
    : currentInvoices.reduce((total, invoice) => {
        return total + getInvoiceDeductibleAmount(invoice);
      }, 0);

  const deductibleNoInvoice = selectedMetrics
    ? Math.max(
        (Number(selectedMetrics.expense_base) || 0) -
          (Number(selectedMetrics.invoice_expenses) || 0) -
          (Number(selectedMetrics.loan_interest) || 0),
        0
      )
    : currentNoInvoiceExpenses.reduce(
        (total, expense) => total + getNoInvoiceDeductibleAmount(expense),
        0
      );
  const loanInterestPeriod = selectedMetrics
    ? Number(selectedMetrics.loan_interest) || 0
    : currentLoanInstallments.reduce(
        (total, installment) => total + (Number(installment.interest_amount) || 0),
        0
      );

  const periodExpenses = selectedMetrics
    ? Number(selectedMetrics.expense_base) || 0
    : deductibleInvoices + deductibleNoInvoice + loanInterestPeriod;
  currentDeductibleExpenses = periodExpenses;

  const annualIncome = fullYearMetrics ? Number(fullYearMetrics.income_base) || 0 : annualBillingBaseTotal;
  const annualExpenses = fullYearMetrics ? Number(fullYearMetrics.expense_base) || 0 : annualDeductibleExpenses;
  const annualOperatingExpenses = annualExpenses - annualLoanInterestTotal;
  const annualOperatingResult = annualIncome - annualOperatingExpenses;
  const annualNet = annualIncome - annualExpenses;
  const companyType = getSelectedCompanyType();
  const annualTaxEstimate = fullYearMetrics
    ? companyType === "individual"
      ? Number(fullYearMetrics.model_130_estimate) || 0
      : Number(fullYearMetrics.corporate_tax_estimate) || 0
    : annualNet > 0
      ? annualNet * (companyType === "company" ? 0.25 : companyType === "individual" ? 0.15 : 0)
      : 0;

  document.getElementById("irpfIncome").textContent = formatCurrency(annualIncome);
  document.getElementById("irpfExpenses").textContent = formatCurrency(annualExpenses);
  document.getElementById("irpfNet").textContent = formatCurrency(annualNet);

  document.getElementById("isIncome").textContent = formatCurrency(annualIncome);
  document.getElementById("isExpenses").textContent = formatCurrency(
    annualOperatingResult
  );
  document.getElementById("isResult").textContent = formatCurrency(annualNet);
  document.getElementById("isBase").textContent = formatCurrency(annualTaxEstimate);

  updatePnlSummary();
  updateNetChart();
  renderFiscalModelsSummary();
}

function updatePnlSummary() {
  const selectedMetrics = currentFinancialMetrics?.selected || null;
  const fullYearMetrics = currentFinancialMetrics?.fullYear || null;
  const useAnnual = pnlDataReady && Boolean(fullYearMetrics);
  const sourceInvoices = useAnnual ? pnlInvoices : currentInvoices;
  const sourceNoInvoice = useAnnual ? pnlNoInvoiceExpenses : currentNoInvoiceExpenses;
  const sourceLoans = useAnnual ? pnlLoanInstallments : currentLoanInstallments;
  const sourceIncome = useAnnual ? pnlIncomeInvoices : currentIncomeInvoices;

  const incomeInvoicesBase = sourceIncome.reduce(
    (sum, invoice) => sum + (Number(invoice.base_amount) || 0),
    0
  );
  const incomeTotal = useAnnual
    ? Number(fullYearMetrics.income_base) || 0
    : selectedMetrics
      ? Number(selectedMetrics.income_base) || 0
      : billingBaseTotal + incomeInvoicesBase;
  const loanInterest = useAnnual
    ? Number(fullYearMetrics.loan_interest) || 0
    : selectedMetrics
      ? Number(selectedMetrics.loan_interest) || 0
      : sourceNoInvoice.reduce((sum, expense) => {
          if (expense.expense_type === "prestamo") {
            return sum + (Number(expense.interest_amount) || 0);
          }
          return sum;
        }, 0) + sourceLoans.reduce(
          (sum, installment) => sum + (Number(installment.interest_amount) || 0),
          0
        );
  const invoiceExpenses = useAnnual
    ? Number(fullYearMetrics.invoice_expenses) || 0
    : selectedMetrics
      ? Number(selectedMetrics.invoice_expenses) || 0
      : sourceInvoices.reduce((sum, invoice) => {
          return sum + getInvoiceDeductibleAmount(invoice);
        }, 0);
  const payrollExpenses = useAnnual
    ? Number(fullYearMetrics.payroll_expenses) || 0
    : selectedMetrics
      ? Number(selectedMetrics.payroll_expenses) || 0
      : sourceNoInvoice.reduce((sum, expense) => {
          if (expense.expense_type === "nomina" || expense.expense_type === "seguridad_social") {
            return sum + getNoInvoiceDeductibleAmount(expense);
          }
          return sum;
        }, 0);
  const amortizationExpenses = useAnnual
    ? Number(fullYearMetrics.amortization_expenses) || 0
    : selectedMetrics
      ? Number(selectedMetrics.amortization_expenses) || 0
      : sourceNoInvoice.reduce((sum, expense) => {
          if (expense.expense_type === "amortizacion") {
            return sum + getNoInvoiceDeductibleAmount(expense);
          }
          return sum;
        }, 0);
  const otherOperatingExpenses = useAnnual
    ? Number(fullYearMetrics.other_operating_expenses) || 0
    : selectedMetrics
      ? Number(selectedMetrics.other_operating_expenses) || 0
      : sourceNoInvoice.reduce((sum, expense) => {
          if (
            expense.expense_type === "kilometraje" ||
            expense.expense_type === "otro" ||
            expense.expense_type === "alquiler_local" ||
            expense.expense_type === "alquiler_cabina"
          ) {
            return sum + getNoInvoiceDeductibleAmount(expense);
          }
          return sum;
        }, 0);
  const operatingExpenses =
    invoiceExpenses + payrollExpenses + amortizationExpenses + otherOperatingExpenses;
  const operatingResultEl = document.getElementById("pnlOperatingResult");
  const financialResultEl = document.getElementById("pnlFinancialResult");
  const preTaxEl = document.getElementById("pnlPreTax");
  const netEl = document.getElementById("pnlNet");
  if (!operatingResultEl || !financialResultEl || !preTaxEl || !netEl) {
    return;
  }

  setPnlInputValue("pnlLine1", incomeTotal, true);
  setPnlInputValue("pnlLine4", invoiceExpenses, true);
  setPnlInputValue("pnlLine6", payrollExpenses, true);
  setPnlInputValue("pnlLine7", otherOperatingExpenses, true);
  setPnlInputValue("pnlLine8", amortizationExpenses, true);
  setPnlInputValue("pnlLine14", loanInterest, true);

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
  const defaultTaxes = useAnnual
    ? companyType === "individual"
      ? Number(fullYearMetrics.model_130_estimate) || 0
      : Number(fullYearMetrics.corporate_tax_estimate) || 0
    : selectedMetrics
      ? companyType === "individual"
        ? Number(selectedMetrics.model_130_estimate) || 0
        : Number(selectedMetrics.corporate_tax_estimate) || 0
      : preTax > 0
        ? preTax * (companyType === "company" ? 0.25 : companyType === "individual" ? 0.15 : 0)
        : 0;
  setPnlInputValue("pnlLine19", defaultTaxes, true);
  const taxes = getPnlInputValue("pnlLine19");
  const netResult = preTax - taxes;

  operatingResultEl.textContent = formatCurrency(operatingResult);
  financialResultEl.textContent = formatCurrency(financialResult);
  preTaxEl.textContent = formatCurrency(preTax);
  netEl.textContent = formatCurrency(netResult);
  syncBalanceStatementFields(netResult);
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
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const margin = 48;
  const contentWidth = pageWidth - margin * 2;
  let y = margin;

  const brandColor = [34, 124, 101];
  const mutedColor = [93, 106, 99];
  const borderColor = [221, 229, 223];
  const rowAlt = [249, 251, 247];

  doc.setFillColor(242, 245, 242);
  doc.rect(0, 0, pageWidth, 84, "F");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.setTextColor(...brandColor);
  doc.text("Ledged", margin, 40);
  doc.setFontSize(18);
  doc.setTextColor(28, 32, 36);
  doc.text("Cuenta de pérdidas y ganancias (estimada)", margin, 64);

  y = 104;
  doc.setFont("helvetica", "normal");
  doc.setFontSize(11);
  doc.setTextColor(40, 44, 48);
  const dateLabel = new Date().toLocaleDateString("es-ES");
  const meta = [];
  if (nameValue) meta.push(`Nombre: ${nameValue}`);
  if (taxIdValue) meta.push(`CIF/NIF: ${taxIdValue}`);
  if (periodLabel) meta.push(`Periodo: ${periodLabel}`);
  meta.push(`Fecha de generación: ${dateLabel}`);

  meta.forEach((line) => {
    doc.text(line, margin, y);
    y += 16;
  });

  y += 8;
  doc.setDrawColor(...borderColor);
  doc.line(margin, y, pageWidth - margin, y);
  y += 18;

  doc.setFont("helvetica", "bold");
  doc.setTextColor(...mutedColor);
  doc.setFontSize(10);
  doc.text("CONCEPTO", margin, y);
  doc.text("IMPORTE (€)", pageWidth - margin, y, { align: "right" });
  y += 10;
  doc.setDrawColor(...borderColor);
  doc.line(margin, y, pageWidth - margin, y);
  y += 12;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10.5);
  doc.setTextColor(36, 40, 44);
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
  const amountColWidth = 110;
  const labelColWidth = contentWidth - amountColWidth - 12;

  const isSectionRow = (label) =>
    label.startsWith("A)") ||
    label.startsWith("B)") ||
    label.startsWith("C)") ||
    label.startsWith("D)");

  const ensureSpace = (height) => {
    if (y + height <= pageHeight - margin) {
      return;
    }
    doc.addPage();
    y = margin;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.setTextColor(...mutedColor);
    doc.text("CONCEPTO", margin, y);
    doc.text("IMPORTE (€)", pageWidth - margin, y, { align: "right" });
    y += 10;
    doc.setDrawColor(...borderColor);
    doc.line(margin, y, pageWidth - margin, y);
    y += 12;
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10.5);
    doc.setTextColor(36, 40, 44);
  };

  rows.forEach(([label, value], index) => {
    const section = isSectionRow(label);
    const labelLines = doc.splitTextToSize(label, labelColWidth);
    const rowHeight = Math.max(labelLines.length, 1) * 14 + 6;
    ensureSpace(rowHeight);

    if (index % 2 === 0) {
      doc.setFillColor(...rowAlt);
      doc.rect(margin, y - 10, contentWidth, rowHeight, "F");
    }

    if (section) {
      doc.setFont("helvetica", "bold");
      doc.setTextColor(...brandColor);
    } else {
      doc.setFont("helvetica", "normal");
      doc.setTextColor(36, 40, 44);
    }

    doc.text(labelLines, margin, y);
    doc.setFont(section ? "helvetica" : "helvetica", section ? "bold" : "normal");
    doc.setTextColor(section ? brandColor[0] : 36, section ? brandColor[1] : 40, section ? brandColor[2] : 44);
    doc.text(value, pageWidth - margin, y, { align: "right" });
    y += rowHeight;
  });

  y += 12;
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(...mutedColor);
  const footer =
    "Estructura basada en el Plan General Contable PYMES. Importes estimados automáticamente y no sustitutivos del asesoramiento fiscal profesional.";
  const footerLines = doc.splitTextToSize(footer, contentWidth);
  ensureSpace(footerLines.length * 12);
  doc.text(footerLines, margin, y);

  const filename = `pnl_${periodLabel || "periodo"}.pdf`.replace(/\s+/g, "_");
  doc.save(filename);
}

function getPnlRowsForEmail() {
  return [
    { label: "1. Importe neto de la cifra de negocios", value: getPnlInputValue("pnlLine1") },
    { label: "2. Variación de existencias de productos terminados y en curso", value: getPnlInputValue("pnlLine2") },
    { label: "3. Trabajos realizados por la empresa para su activo", value: getPnlInputValue("pnlLine3") },
    { label: "4. Aprovisionamientos", value: getPnlInputValue("pnlLine4") },
    { label: "5. Otros ingresos de explotación", value: getPnlInputValue("pnlLine5") },
    { label: "6. Gastos de personal", value: getPnlInputValue("pnlLine6") },
    { label: "7. Otros gastos de explotación", value: getPnlInputValue("pnlLine7") },
    { label: "8. Amortización del inmovilizado", value: getPnlInputValue("pnlLine8") },
    { label: "9. Imputación de subvenciones de inmovilizado no financiero y otras", value: getPnlInputValue("pnlLine9") },
    { label: "10. Excesos de provisiones", value: getPnlInputValue("pnlLine10") },
    { label: "11. Deterioro y resultado por enajenación del inmovilizado", value: getPnlInputValue("pnlLine11") },
    { label: "12. Otros resultados", value: getPnlInputValue("pnlLine12") },
    { label: "A) RESULTADO DE EXPLOTACIÓN", value: document.getElementById("pnlOperatingResult")?.textContent || "" },
    { label: "13.a Imputación de subvenciones, donaciones y legados de carácter financiero", value: getPnlInputValue("pnlLine13a") },
    { label: "13.b Otros ingresos financieros", value: getPnlInputValue("pnlLine13b") },
    { label: "14. Gastos financieros", value: getPnlInputValue("pnlLine14") },
    { label: "15. Variación de valor razonable en instrumentos financieros", value: getPnlInputValue("pnlLine15") },
    { label: "16. Diferencias de cambio", value: getPnlInputValue("pnlLine16") },
    { label: "17. Deterioro y resultado por enajenación de instrumentos financieros", value: getPnlInputValue("pnlLine17") },
    { label: "18.a Incorporación al activo de gastos financieros", value: getPnlInputValue("pnlLine18a") },
    { label: "18.b Ingresos financieros derivados de convenios de acreedores", value: getPnlInputValue("pnlLine18b") },
    { label: "18.c Resto de ingresos y gastos", value: getPnlInputValue("pnlLine18c") },
    { label: "B) RESULTADO FINANCIERO", value: document.getElementById("pnlFinancialResult")?.textContent || "" },
    { label: "C) RESULTADO ANTES DE IMPUESTOS", value: document.getElementById("pnlPreTax")?.textContent || "" },
    { label: "19. Impuestos sobre beneficios", value: getPnlInputValue("pnlLine19") },
    { label: "D) RESULTADO DEL EJERCICIO", value: document.getElementById("pnlNet")?.textContent || "" },
  ];
}

function exportBalancePdf() {
  if (!window.jspdf || !balanceName || !balanceTaxId) {
    return;
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "a4" });

  const nameValue = balanceName.value.trim();
  const taxIdValue = balanceTaxId.value.trim();
  const periodLabel = headerPeriodLabel ? headerPeriodLabel.textContent.trim() : "";
  const generated = new Date().toLocaleDateString("es-ES");

  const pageWidth = doc.internal.pageSize.getWidth();
  const margin = 48;
  const contentWidth = pageWidth - margin * 2;
  let y = margin;

  const brandColor = [34, 124, 101];
  const mutedColor = [93, 106, 99];

  doc.setFillColor(242, 245, 242);
  doc.rect(0, 0, pageWidth, 84, "F");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.setTextColor(...brandColor);
  doc.text("Ledged", margin, 40);
  doc.setFontSize(18);
  doc.setTextColor(28, 32, 36);
  doc.text("Balance de situación (estimado)", margin, 64);

  y = 104;
  doc.setFont("helvetica", "normal");
  doc.setFontSize(11);
  doc.setTextColor(40, 44, 48);
  const meta = [];
  if (nameValue) meta.push(`Nombre: ${nameValue}`);
  if (taxIdValue) meta.push(`CIF/NIF: ${taxIdValue}`);
  if (periodLabel) meta.push(`Periodo: ${periodLabel}`);
  meta.push(`Fecha de generación: ${generated}`);

  meta.forEach((line) => {
    doc.text(line, margin, y);
    y += 16;
  });

  y += 8;
  doc.setDrawColor(221, 229, 223);
  doc.line(margin, y, pageWidth - margin, y);
  y += 18;

  const assetRows = [
    ["A) ACTIVO NO CORRIENTE", ""],
    ["I. Inmovilizado intangible", formatCurrency(getBalanceInputValue("bsAssetIntangible"))],
    ["II. Inmovilizado material", formatCurrency(getBalanceInputValue("bsAssetTangible"))],
    ["III. Inversiones inmobiliarias", formatCurrency(getBalanceInputValue("bsAssetInvestmentProperty"))],
    ["IV. Inversiones financieras a largo plazo", formatCurrency(getBalanceInputValue("bsAssetLongTermInv"))],
    ["V. Activos por impuesto diferido", formatCurrency(getBalanceInputValue("bsAssetDeferredTax"))],
    ["Subtotal activo no corriente", bsAssetNonCurrentTotal ? bsAssetNonCurrentTotal.textContent : formatCurrency(0)],
    ["B) ACTIVO CORRIENTE", ""],
    ["I. Existencias", formatCurrency(getBalanceInputValue("bsAssetInventory"))],
    ["II. Deudores comerciales y otras cuentas a cobrar", formatCurrency(getBalanceInputValue("bsAssetReceivables"))],
    ["III. Inversiones financieras a corto plazo", formatCurrency(getBalanceInputValue("bsAssetShortTermInv"))],
    ["IV. Efectivo y otros activos líquidos equivalentes", formatCurrency(getBalanceInputValue("bsAssetCash"))],
    ["V. Activos por impuesto corriente", formatCurrency(getBalanceInputValue("bsAssetCurrentTax"))],
    ["Subtotal activo corriente", bsAssetCurrentTotal ? bsAssetCurrentTotal.textContent : formatCurrency(0)],
    ["TOTAL ACTIVO", bsTotalAssets ? bsTotalAssets.textContent : formatCurrency(0)],
  ];
  const liabilityRows = [
    ["A) PATRIMONIO NETO", ""],
    ["I. Capital", formatCurrency(getBalanceInputValue("bsEquityCapital"))],
    ["II. Reservas", formatCurrency(getBalanceInputValue("bsEquityReserves"))],
    ["III. Resultado del ejercicio", formatCurrency(getBalanceInputValue("bsEquityResult"))],
    ["IV. Subvenciones, donaciones y legados", formatCurrency(getBalanceInputValue("bsEquityGrants"))],
    ["Subtotal patrimonio neto", bsEquityTotal ? bsEquityTotal.textContent : formatCurrency(0)],
    ["B) PASIVO NO CORRIENTE", ""],
    ["I. Deudas a largo plazo", formatCurrency(getBalanceInputValue("bsLiabLongTermDebt"))],
    ["II. Otras obligaciones a largo plazo", formatCurrency(getBalanceInputValue("bsLiabLongTermOther"))],
    ["Subtotal pasivo no corriente", bsLiabNonCurrentTotal ? bsLiabNonCurrentTotal.textContent : formatCurrency(0)],
    ["C) PASIVO CORRIENTE", ""],
    ["I. Deudas a corto plazo", formatCurrency(getBalanceInputValue("bsLiabShortTermDebt"))],
    ["II. Proveedores y otras cuentas a pagar", formatCurrency(getBalanceInputValue("bsLiabPayables"))],
    ["III. Otras obligaciones corrientes", formatCurrency(getBalanceInputValue("bsLiabOtherCurrent"))],
    ["Subtotal pasivo corriente", bsLiabCurrentTotal ? bsLiabCurrentTotal.textContent : formatCurrency(0)],
    ["TOTAL PATRIMONIO NETO Y PASIVO", bsTotalLiabilities ? bsTotalLiabilities.textContent : formatCurrency(0)],
  ];

  doc.setFont("helvetica", "bold");
  doc.setTextColor(...mutedColor);
  doc.text("Activo", margin, y);
  doc.text("Patrimonio neto y pasivo", margin + contentWidth / 2 + 10, y);
  y += 18;

  doc.setFont("helvetica", "normal");
  doc.setTextColor(36, 40, 44);
  let yLeft = y;
  assetRows.forEach((row) => {
    const isHeader = row[1] === "" || row[0].startsWith("TOTAL");
    doc.setFont("helvetica", isHeader ? "bold" : "normal");
    doc.text(String(row[0]), 40, yLeft);
    if (row[1]) {
      doc.text(String(row[1]), 220, yLeft, { align: "right" });
    }
    yLeft += 18;
  });

  let yRight = y;
  liabilityRows.forEach((row) => {
    const isHeader = row[1] === "" || row[0].startsWith("TOTAL");
    doc.setFont("helvetica", isHeader ? "bold" : "normal");
    doc.text(String(row[0]), margin + contentWidth / 2 + 10, yRight);
    if (row[1]) {
      doc.text(String(row[1]), pageWidth - margin, yRight, { align: "right" });
    }
    yRight += 18;
  });

  doc.setFontSize(9);
  doc.setTextColor(...mutedColor);
  doc.text(
    "Balance de situación estimado. Importes editables y no sustitutivos del asesoramiento fiscal profesional.",
    margin,
    760,
    { maxWidth: contentWidth }
  );

  const filename = `balance_${periodLabel || "periodo"}.pdf`.replace(/\s+/g, "_");
  doc.save(filename);
}

function sendPnlEmail() {
  const periodLabel = headerPeriodLabel ? headerPeriodLabel.textContent.trim() : "";
  const payload = {
    name: pnlName ? pnlName.value.trim() : "",
    tax_id: pnlTaxId ? pnlTaxId.value.trim() : "",
    period_label: periodLabel,
    lines: getPnlRowsForEmail(),
    totals: {
      operating: document.getElementById("pnlOperatingResult")?.textContent || "",
      financial: document.getElementById("pnlFinancialResult")?.textContent || "",
      pretax: document.getElementById("pnlPreTax")?.textContent || "",
      net: document.getElementById("pnlNet")?.textContent || "",
    },
  };

  fetch(withCompanyParam("/api/pnl/email"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["No se pudo enviar el P&G."]).join("\n"));
        return;
      }
      alert("P&G enviado correctamente.");
    })
    .catch(() => {
      alert("No se pudo enviar el P&G.");
    });
}

function sendBalanceEmail() {
  const periodLabel = headerPeriodLabel ? headerPeriodLabel.textContent.trim() : "";
  const payload = {
    name: balanceName ? balanceName.value.trim() : "",
    tax_id: balanceTaxId ? balanceTaxId.value.trim() : "",
    period_label: periodLabel,
    lines: getBalanceRowsForEmail(),
  };

  fetch(withCompanyParam("/api/balance/email"), {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["No se pudo enviar el balance."]).join("\n"));
        return;
      }
      alert("Balance enviado correctamente.");
    })
    .catch(() => {
      alert("No se pudo enviar el balance.");
    });
}

function getBalanceRowsForEmail() {
  return [
    { label: "A) ACTIVO NO CORRIENTE", value: "" },
    { label: "I. Inmovilizado intangible", value: getBalanceInputValue("bsAssetIntangible") },
    { label: "II. Inmovilizado material", value: getBalanceInputValue("bsAssetTangible") },
    { label: "III. Inversiones inmobiliarias", value: getBalanceInputValue("bsAssetInvestmentProperty") },
    { label: "IV. Inversiones financieras a largo plazo", value: getBalanceInputValue("bsAssetLongTermInv") },
    { label: "V. Activos por impuesto diferido", value: getBalanceInputValue("bsAssetDeferredTax") },
    { label: "Subtotal activo no corriente", value: bsAssetNonCurrentTotal ? bsAssetNonCurrentTotal.textContent : "" },
    { label: "B) ACTIVO CORRIENTE", value: "" },
    { label: "I. Existencias", value: getBalanceInputValue("bsAssetInventory") },
    { label: "II. Deudores comerciales y otras cuentas a cobrar", value: getBalanceInputValue("bsAssetReceivables") },
    { label: "III. Inversiones financieras a corto plazo", value: getBalanceInputValue("bsAssetShortTermInv") },
    { label: "IV. Efectivo y otros activos líquidos equivalentes", value: getBalanceInputValue("bsAssetCash") },
    { label: "V. Activos por impuesto corriente", value: getBalanceInputValue("bsAssetCurrentTax") },
    { label: "Subtotal activo corriente", value: bsAssetCurrentTotal ? bsAssetCurrentTotal.textContent : "" },
    { label: "TOTAL ACTIVO", value: bsTotalAssets ? bsTotalAssets.textContent : "" },
    { label: "A) PATRIMONIO NETO", value: "" },
    { label: "I. Capital", value: getBalanceInputValue("bsEquityCapital") },
    { label: "II. Reservas", value: getBalanceInputValue("bsEquityReserves") },
    { label: "III. Resultado del ejercicio", value: getBalanceInputValue("bsEquityResult") },
    { label: "IV. Subvenciones, donaciones y legados", value: getBalanceInputValue("bsEquityGrants") },
    { label: "Subtotal patrimonio neto", value: bsEquityTotal ? bsEquityTotal.textContent : "" },
    { label: "B) PASIVO NO CORRIENTE", value: "" },
    { label: "I. Deudas a largo plazo", value: getBalanceInputValue("bsLiabLongTermDebt") },
    { label: "II. Otras obligaciones a largo plazo", value: getBalanceInputValue("bsLiabLongTermOther") },
    { label: "Subtotal pasivo no corriente", value: bsLiabNonCurrentTotal ? bsLiabNonCurrentTotal.textContent : "" },
    { label: "C) PASIVO CORRIENTE", value: "" },
    { label: "I. Deudas a corto plazo", value: getBalanceInputValue("bsLiabShortTermDebt") },
    { label: "II. Proveedores y otras cuentas a pagar", value: getBalanceInputValue("bsLiabPayables") },
    { label: "III. Otras obligaciones corrientes", value: getBalanceInputValue("bsLiabOtherCurrent") },
    { label: "Subtotal pasivo corriente", value: bsLiabCurrentTotal ? bsLiabCurrentTotal.textContent : "" },
    { label: "TOTAL PATRIMONIO NETO Y PASIVO", value: bsTotalLiabilities ? bsTotalLiabilities.textContent : "" },
  ];
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

  updateDashboardTotals();
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

function syncReportSelectorsFromMain() {
  if (reportYearSelect && yearSelect) {
    reportYearSelect.value = yearSelect.value;
  }
  if (!reportQuarterSelect || !reportStartMonthSelect || !reportEndMonthSelect || !monthSelect || !periodSelect) {
    toggleReportCustomRange();
    return;
  }

  const selectedMonth = Number(monthSelect.value || 1);
  if (getSelectedPeriod() === "quarterly") {
    const quarter = Math.min(4, Math.max(1, Math.ceil(selectedMonth / 3)));
    reportQuarterSelect.value = String(quarter);
    reportStartMonthSelect.value = String((quarter - 1) * 3 + 1);
    reportEndMonthSelect.value = String(quarter * 3);
  } else {
    reportQuarterSelect.value = "custom";
    reportStartMonthSelect.value = String(selectedMonth);
    reportEndMonthSelect.value = String(selectedMonth);
  }
  toggleReportCustomRange();
}

function applyReportSelectorsToMain() {
  if (!monthSelect || !yearSelect || !periodSelect) {
    return;
  }

  if (reportYearSelect?.value) {
    yearSelect.value = reportYearSelect.value;
  }

  const reportMode = reportQuarterSelect?.value || "custom";
  if (reportMode === "custom") {
    const startMonth = Number(reportStartMonthSelect?.value || monthSelect.value || 1);
    const endMonth = Number(reportEndMonthSelect?.value || monthSelect.value || startMonth);
    const isQuarterRange =
      (startMonth === 1 && endMonth === 3) ||
      (startMonth === 4 && endMonth === 6) ||
      (startMonth === 7 && endMonth === 9) ||
      (startMonth === 10 && endMonth === 12);

    periodSelect.value = isQuarterRange ? "quarterly" : "monthly";
    monthSelect.value = String(endMonth);
  } else {
    const quarter = Number(reportMode);
    periodSelect.value = "quarterly";
    if (quarter >= 1 && quarter <= 4) {
      monthSelect.value = String(quarter * 3);
      if (reportStartMonthSelect) {
        reportStartMonthSelect.value = String((quarter - 1) * 3 + 1);
      }
      if (reportEndMonthSelect) {
        reportEndMonthSelect.value = String(quarter * 3);
      }
    }
  }

  document.body.classList.toggle(
    "period-quarterly",
    getSelectedPeriod() === "quarterly"
  );
  persistFilters();
  updateHeaderContext();
  calendarOverride = false;
  syncCalendarWithFilters();
  refreshAllData();
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

function padMonthDay(value) {
  return String(value).padStart(2, "0");
}

function buildIntegrationRangeFromMainFilters() {
  const { month, year } = getSelectedMonthYear();
  const normalizedMonth = Number(month || 1);
  const normalizedYear = Number(year || new Date().getFullYear());
  let startMonth = normalizedMonth;
  let endMonth = normalizedMonth;
  if (getSelectedPeriod() === "quarterly") {
    startMonth = Math.max(1, normalizedMonth - 2);
    endMonth = normalizedMonth;
  }
  const startDate = `${normalizedYear}-${padMonthDay(startMonth)}-01`;
  const endDate = `${normalizedYear}-${padMonthDay(endMonth)}-${padMonthDay(
    new Date(normalizedYear, endMonth, 0).getDate()
  )}`;
  return { startDate, endDate };
}

function syncIntegrationRangeFromMainFilters() {
  if (!integrationStartDate || !integrationEndDate) {
    return;
  }
  const { startDate, endDate } = buildIntegrationRangeFromMainFilters();
  integrationStartDate.value = startDate;
  integrationEndDate.value = endDate;
}

function buildIntegrationParams() {
  const params = new URLSearchParams();
  if (integrationStartDate?.value) {
    params.set("start_date", integrationStartDate.value);
  }
  if (integrationEndDate?.value) {
    params.set("end_date", integrationEndDate.value);
  }
  if (integrationFormat?.value) {
    params.set("format", integrationFormat.value);
  }
  return params;
}

function renderAccountingIntegrationSummary() {
  const stats = currentAccountingIntegrationSummary?.stats || {};
  if (integrationPurchasesCount) {
    integrationPurchasesCount.textContent = String(stats.purchaseDocuments || 0);
  }
  if (integrationSalesCount) {
    integrationSalesCount.textContent = String(stats.salesDocuments || 0);
  }
  if (integrationJournalCount) {
    integrationJournalCount.textContent = String(stats.journalEntries || 0);
  }
  if (integrationDocumentsCount) {
    integrationDocumentsCount.textContent = String(stats.documentFiles || 0);
  }
  if (integrationStatus) {
    if (!currentAccountingIntegrationSummary) {
      integrationStatus.textContent = "";
      return;
    }
    const companyName = currentAccountingIntegrationSummary.company?.displayName || "Empresa seleccionada";
    const periodLabel = currentAccountingIntegrationSummary.periodLabel || "";
    integrationStatus.textContent = `${companyName} · ${periodLabel} · ${stats.journalLines || 0} líneas contables preparadas.`;
  }
}

function loadAccountingIntegrationSummary() {
  if (!getSelectedCompanyId()) {
    currentAccountingIntegrationSummary = null;
    renderAccountingIntegrationSummary();
    return Promise.resolve();
  }
  const params = buildIntegrationParams();
  const query = params.toString();
  const url = withCompanyParam(`/api/accounting-integrations/summary${query ? `?${query}` : ""}`);
  return fetch(url)
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        throw new Error((data.errors || ["No se pudo cargar el resumen de integraciones."]).join("\n"));
      }
      currentAccountingIntegrationSummary = data;
      renderAccountingIntegrationSummary();
    })
    .catch((error) => {
      currentAccountingIntegrationSummary = null;
      renderAccountingIntegrationSummary();
      if (integrationStatus) {
        integrationStatus.textContent = error.message || "No se pudo cargar el resumen de integraciones.";
      }
    });
}

function triggerAccountingExport(kind) {
  if (!getSelectedCompanyId()) {
    alert("Selecciona una empresa antes de exportar.");
    return;
  }
  const params = buildIntegrationParams();
  params.delete("format");
  if (kind !== "package" && integrationFormat?.value) {
    params.set("format", integrationFormat.value);
  }
  const query = params.toString();
  const baseUrl =
    kind === "package"
      ? `/api/accounting-integrations/export-package${query ? `?${query}` : ""}`
      : `/api/accounting-integrations/export/${kind}${query ? `?${query}` : ""}`;
  if (integrationStatus) {
    integrationStatus.textContent =
      kind === "package" ? "Preparando paquete contable..." : "Preparando exportación...";
  }
  window.location.href = withCompanyParam(baseUrl);
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
  ensureCompositeSectionState(sectionId);
  document.body.classList.remove("sidebar-open");
}

function initNavigation() {
  const legacySectionMap = {
    invoices: "expenses",
    payments: "dashboard",
    taxes: "reports",
    pnl: "statements",
    balance: "statements",
    "document-center": "expenses",
  };
  const querySection = legacySectionMap[new URLSearchParams(window.location.search).get("section")] ||
    new URLSearchParams(window.location.search).get("section");
  const storedSection = legacySectionMap[localStorage.getItem("activeSection")] || localStorage.getItem("activeSection");
  const availableSections = Array.from(sections || []).map(
    (section) => section.dataset.section
  );
  const defaultSection = availableSections.includes(querySection)
    ? querySection
    : availableSections.includes(storedSection)
    ? storedSection
    : availableSections[0] || "dashboard";
  setActiveSection(defaultSection);

  navLinks.forEach((link) => {
    link.addEventListener("click", () => {
      setActiveSection(link.dataset.section);
    });
  });

  sectionTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      const parent = tab.dataset.parent;
      const subsection = tab.dataset.subsection;
      if (parent === "expenses") {
        setExpenseSubview(subsection);
        return;
      }
      setGroupedSubview(parent, subsection);
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
  [
    [invoiceSearchInput, () => renderInvoices(currentInvoices || [])],
    [noInvoiceSearchInput, () => renderNoInvoiceExpenses(currentNoInvoiceExpenses || [])],
    [loanSearchInput, () => renderLoanInstallments(currentLoanInstallments || [])],
    [incomeSearchInput, () => renderIncomeInvoices(currentIncomeInvoices || [])],
    [billingSearchInput, () => renderBillingEntries(currentBillingEntries || [])],
    [archiveSearchInput, renderArchive],
  ].forEach(([input, handler]) => {
    if (input) {
      input.addEventListener("input", handler);
    }
  });
  [
    [archiveTypeFilter, renderArchive],
    [archiveGroupBy, renderArchive],
    [archiveMonthFilter, renderArchive],
    [archiveStatusFilter, renderArchive],
  ].forEach(([input, handler]) => {
    if (input) {
      input.addEventListener("change", handler);
    }
  });
  if (archiveExportBtn) {
    archiveExportBtn.addEventListener("click", exportArchiveCsv);
  }
  const selectFilesBtn = document.getElementById("selectFiles");
  const selectFolderBtn = document.getElementById("selectFolder");
  const incomeSelectFilesBtn = document.getElementById("incomeSelectFiles");
  const incomeSelectFolderBtn = document.getElementById("incomeSelectFolder");
  const lowQualityAccept = document.getElementById("lowQualityAccept");
  const lowQualityClose = document.getElementById("lowQualityClose");
  if (lowQualityAccept) {
    lowQualityAccept.addEventListener("click", hideLowQualityModal);
  }
  if (lowQualityClose) {
    lowQualityClose.addEventListener("click", hideLowQualityModal);
  }

  const vatWarningAccept = document.getElementById("vatWarningAccept");
  const vatWarningClose = document.getElementById("vatWarningClose");
  if (vatWarningAccept) {
    vatWarningAccept.addEventListener("click", hideVatWarningModal);
  }
  if (vatWarningClose) {
    vatWarningClose.addEventListener("click", hideVatWarningModal);
  }
  if (paymentPrevMonth) {
    paymentPrevMonth.addEventListener("click", () => shiftCalendarMonth(-1));
  }
  if (paymentNextMonth) {
    paymentNextMonth.addEventListener("click", () => shiftCalendarMonth(1));
  }
  if (paymentClearAllBtn) {
    paymentClearAllBtn.addEventListener("click", () => {
      markPaymentsAsPaid(getAlertPendingPayments(currentPayments));
    });
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
  if (incomeFolderInput && incomeSelectFolderBtn) {
    incomeSelectFolderBtn.addEventListener("click", () => {
      incomeFolderInput.click();
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
  if (incomeFolderInput) {
    incomeFolderInput.addEventListener("change", (event) => {
      addIncomeFiles(event.target.files);
      incomeFolderInput.value = "";
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

  if (loanPlanSelectBtn && loanPlanFile) {
    loanPlanSelectBtn.addEventListener("click", () => {
      loanPlanFile.click();
    });
  }

  if (loanPlanFile) {
    loanPlanFile.addEventListener("change", () => {
      const file = loanPlanFile.files && loanPlanFile.files[0];
      selectedLoanPlanFile = file || null;
      if (loanPlanFileName) {
        loanPlanFileName.textContent = file ? file.name : "Ningún archivo seleccionado.";
      }
      resetLoanPlanPreview();
    });
  }

  if (loanPlanDropZone) {
    loanPlanDropZone.addEventListener("dragover", (event) => {
      event.preventDefault();
      loanPlanDropZone.classList.add("dragover");
    });
    loanPlanDropZone.addEventListener("dragleave", () => {
      loanPlanDropZone.classList.remove("dragover");
    });
    loanPlanDropZone.addEventListener("drop", (event) => {
      event.preventDefault();
      loanPlanDropZone.classList.remove("dragover");
      const file = event.dataTransfer?.files?.[0];
      if (!file) {
        return;
      }
      selectedLoanPlanFile = file;
      if (loanPlanFile) {
        const transfer = new DataTransfer();
        transfer.items.add(file);
        loanPlanFile.files = transfer.files;
      }
      if (loanPlanFileName) {
        loanPlanFileName.textContent = file.name;
      }
      resetLoanPlanPreview();
    });
  }

  if (loanPlanSaveBtn) {
    loanPlanSaveBtn.addEventListener("click", saveLoanPlanDraft);
  }

  if (uploadBtn) {
    uploadBtn.addEventListener("click", uploadPending);
  }
  if (incomeUploadBtn) {
    incomeUploadBtn.addEventListener("click", uploadIncomePending);
  }
  attachAmountInputBehavior(billingBaseInput);
  attachAmountInputBehavior(billingVatAmountInput);
  attachAmountInputBehavior(billingTotalInput);
  attachAmountInputBehavior(noInvoiceAmount);
  attachAmountInputBehavior(noInvoiceInterest);
  attachAmountInputBehavior(noInvoiceWithholding);
  attachAmountInputBehavior(noInvoiceVatBase);
  attachAmountInputBehavior(noInvoiceVatAmount);
  attachAmountInputBehavior(loanTotalInput);
  attachAmountInputBehavior(loanInterestInput);
  attachAmountInputBehavior(loanPrincipalInput);
  attachAmountInputBehavior(documentCenterBaseAmount);
  attachAmountInputBehavior(documentCenterVatAmount);
  attachAmountInputBehavior(documentCenterTotalAmount);
  attachAmountInputBehavior(documentCenterPayrollGross);
  attachAmountInputBehavior(documentCenterPayrollNet);
  attachAmountInputBehavior(documentCenterPayrollDeductions);
  attachAmountInputBehavior(documentCenterPayrollEmployerCost);
  attachAmountInputBehavior(documentCenterTaxModelAmount);
  attachAmountInputBehavior(documentCenterTaxModelOffsetAmount);
  attachAmountInputBehavior(documentCenterTaxModelRefundAmount);
  if (companySaveBtn) {
    companySaveBtn.addEventListener("click", saveCompany);
  }
  if (companyType) {
    companyType.addEventListener("change", syncCompanyFiscalProfileInputs);
    syncCompanyFiscalProfileInputs();
  }
  if (accountSaveBtn) {
    accountSaveBtn.addEventListener("click", saveAccount);
  }
  if (documentCenterUploadBtn) {
    documentCenterUploadBtn.addEventListener("click", uploadDocumentCenterBatch);
  }
  if (documentCenterDocType) {
    documentCenterDocType.addEventListener("change", () => {
      toggleDocumentCenterSpecialFields(documentCenterDocType.value || "unknown");
    });
  }
  if (documentCenterBatchFilter) {
    documentCenterBatchFilter.addEventListener("change", () => {
      selectedDocumentBatchId = documentCenterBatchFilter.value || "";
      loadDocumentCenterDocuments();
      renderDocumentCenterBatches();
    });
  }
  if (documentCenterStatusFilter) {
    documentCenterStatusFilter.addEventListener("change", () => {
      loadDocumentCenterDocuments();
    });
  }
  if (documentCenterRegisterReadyBtn) {
    documentCenterRegisterReadyBtn.addEventListener("click", registerReadyDocumentCenterBatch);
  }
  if (documentCenterSaveBtn) {
    documentCenterSaveBtn.addEventListener("click", saveDocumentCenterDocument);
  }
  if (documentCenterApproveBtn) {
    documentCenterApproveBtn.addEventListener("click", () => {
      saveDocumentCenterDocument().then(() => changeDocumentCenterState("approve"));
    });
  }
  if (documentCenterRegisterBtn) {
    documentCenterRegisterBtn.addEventListener("click", () => {
      saveDocumentCenterDocument().then(() => changeDocumentCenterState("register"));
    });
  }
  if (documentCenterRejectBtn) {
    documentCenterRejectBtn.addEventListener("click", () => {
      const reason = window.prompt("Motivo del rechazo:", "Documento rechazado manualmente.");
      if (reason === null) {
        return;
      }
      changeDocumentCenterState("reject", { reason });
    });
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
      syncReportSelectorsFromMain();
      syncIntegrationRangeFromMainFilters();
      refreshAllData();
    });
  }
  if (yearSelect) {
    yearSelect.addEventListener("change", () => {
      persistFilters();
      updateHeaderContext();
      calendarOverride = false;
      syncCalendarWithFilters();
      syncReportSelectorsFromMain();
      syncIntegrationRangeFromMainFilters();
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
      syncReportSelectorsFromMain();
      syncIntegrationRangeFromMainFilters();
      refreshAllData();
    });
  }
  if (companySelect) {
    companySelect.addEventListener("change", () => {
      selectedCompanyId = companySelect.value;
      loadBalanceManualDataFromCompany();
      persistFilters();
      applyCompanyTaxModules();
      updatePnlSummary();
      updateHeaderContext();
      renderFiscalModelsSummary();
      calendarOverride = false;
      syncCalendarWithFilters();
      syncIntegrationRangeFromMainFilters();
      renderTable();
      renderIncomeTable();
      loadYears().then(() => {
        syncReportSelectorsFromMain();
        refreshAllData();
      });
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
  if (noInvoiceType && noInvoiceDeductible && noInvoiceInterestField && noInvoiceInterest) {
    noInvoiceType.addEventListener("change", () => {
      toggleLoanInterestField({
        typeValue: noInvoiceType.value,
        interestField: noInvoiceInterestField,
        interestInput: noInvoiceInterest,
        deductibleSelect: noInvoiceDeductible,
      });
      toggleNoInvoiceWithholdingField({
        typeValue: noInvoiceType.value,
      });
      toggleNoInvoiceVatFields({
        typeValue: noInvoiceType.value,
        vatDeductibleValue: noInvoiceVatDeductible?.value,
      });
    });
    toggleLoanInterestField({
      typeValue: noInvoiceType.value,
      interestField: noInvoiceInterestField,
      interestInput: noInvoiceInterest,
      deductibleSelect: noInvoiceDeductible,
    });
    toggleNoInvoiceWithholdingField({
      typeValue: noInvoiceType.value,
    });
    toggleNoInvoiceVatFields({
      typeValue: noInvoiceType.value,
      vatDeductibleValue: noInvoiceVatDeductible?.value,
    });
  }
  if (noInvoiceVatDeductible) {
    noInvoiceVatDeductible.addEventListener("change", () => {
      toggleNoInvoiceVatFields({
        typeValue: noInvoiceType?.value,
        vatDeductibleValue: noInvoiceVatDeductible.value,
      });
      syncNoInvoiceVatCalculation();
    });
  }
  if (noInvoiceVatSelect) {
    noInvoiceVatSelect.addEventListener("change", syncNoInvoiceVatCalculation);
  }
  if (noInvoiceAmount) {
    noInvoiceAmount.addEventListener("input", syncNoInvoiceVatCalculation);
  }
  if (noInvoiceSaveBtn) {
    noInvoiceSaveBtn.addEventListener("click", saveNoInvoiceExpense);
  }
  if (loanTotalInput) {
    loanTotalInput.addEventListener("input", syncLoanPrincipal);
  }
  if (loanInterestInput) {
    loanInterestInput.addEventListener("input", syncLoanPrincipal);
  }
  if (loanSaveBtn) {
    loanSaveBtn.addEventListener("click", saveLoanInstallment);
  }
  if (loanImportBtn) {
    loanImportBtn.addEventListener("click", importLoanPlan);
  }
  if (exportPnlBtn) {
    exportPnlBtn.addEventListener("click", exportPnlPdf);
  }
  if (pnlEmailBtn) {
    pnlEmailBtn.addEventListener("click", sendPnlEmail);
  }
  if (balancePdfBtn) {
    balancePdfBtn.addEventListener("click", exportBalancePdf);
  }
  if (balanceEmailBtn) {
    balanceEmailBtn.addEventListener("click", sendBalanceEmail);
  }
  if (reportQuarterSelect) {
    reportQuarterSelect.addEventListener("change", () => {
      toggleReportCustomRange();
      applyReportSelectorsToMain();
    });
  }
  if (reportYearSelect) {
    reportYearSelect.addEventListener("change", applyReportSelectorsToMain);
  }
  if (reportStartMonthSelect) {
    reportStartMonthSelect.addEventListener("change", () => {
      if (reportQuarterSelect?.value === "custom") {
        applyReportSelectorsToMain();
      }
    });
  }
  if (reportEndMonthSelect) {
    reportEndMonthSelect.addEventListener("change", () => {
      if (reportQuarterSelect?.value === "custom") {
        applyReportSelectorsToMain();
      }
    });
  }
  if (reportDownloadBtn) {
    reportDownloadBtn.addEventListener("click", downloadQuarterlyReport);
  }
  if (reportEmailBtn) {
    reportEmailBtn.addEventListener("click", sendQuarterlyReportEmail);
  }
  if (integrationStartDate) {
    integrationStartDate.addEventListener("change", loadAccountingIntegrationSummary);
  }
  if (integrationEndDate) {
    integrationEndDate.addEventListener("change", loadAccountingIntegrationSummary);
  }
  if (integrationFormat) {
    integrationFormat.addEventListener("change", loadAccountingIntegrationSummary);
  }
  if (integrationExportPurchasesBtn) {
    integrationExportPurchasesBtn.addEventListener("click", () => triggerAccountingExport("purchases"));
  }
  if (integrationExportSalesBtn) {
    integrationExportSalesBtn.addEventListener("click", () => triggerAccountingExport("sales"));
  }
  if (integrationExportJournalBtn) {
    integrationExportJournalBtn.addEventListener("click", () => triggerAccountingExport("journal"));
  }
  if (integrationExportPackageBtn) {
    integrationExportPackageBtn.addEventListener("click", () => triggerAccountingExport("package"));
  }
  bindPnlInputs();
}

function init() {
  const now = new Date();
  populateMonthSelects();
  populateReportMonthSelect(reportStartMonthSelect);
  populateReportMonthSelect(reportEndMonthSelect);

  bindBalanceInputs();
  toggleReportCustomRange();
  bindEvents();
  initNavigation();
  applyRoleVisibility();
  restoreFilters(now);
  syncCalendarWithFilters();
  syncIntegrationRangeFromMainFilters();
  loadStaff()
    .then(() => loadCompanies())
    .then(() => loadYears())
    .then(() => {
      document.body.classList.toggle("period-quarterly", getSelectedPeriod() === "quarterly");
      billingMonthSelect.value = monthSelect.value;
      billingYearSelect.value = yearSelect.value;
      syncReportSelectorsFromMain();
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
      if (loanPaymentDateInput && !loanPaymentDateInput.value) {
        loanPaymentDateInput.value = now.toISOString().slice(0, 10);
      }
      if (documentCenterPeriod && !documentCenterPeriod.value) {
        documentCenterPeriod.value = getPeriodLabel() || `${now.getFullYear()}`;
      }
      updateHeaderContext();
      refreshAllData();
    });
}

document.addEventListener("DOMContentLoaded", init);
