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
let currentDeductibleExpenses = 0;
let annualBillingBaseTotal = 0;
let annualDeductibleExpenses = 0;
let currentPayments = null;
let selectedPaymentDay = null;

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

function formatCurrency(value) {
  const number = Number(value || 0);
  return `${number.toFixed(2).replace(".", ",")} €`;
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

const allowedExtensions = new Set([".pdf", ".jpg", ".jpeg", ".png"]);

const monthSelect = document.getElementById("monthSelect");
const yearSelect = document.getElementById("yearSelect");
const periodSelect = document.getElementById("periodSelect");
const billingMonthSelect = document.getElementById("billingMonthSelect");
const billingYearSelect = document.getElementById("billingYearSelect");
const billingBaseInput = document.getElementById("billingBaseInput");
const billingVatSelect = document.getElementById("billingVatSelect");
const billingSaveBtn = document.getElementById("billingSaveBtn");
const billingEntriesBody = document.querySelector("#billingEntriesTable tbody");
const billingEntriesEmpty = document.getElementById("billingEntriesEmpty");
const invoicesTableBody = document.querySelector("#invoicesTable tbody");
const invoicesEmpty = document.getElementById("invoicesEmpty");
const taxpayerSelect = document.getElementById("taxpayerSelect");
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
const paymentCalendar = document.getElementById("paymentCalendar");
const paymentCalendarTitle = document.getElementById("paymentCalendarTitle");
const paymentDayTitle = document.getElementById("paymentDayTitle");
const paymentDayList = document.getElementById("paymentDayList");
const paymentDayTotal = document.getElementById("paymentDayTotal");

function isAllowedFile(fileName) {
  const lower = fileName.toLowerCase();
  const dotIndex = lower.lastIndexOf(".");
  if (dotIndex === -1) {
    return false;
  }
  return allowedExtensions.has(lower.slice(dotIndex));
}

function populateMonthSelects() {
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
  return fetch("/api/years")
    .then((res) => res.json())
    .then((data) => {
      const currentYear = new Date().getFullYear();
      const yearSet = new Set((data.years || []).map(Number));
      yearSet.add(currentYear);
      const years = Array.from(yearSet).sort((a, b) => a - b);
      setYearOptions(yearSelect, years);
      setYearOptions(billingYearSelect, years);
    });
}

function getSelectedPeriod() {
  return periodSelect.value || "monthly";
}

function getSelectedMonthYear() {
  return {
    month: Number(monthSelect.value),
    year: Number(yearSelect.value),
  };
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

function persistFilters() {
  localStorage.setItem("selectedMonth", monthSelect.value);
  localStorage.setItem("selectedYear", yearSelect.value);
  localStorage.setItem("selectedPeriod", periodSelect.value);
  localStorage.setItem("taxpayerType", taxpayerSelect.value);
}

function restoreFilters(now) {
  const storedMonth = localStorage.getItem("selectedMonth");
  const storedYear = localStorage.getItem("selectedYear");
  const storedPeriod = localStorage.getItem("selectedPeriod");
  const storedTaxpayer = localStorage.getItem("taxpayerType");

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

  if (storedPeriod) {
    periodSelect.value = storedPeriod;
  }

  if (storedTaxpayer) {
    taxpayerSelect.value = storedTaxpayer === "empresa" ? "sociedad" : storedTaxpayer;
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
      supplier: "",
      base: "",
      vat: "21",
      vatAmount: "",
      total: "",
      analysisText: "",
      analysisPending: true,
      analysisError: false,
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
    supplierInput.addEventListener("input", () => {
      item.supplier = supplierInput.value;
      item.touched.supplier = true;
    });
    supplierTd.appendChild(supplierInput);

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
    vatSelect.disabled = item.analysisPending;
    vatSelect.addEventListener("change", () => {
      item.vat = resolveVatRateValue(vatSelect.value);
      item.touched.vat = true;
    });
    vatTd.appendChild(vatSelect);

    const vatAmountTd = document.createElement("td");
    const vatAmountInput = document.createElement("input");
    vatAmountInput.type = "number";
    vatAmountInput.step = "0.01";
    vatAmountInput.min = "0";
    vatAmountInput.placeholder = "0,00";
    vatAmountInput.value = item.vatAmount;
    vatAmountInput.disabled = item.analysisPending;
    vatAmountInput.addEventListener("input", () => {
      item.vatAmount = vatAmountInput.value;
      item.touched.vatAmount = true;
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
    totalInput.addEventListener("input", () => {
      item.total = totalInput.value;
      item.touched.total = true;
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
        ? "No se ha podido analizar la factura automáticamente. Puedes introducir los datos manualmente."
        : "Analizando factura… Las facturas escaneadas pueden tardar hasta 1 minuto.";
      statusWrapper.appendChild(message);
      statusTd.appendChild(statusWrapper);
      statusRow.appendChild(statusTd);
      uploadTableBody.appendChild(statusRow);
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
    if (!item.base || Number(item.base) < 0) {
      errors.push(`Base imponible inválida: ${item.file.name}`);
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

  fetch("/api/analyze-invoice", {
    method: "POST",
    body: formData,
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        item.analysisPending = false;
        item.analysisError = true;
        renderTable();
        return;
      }
      const extracted = data.extracted || {};
      item.storedFilename = data.storedFilename || "";
      item.analysisText = extracted.analysis_text || "";

      if (!item.touched.supplier && extracted.provider_name) {
        item.supplier = extracted.provider_name;
      }
      if (!item.touched.date && extracted.invoice_date) {
        item.date = extracted.invoice_date;
      }
      if (!item.paymentDate) {
        item.paymentDate = computePaymentDate(item.date, extracted.payment_date);
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

      item.analysisPending = false;
      const hasExtractedValue = [
        extracted.provider_name,
        extracted.invoice_date,
        extracted.base_amount,
        extracted.vat_rate,
        extracted.vat_amount,
        extracted.total_amount,
      ].some((value) => value !== null && value !== undefined && value !== "");
      item.analysisError = !hasExtractedValue && !item.analysisText;
      renderTable();
    })
    .catch(() => {
      item.analysisPending = false;
      item.analysisError = true;
      renderTable();
    });
}

function uploadPending() {
  if (pendingFiles.length === 0) {
    alert("No hay facturas para subir.");
    return;
  }

  const errors = validatePending();
  if (errors.length) {
    alert(errors.slice(0, 3).join("\n"));
    return;
  }

  uploadBtn.disabled = true;
  const payload = {
    entries: pendingFiles.map((item) => ({
      storedFilename: item.storedFilename,
      originalFilename: item.originalFilename,
      date: item.date,
      paymentDate: computePaymentDate(item.date, item.paymentDate),
      supplier: item.supplier.trim(),
      base: item.base,
      vat: item.vat,
      vatAmount: item.vatAmount,
      total: item.total,
      analysisText: item.analysisText,
    })),
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
  return fetch(`/api/summary?month=${month}&year=${year}`).then((res) => res.json());
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

function updateLineChart(labels, values, datasetLabel) {
  const ctx = document.getElementById("lineChart");
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

  if (!labels || labels.length === 0) {
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
  const data = currentBillingSummary;
  if (!data) {
    return;
  }
  const period = getSelectedPeriod();
  let labels = [];
  let values = [];
  if (period === "quarterly") {
    labels = data.monthlyTotals.map((item) => monthNames[item.month - 1]);
    values = data.monthlyTotals.map((item) => item.total);
  } else {
    const { month } = getSelectedMonthYear();
    labels = month ? [monthNames[month - 1]] : [];
    values = [billingBaseTotal];
  }

  const ctx = document.getElementById("billingLineChart");
  if (!billingLineChart) {
    billingLineChart = new Chart(ctx, {
      type: "line",
      data: {
        labels,
        datasets: [
          {
            label: "Base facturada",
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
  return fetch(`/api/billing/summary?month=${month}&year=${year}`).then((res) => res.json());
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
  const monthLabel = formatMonthYear(Number(monthSelect.value), Number(yearSelect.value));
  paymentDayTitle.textContent = `Pagos del ${day} ${monthLabel}`;

  let total = 0;
  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "payment-day-item";
    const supplier = document.createElement("span");
    supplier.textContent = item.supplier || "Proveedor";
    const dateLabel = document.createElement("span");
    dateLabel.textContent = item.payment_date;
    const amount = document.createElement("span");
    amount.textContent = formatCurrency(item.amount);
    row.appendChild(supplier);
    row.appendChild(dateLabel);
    row.appendChild(amount);
    paymentDayList.appendChild(row);
    total += Number(item.amount || 0);
  });
  paymentDayTotal.textContent = `Total del día: ${formatCurrency(total)}`;
}

function fetchBillingEntries(month, year) {
  return fetch(`/api/billing/entries?month=${month}&year=${year}`)
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
  return Promise.all([
    refreshSummary(),
    refreshBillingData(),
    refreshInvoices(),
    refreshPayments(),
    refreshNoInvoiceExpenses(),
    refreshAnnualTaxData(),
  ]);
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
  return fetch(`/api/invoices?month=${month}&year=${year}`)
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
  return fetch(`/api/payments?month=${month}&year=${year}`)
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
  const { month, year } = getSelectedMonthYear();
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
    vatTd.textContent = `${invoice.vat_rate}%`;

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

  const categorySelect = createExpenseCategorySelect(invoice.expense_category);

  dateTd.textContent = "";
  dateTd.appendChild(dateInput);
  supplierTd.textContent = "";
  supplierTd.appendChild(supplierInput);
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
  fetch(`/api/invoices/${invoiceId}`, {
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
      refreshAllData();
    })
    .catch(() => {
      alert("No se pudo actualizar la factura.");
    });
}

function deleteInvoice(invoiceId) {
  fetch(`/api/invoices/${invoiceId}`, {
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
  return fetch(`/api/expenses/no-invoice?month=${month}&year=${year}`)
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
  fetch(`/api/expenses/no-invoice/${expenseId}`, {
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
      refreshNoInvoiceExpenses();
    })
    .catch(() => {
      alert("No se pudo actualizar el gasto.");
    });
}

function deleteNoInvoiceExpense(expenseId) {
  fetch(`/api/expenses/no-invoice/${expenseId}`, {
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
  if (!entries.length) {
    billingEntriesEmpty.style.display = "block";
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

  fetch(`/api/billing/${entryId}`, {
    method: "PUT",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
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
  fetch(`/api/billing/${entryId}`, {
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
  const preTax = incomeTotal - expensesTotal;
  const taxpayerType = taxpayerSelect.value;
  const taxRate = taxpayerType === "sociedad" ? 0.25 : 0.15;
  const taxes = preTax > 0 ? preTax * taxRate : 0;
  const netResult = preTax - taxes;

  const incomeEl = document.getElementById("pnlIncome");
  const expensesEl = document.getElementById("pnlExpenses");
  const preTaxEl = document.getElementById("pnlPreTax");
  const taxesEl = document.getElementById("pnlTaxes");
  const netEl = document.getElementById("pnlNet");

  if (!incomeEl) {
    return;
  }

  incomeEl.textContent = formatCurrency(incomeTotal);
  expensesEl.textContent = formatCurrency(expensesTotal);
  preTaxEl.textContent = formatCurrency(preTax);
  taxesEl.textContent = formatCurrency(taxes);
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

  const incomeValue = document.getElementById("pnlIncome").textContent;
  const expensesValue = document.getElementById("pnlExpenses").textContent;
  const preTaxValue = document.getElementById("pnlPreTax").textContent;
  const taxesValue = document.getElementById("pnlTaxes").textContent;
  const netValue = document.getElementById("pnlNet").textContent;

  const doc = new jsPDF({ unit: "pt", format: "a4" });
  const margin = 48;
  let y = margin;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.text("Cuenta de resultados", margin, y);

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
    ["Ingresos", incomeValue],
    ["Gastos deducibles", expensesValue],
    ["Resultado antes de impuestos", preTaxValue],
    ["Impuestos estimados", taxesValue],
    ["Resultado neto", netValue],
  ];

  rows.forEach(([label, value]) => {
    doc.text(label, margin, y);
    doc.text(value, margin + 320, y, { align: "right" });
    y += 18;
  });

  const filename = `pnl_${periodLabel || "periodo"}.pdf`.replace(/\s+/g, "_");
  doc.save(filename);
}

function saveBillingEntry() {
  const month = Number(billingMonthSelect.value || monthSelect.value);
  const year = Number(billingYearSelect.value || yearSelect.value);
  const baseValue = billingBaseInput.value;
  const vatValue = billingVatSelect.value;

  if (!baseValue || Number(baseValue) < 0) {
    alert("Base facturada inválida.");
    return;
  }
  if (!month || !year) {
    alert("Selecciona mes y año.");
    return;
  }

  billingSaveBtn.disabled = true;
  fetch("/api/billing", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      month,
      year,
      base: baseValue,
      vat: vatValue,
    }),
  })
    .then((res) => res.json())
    .then((data) => {
      if (!data.ok) {
        alert((data.errors || ["Error al guardar."]).join("\n"));
        return;
      }
      billingBaseInput.value = "";
      refreshBillingData();
    })
    .catch(() => {
      alert("No se pudo guardar la facturación.");
    })
    .finally(() => {
      billingSaveBtn.disabled = false;
    });
}

function applyTaxpayerSelection(value) {
  const target = value === "sociedad" ? "is" : "irpf";
  document.querySelectorAll("[data-tax-module]").forEach((panel) => {
    panel.style.display = panel.dataset.taxModule === target ? "" : "none";
  });
}

function initTaxpayerSelector() {
  applyTaxpayerSelection(taxpayerSelect.value);
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
  const defaultSection = storedSection || "dashboard";
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
  document.getElementById("selectFiles").addEventListener("click", () => {
    fileInput.click();
  });
  document.getElementById("selectFolder").addEventListener("click", () => {
    folderInput.click();
  });
  fileInput.addEventListener("change", (event) => {
    addFiles(event.target.files);
    fileInput.value = "";
  });
  folderInput.addEventListener("change", (event) => {
    addFiles(event.target.files);
    folderInput.value = "";
  });

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

  uploadBtn.addEventListener("click", uploadPending);
  monthSelect.addEventListener("change", () => {
    persistFilters();
    refreshAllData();
  });
  yearSelect.addEventListener("change", () => {
    persistFilters();
    refreshAllData();
  });
  periodSelect.addEventListener("change", () => {
    document.body.classList.toggle("period-quarterly", getSelectedPeriod() === "quarterly");
    persistFilters();
    refreshAllData();
  });
  taxpayerSelect.addEventListener("change", () => {
    persistFilters();
    applyTaxpayerSelection(taxpayerSelect.value);
    updatePnlSummary();
  });
  billingSaveBtn.addEventListener("click", saveBillingEntry);
  noInvoiceSaveBtn.addEventListener("click", saveNoInvoiceExpense);
  if (exportPnlBtn) {
    exportPnlBtn.addEventListener("click", exportPnlPdf);
  }
}

function init() {
  const now = new Date();
  populateMonthSelects();
  bindEvents();
  initNavigation();
  loadYears().then(() => {
    restoreFilters(now);
    initTaxpayerSelector();
    document.body.classList.toggle("period-quarterly", getSelectedPeriod() === "quarterly");
    billingMonthSelect.value = monthSelect.value;
    billingYearSelect.value = yearSelect.value;
    if (!noInvoiceDate.value) {
      noInvoiceDate.value = now.toISOString().slice(0, 10);
    }
    refreshAllData();
  });
}

document.addEventListener("DOMContentLoaded", init);
