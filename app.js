const state = {
  files: [],
  rows: [],
  transactions: [],
  orders: [],
  productTitles: [],
  analysis: null,
  costs: {
    defaultProductCost: "",
    defaultShippingCost: "",
    defaultPackagingCost: "",
    defaultOtherCost: "",
    fixedCosts: "",
    productCosts: {},
    orderOverrides: {}
  }
};

const money = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD"
});

const els = {
  fileInput: document.getElementById("fileInput"),
  dropzone: document.getElementById("dropzone"),
  fileName: document.getElementById("fileName"),
  fileMeta: document.getElementById("fileMeta"),
  metricProfit: document.getElementById("metricProfit"),
  metricMargin: document.getElementById("metricMargin"),
  metricSales: document.getElementById("metricSales"),
  metricOrders: document.getElementById("metricOrders"),
  metricDeductions: document.getElementById("metricDeductions"),
  metricCosts: document.getElementById("metricCosts"),
  breakdownList: document.getElementById("breakdownList"),
  profitBridge: document.getElementById("profitBridge"),
  productCostBody: document.getElementById("productCostBody"),
  ordersBody: document.getElementById("ordersBody"),
  transactionsBody: document.getElementById("transactionsBody"),
  downloadReport: document.getElementById("downloadReport"),
  downloadSummary: document.getElementById("downloadSummary"),
  demoReset: document.getElementById("demoReset")
};

const requiredHeaders = ["Date", "Type", "Title", "Info", "Amount", "Fees & Taxes", "Net"];

init();

function init() {
  loadCosts();
  bindEvents();
  populateDefaultInputs();
  renderAll();
  refreshIcons();
}

function bindEvents() {
  els.fileInput.addEventListener("change", (event) => {
    const files = [...event.target.files];
    if (files.length) handleFiles(files);
  });

  ["dragenter", "dragover"].forEach((type) => {
    els.dropzone.addEventListener(type, (event) => {
      event.preventDefault();
      els.dropzone.classList.add("dragover");
    });
  });

  ["dragleave", "drop"].forEach((type) => {
    els.dropzone.addEventListener(type, (event) => {
      event.preventDefault();
      els.dropzone.classList.remove("dragover");
    });
  });

  els.dropzone.addEventListener("drop", (event) => {
    const files = [...event.dataTransfer.files];
    if (files.length) handleFiles(files);
  });

  document.querySelectorAll(".tab").forEach((button) => {
    button.addEventListener("click", () => setActiveTab(button.dataset.tab));
  });

  document.querySelectorAll("[data-cost-key]").forEach((input) => {
    input.addEventListener("input", () => {
      state.costs[input.dataset.costKey] = input.value;
      saveCosts();
      recalcAndRender();
    });
  });

  els.productCostBody.addEventListener("input", (event) => {
    const input = event.target.closest("[data-product-title]");
    if (!input) return;
    state.costs.productCosts[input.dataset.productTitle] = input.value;
    saveCosts();
    recalcAndRender();
  });

  els.ordersBody.addEventListener("input", (event) => {
    const input = event.target.closest("[data-order-id]");
    if (!input) return;
    const orderId = input.dataset.orderId;
    const field = input.dataset.orderField;
    state.costs.orderOverrides[orderId] = {
      ...(state.costs.orderOverrides[orderId] || {}),
      [field]: input.value
    };
    saveCosts();
    recalcAndRender();
  });

  els.downloadSummary.addEventListener("click", downloadSummaryCsv);
  els.downloadReport.addEventListener("click", downloadWorkbookReport);

  els.demoReset.addEventListener("click", () => {
    state.costs = {
      defaultProductCost: "",
      defaultShippingCost: "",
      defaultPackagingCost: "",
      defaultOtherCost: "",
      fixedCosts: "",
      productCosts: {},
      orderOverrides: {}
    };
    saveCosts();
    populateDefaultInputs();
    renderAll();
  });
}

async function handleFiles(files) {
  try {
    const loadedSources = [];
    const allRows = [];
    const allTransactions = [];

    for (const file of files) {
      const tables = await readSpreadsheet(file);
      let matchedInFile = 0;

      for (const table of tables) {
        const headerMap = getHeaderMap(table.rows[0] || []);
        const missing = requiredHeaders.filter((header) => headerMap[header] == null);

        if (missing.length) {
          if (tables.length === 1) {
            throw new Error(`${table.label}: missing expected column${missing.length > 1 ? "s" : ""}: ${missing.join(", ")}`);
          }
          continue;
        }

        const transactions = normalizeRows(table.rows, headerMap, table.label);
        matchedInFile += 1;
        allRows.push(...table.rows.slice(1));
        allTransactions.push(...transactions);
        loadedSources.push({
          label: table.label,
          rows: transactions.length
        });
      }

      if (!matchedInFile) {
        throw new Error(`${file.name}: no worksheet with Etsy statement columns was found.`);
      }
    }

    const { transactions, duplicatesSkipped } = dedupeTransactions(allTransactions);
    if (!transactions.length) {
      throw new Error("No Etsy transaction rows were found in the selected files.");
    }

    state.files = loadedSources;
    state.rows = allRows;
    state.transactions = transactions;
    state.analysis = analyzeTransactions(state.transactions);
    state.orders = state.analysis.orders;
    state.productTitles = state.analysis.productTitles;

    els.fileName.textContent = formatFileTitle(files, loadedSources);
    els.fileMeta.textContent = formatFileMeta(loadedSources, transactions.length, state.orders.length, duplicatesSkipped);
    els.fileMeta.title = loadedSources.map((source) => `${source.label}: ${source.rows} rows`).join("\n");
    document.body.classList.add("loaded");
    renderAll();
  } catch (error) {
    els.fileName.textContent = "Could not read files";
    els.fileMeta.textContent = error.message;
    els.fileMeta.title = "";
    document.body.classList.remove("loaded");
  }
}

function readSpreadsheet(file) {
  const extension = file.name.split(".").pop().toLowerCase();

  if (extension === "csv") {
    return file.text().then((text) => [{
      fileName: file.name,
      sheetName: "",
      label: file.name,
      rows: parseCsv(text.replace(/^\uFEFF/, ""))
    }]);
  }

  if (["xlsx", "xls"].includes(extension)) {
    if (!window.XLSX) {
      return Promise.reject(new Error("Excel support could not load. Try exporting from Etsy as CSV."));
    }
    return file.arrayBuffer().then((buffer) => {
      const workbook = XLSX.read(buffer, { type: "array" });
      return workbook.SheetNames.map((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        return {
          fileName: file.name,
          sheetName,
          label: `${file.name} / ${sheetName}`,
          rows: XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" })
        };
      }).filter((table) => table.rows.length);
    });
  }

  return Promise.reject(new Error("Please upload a .csv, .xlsx, or .xls file."));
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (quoted) {
      if (char === "\"" && next === "\"") {
        field += "\"";
        index += 1;
      } else if (char === "\"") {
        quoted = false;
      } else {
        field += char;
      }
      continue;
    }

    if (char === "\"") {
      quoted = true;
    } else if (char === ",") {
      row.push(field);
      field = "";
    } else if (char === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (char !== "\r") {
      field += char;
    }
  }

  if (field || row.length) {
    row.push(field);
    rows.push(row);
  }

  return rows.filter((csvRow) => csvRow.some((cell) => String(cell).trim() !== ""));
}

function getHeaderMap(headers) {
  return headers.reduce((map, header, index) => {
    const clean = String(header).replace(/^\uFEFF/, "").trim();
    map[clean] = index;
    return map;
  }, {});
}

function normalizeRows(rows, headerMap, source) {
  return rows.slice(1).map((row, index) => {
    const get = (header) => String(row[headerMap[header]] ?? "").trim();
    const title = get("Title");
    const info = get("Info");
    const orderId = extractOrderId(`${title} ${info}`);

    return {
      date: get("Date"),
      type: get("Type"),
      title,
      info,
      currency: get("Currency") || "USD",
      amount: parseMoney(get("Amount")),
      feesTaxes: parseMoney(get("Fees & Taxes")),
      net: parseMoney(get("Net")),
      taxDetails: get("Tax Details"),
      orderId,
      productTitle: extractProductTitle(title),
      source,
      sourceRow: index + 2
    };
  });
}

function analyzeTransactions(transactions) {
  const orderMap = new Map();
  const productCounts = new Map();
  const totals = {
    grossSales: 0,
    statementNet: 0,
    deposits: 0,
    transactionFees: 0,
    processingFees: 0,
    listingFees: 0,
    marketing: 0,
    taxes: 0,
    otherEtsy: 0,
    etsyDeductions: 0
  };

  transactions.forEach((tx) => {
    if (tx.type !== "Deposit") {
      totals.statementNet += tx.net;
    } else {
      totals.deposits += extractMoneyFromText(tx.title);
    }

    if (tx.type === "Sale") totals.grossSales += tx.amount;
    if (tx.type !== "Sale" && tx.type !== "Deposit" && tx.net < 0) totals.etsyDeductions += Math.abs(tx.net);

    const title = tx.title.toLowerCase();
    if (tx.type === "Fee" && title.startsWith("transaction fee:")) {
      totals.transactionFees += Math.abs(tx.net);
    } else if (tx.type === "Fee" && title.includes("processing fee")) {
      totals.processingFees += Math.abs(tx.net);
    } else if (tx.type === "Fee" && title.includes("listing fee")) {
      totals.listingFees += Math.abs(tx.net);
    } else if (tx.type === "Marketing") {
      totals.marketing += Math.abs(tx.net);
    } else if (tx.type === "Tax") {
      totals.taxes += Math.abs(tx.net);
    } else if (tx.type !== "Sale" && tx.type !== "Deposit" && tx.net < 0) {
      totals.otherEtsy += Math.abs(tx.net);
    }

    if (!tx.orderId) return;
    if (!orderMap.has(tx.orderId)) {
      orderMap.set(tx.orderId, {
        id: tx.orderId,
        date: tx.date,
        dateValue: parseDateValue(tx.date),
        gross: 0,
        net: 0,
        etsyDeductions: 0,
        taxes: 0,
        fees: 0,
        marketing: 0,
        productTitles: new Set()
      });
    }

    const order = orderMap.get(tx.orderId);
    if (parseDateValue(tx.date) > order.dateValue) {
      order.date = tx.date;
      order.dateValue = parseDateValue(tx.date);
    }

    order.net += tx.net;
    if (tx.type === "Sale") order.gross += tx.amount;
    if (tx.type !== "Sale" && tx.net < 0) order.etsyDeductions += Math.abs(tx.net);
    if (tx.type === "Tax") order.taxes += Math.abs(tx.net);
    if (tx.type === "Fee") order.fees += Math.abs(tx.net);
    if (tx.type === "Marketing") order.marketing += Math.abs(tx.net);
    if (tx.productTitle) {
      order.productTitles.add(tx.productTitle);
      productCounts.set(tx.productTitle, (productCounts.get(tx.productTitle) || 0) + 1);
    }
  });

  const orders = [...orderMap.values()]
    .map((order) => ({
      ...order,
      productTitles: [...order.productTitles],
      productTitle: [...order.productTitles][0] || "Unknown item"
    }))
    .sort((a, b) => b.dateValue - a.dateValue);

  const productTitles = [...productCounts.entries()]
    .map(([title, count]) => ({ title, count }))
    .sort((a, b) => b.count - a.count || a.title.localeCompare(b.title));

  return { totals, orders, productTitles };
}

function recalcAndRender() {
  if (state.analysis) {
    state.analysis = analyzeTransactions(state.transactions);
    state.orders = state.analysis.orders;
    state.productTitles = state.analysis.productTitles;
  }
  renderMetrics();
  renderSummary();
  renderOrders();
}

function renderAll() {
  renderMetrics();
  renderSummary();
  renderProductCosts();
  renderOrders();
  renderTransactions();
  refreshIcons();
}

function renderMetrics() {
  const totals = calculatedTotals();
  els.metricProfit.textContent = formatMoney(totals.estimatedProfit);
  els.metricProfit.classList.toggle("bad", totals.estimatedProfit < 0);
  els.metricMargin.textContent = `${formatPercent(totals.margin)} margin`;
  els.metricSales.textContent = formatMoney(totals.grossSales);
  els.metricOrders.textContent = `${state.orders.length} order${state.orders.length === 1 ? "" : "s"}`;
  els.metricDeductions.textContent = formatMoney(totals.etsyDeductions);
  els.metricCosts.textContent = formatMoney(totals.totalCosts);
}

function renderSummary() {
  const totals = calculatedTotals();
  const rows = [
    ["Gross sales", totals.grossSales, "green"],
    ["Transaction fees", totals.transactionFees, "red"],
    ["Processing fees", totals.processingFees, "red"],
    ["Listing fees", totals.listingFees, "amber"],
    ["Ads and offsite ads", totals.marketing, "blue"],
    ["Buyer tax remitted", totals.taxes, "teal"],
    ["Other Etsy adjustments", totals.otherEtsy, "red"],
    ["Your entered costs", totals.totalCosts, "amber"]
  ];
  const max = Math.max(...rows.map(([, value]) => Math.abs(value)), 1);

  els.breakdownList.innerHTML = rows.map(([label, value, color]) => `
    <div class="breakdown-row">
      <span class="breakdown-label">${escapeHtml(label)}</span>
      <span class="bar-track"><span class="bar-fill ${color}" style="width: ${Math.max(2, Math.abs(value) / max * 100)}%"></span></span>
      <span class="number">${formatMoney(value)}</span>
    </div>
  `).join("");

  const bridgeRows = [
    ["Gross sales", totals.grossSales],
    ["Minus Etsy deductions", -totals.etsyDeductions],
    ["Net after Etsy rows", totals.statementNet],
    ["Minus product and fulfillment costs", -totals.variableCosts],
    ["Minus fixed shop costs", -totals.fixedCosts],
    ["Estimated profit", totals.estimatedProfit]
  ];

  els.profitBridge.innerHTML = bridgeRows.map(([label, value]) => `
    <div class="bridge-row">
      <span>${escapeHtml(label)}</span>
      <strong class="${value < 0 ? "bad" : "good"}">${formatMoney(value)}</strong>
    </div>
  `).join("");
}

function renderProductCosts() {
  if (!state.productTitles.length) {
    els.productCostBody.innerHTML = `<tr><td colspan="3">Upload a statement to detect product titles.</td></tr>`;
    return;
  }

  els.productCostBody.innerHTML = state.productTitles.map(({ title, count }) => `
    <tr>
      <td class="truncate" title="${escapeAttr(title)}">${escapeHtml(title)}</td>
      <td class="number">${count}</td>
      <td>
        <input class="mini-input" data-product-title="${escapeAttr(title)}" type="number" min="0" step="0.01" placeholder="${escapeAttr(state.costs.defaultProductCost || "0.00")}" value="${escapeAttr(state.costs.productCosts[title] || "")}">
      </td>
    </tr>
  `).join("");
}

function renderOrders() {
  if (!state.orders.length) {
    els.ordersBody.innerHTML = `<tr><td colspan="11">Upload a statement to see order profit.</td></tr>`;
    return;
  }

  els.ordersBody.innerHTML = state.orders.map((order) => {
    const costs = calculateOrderCosts(order);
    const overrides = state.costs.orderOverrides[order.id] || {};
    const profit = order.net - costs.total;
    return `
      <tr>
        <td><span class="order-id">#${escapeHtml(order.id)}</span></td>
        <td>${escapeHtml(order.date)}</td>
        <td class="truncate" title="${escapeAttr(order.productTitles.join(", ") || "Unknown item")}">${escapeHtml(order.productTitle)}</td>
        <td class="number">${formatMoney(order.gross)}</td>
        <td class="number">${formatMoney(order.etsyDeductions)}</td>
        <td><input class="mini-input qty-input" data-order-id="${escapeAttr(order.id)}" data-order-field="qty" type="number" min="1" step="1" placeholder="1" value="${escapeAttr(overrides.qty || "")}"></td>
        <td><input class="mini-input" data-order-id="${escapeAttr(order.id)}" data-order-field="product" type="number" min="0" step="0.01" placeholder="${escapeAttr(formatPlain(costs.productUnit))}" value="${escapeAttr(overrides.product || "")}"></td>
        <td><input class="mini-input" data-order-id="${escapeAttr(order.id)}" data-order-field="shipping" type="number" min="0" step="0.01" placeholder="${escapeAttr(formatPlain(costs.shipping))}" value="${escapeAttr(overrides.shipping || "")}"></td>
        <td><input class="mini-input" data-order-id="${escapeAttr(order.id)}" data-order-field="packaging" type="number" min="0" step="0.01" placeholder="${escapeAttr(formatPlain(costs.packaging))}" value="${escapeAttr(overrides.packaging || "")}"></td>
        <td><input class="mini-input" data-order-id="${escapeAttr(order.id)}" data-order-field="other" type="number" min="0" step="0.01" placeholder="${escapeAttr(formatPlain(costs.other))}" value="${escapeAttr(overrides.other || "")}"></td>
        <td class="number ${profit < 0 ? "bad" : "good"}">${formatMoney(profit)}</td>
      </tr>
    `;
  }).join("");
}

function renderTransactions() {
  if (!state.transactions.length) {
    els.transactionsBody.innerHTML = `<tr><td colspan="8">Upload a statement to see cleaned transactions.</td></tr>`;
    return;
  }

  els.transactionsBody.innerHTML = state.transactions.map((tx) => `
    <tr>
      <td class="truncate" title="${escapeAttr(tx.source || "")}">${escapeHtml(tx.source || "")}</td>
      <td>${escapeHtml(tx.date)}</td>
      <td><span class="type-pill ${typeClass(tx.type)}">${escapeHtml(tx.type)}</span></td>
      <td class="truncate" title="${escapeAttr(tx.title)}">${escapeHtml(tx.title)}</td>
      <td>${tx.orderId ? `<span class="order-id">#${escapeHtml(tx.orderId)}</span>` : ""}</td>
      <td class="number">${formatMoney(tx.amount)}</td>
      <td class="number">${formatMoney(tx.feesTaxes)}</td>
      <td class="number ${tx.net < 0 ? "bad" : "good"}">${formatMoney(tx.net)}</td>
    </tr>
  `).join("");
}

function calculatedTotals() {
  const base = state.analysis?.totals || {
    grossSales: 0,
    statementNet: 0,
    deposits: 0,
    transactionFees: 0,
    processingFees: 0,
    listingFees: 0,
    marketing: 0,
    taxes: 0,
    otherEtsy: 0,
    etsyDeductions: 0
  };

  const variableCosts = state.orders.reduce((sum, order) => sum + calculateOrderCosts(order).total, 0);
  const fixedCosts = toNumber(state.costs.fixedCosts);
  const totalCosts = variableCosts + fixedCosts;
  const estimatedProfit = base.statementNet - totalCosts;
  const margin = base.grossSales ? estimatedProfit / base.grossSales : 0;

  return { ...base, variableCosts, fixedCosts, totalCosts, estimatedProfit, margin };
}

function calculateOrderCosts(order) {
  const overrides = state.costs.orderOverrides[order.id] || {};
  const title = order.productTitle;
  const productRule = title && state.costs.productCosts[title] !== "" ? state.costs.productCosts[title] : "";
  const productUnit = valueOrFallback(overrides.product, valueOrFallback(productRule, state.costs.defaultProductCost));
  const qty = Math.max(1, valueOrFallback(overrides.qty, 1));
  const shipping = valueOrFallback(overrides.shipping, state.costs.defaultShippingCost);
  const packaging = valueOrFallback(overrides.packaging, state.costs.defaultPackagingCost);
  const other = valueOrFallback(overrides.other, state.costs.defaultOtherCost);
  const total = productUnit * qty + shipping + packaging + other;
  return { productUnit, qty, shipping, packaging, other, total };
}

function setActiveTab(name) {
  document.querySelectorAll(".tab").forEach((button) => {
    button.classList.toggle("active", button.dataset.tab === name);
  });
  document.querySelectorAll(".view").forEach((view) => {
    view.classList.toggle("active", view.id === `${name}View`);
  });
}

function downloadSummaryCsv() {
  const totals = calculatedTotals();
  const lines = [
    ["Metric", "Value"],
    ["Sources loaded", state.files.map((source) => source.label).join(" | ")],
    ["Gross sales", totals.grossSales],
    ["Etsy deductions", totals.etsyDeductions],
    ["Net after Etsy rows", totals.statementNet],
    ["Your variable costs", totals.variableCosts],
    ["Fixed shop costs", totals.fixedCosts],
    ["Estimated profit", totals.estimatedProfit],
    [],
    ["Order", "Date", "Item", "Gross", "Etsy deductions", "Entered costs", "Estimated profit"]
  ];

  state.orders.forEach((order) => {
    const costs = calculateOrderCosts(order);
    lines.push([
      order.id,
      order.date,
      order.productTitles.join(" | "),
      order.gross,
      order.etsyDeductions,
      costs.total,
      order.net - costs.total
    ]);
  });

  const csv = lines.map((line) => line.map(csvCell).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "etsy-profit-summary.csv";
  link.click();
  URL.revokeObjectURL(link.href);
}

function downloadWorkbookReport() {
  if (!state.transactions.length) {
    els.fileName.textContent = "No report to download yet";
    els.fileMeta.textContent = "Upload Etsy statement files first, then download the Excel report.";
    return;
  }

  if (!window.XLSX) {
    downloadSummaryCsv();
    return;
  }

  const workbook = XLSX.utils.book_new();
  workbook.Props = {
    Title: "Etsy Profit Clarity Report",
    Subject: "Etsy sales, fees, costs, and profit",
    Author: "Etsy Profit Clarity",
    CreatedDate: new Date()
  };

  appendSheet(workbook, "Summary", buildSummarySheet(), [
    { wch: 30 },
    { wch: 18 },
    { wch: 52 }
  ]);
  appendSheet(workbook, "Orders", buildOrdersSheet(), [
    { wch: 16 },
    { wch: 18 },
    { wch: 42 },
    { wch: 14 },
    { wch: 16 },
    { wch: 14 },
    { wch: 10 },
    { wch: 18 },
    { wch: 16 },
    { wch: 16 },
    { wch: 14 },
    { wch: 18 },
    { wch: 18 }
  ]);
  appendSheet(workbook, "Transactions", buildTransactionsSheet(), [
    { wch: 30 },
    { wch: 18 },
    { wch: 14 },
    { wch: 48 },
    { wch: 16 },
    { wch: 14 },
    { wch: 16 },
    { wch: 14 },
    { wch: 36 }
  ]);
  appendSheet(workbook, "Costs", buildCostsSheet(), [
    { wch: 44 },
    { wch: 16 },
    { wch: 16 },
    { wch: 52 }
  ]);

  const dateStamp = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `etsy-profit-report-${dateStamp}.xlsx`);
}

function buildSummarySheet() {
  const totals = calculatedTotals();
  return [
    ["Etsy Profit Clarity Report"],
    ["Generated", new Date().toLocaleString()],
    ["Sources", state.files.map((source) => source.label).join(" | ") || "Uploaded Etsy statement"],
    [],
    ["Metric", "Value", "Notes"],
    ["Gross sales", roundMoney(totals.grossSales), "Total Etsy sale/payment rows before deductions"],
    ["Etsy deductions", roundMoney(totals.etsyDeductions), "Fees, buyer tax, ads, listing fees, and Etsy adjustments"],
    ["Net after Etsy rows", roundMoney(totals.statementNet), "Statement net excluding bank deposit transfer rows"],
    ["Product and fulfillment costs", roundMoney(totals.variableCosts), "Costs entered by default, product, or order"],
    ["Fixed shop costs", roundMoney(totals.fixedCosts), "Fixed costs entered in the app"],
    ["Estimated profit", roundMoney(totals.estimatedProfit), "Net after Etsy rows minus entered seller costs"],
    ["Profit margin", totals.margin, "Estimated profit divided by gross sales"],
    ["Orders", state.orders.length, "Grouped by Etsy order number"],
    ["Transactions", state.transactions.length, "Cleaned rows used in this report"],
    [],
    ["Deduction Breakdown", "Amount", "Included in Etsy deductions"],
    ["Transaction fees", roundMoney(totals.transactionFees), "Yes"],
    ["Processing fees", roundMoney(totals.processingFees), "Yes"],
    ["Listing fees", roundMoney(totals.listingFees), "Yes"],
    ["Ads and offsite ads", roundMoney(totals.marketing), "Yes"],
    ["Buyer tax remitted", roundMoney(totals.taxes), "Yes"],
    ["Other Etsy adjustments", roundMoney(totals.otherEtsy), "Yes"],
    [],
    ["Bank deposits", roundMoney(totals.deposits), "Tracked separately because deposits are transfers, not revenue or expense"]
  ];
}

function buildOrdersSheet() {
  const rows = [[
    "Order",
    "Date",
    "Item",
    "Gross Sales",
    "Etsy Deductions",
    "Net After Etsy",
    "Qty",
    "Product Unit Cost",
    "Shipping Cost",
    "Packaging Cost",
    "Other Cost",
    "Total Entered Costs",
    "Estimated Profit"
  ]];

  state.orders.forEach((order) => {
    const costs = calculateOrderCosts(order);
    rows.push([
      `#${order.id}`,
      order.date,
      order.productTitles.join(" | ") || order.productTitle,
      roundMoney(order.gross),
      roundMoney(order.etsyDeductions),
      roundMoney(order.net),
      costs.qty,
      roundMoney(costs.productUnit),
      roundMoney(costs.shipping),
      roundMoney(costs.packaging),
      roundMoney(costs.other),
      roundMoney(costs.total),
      roundMoney(order.net - costs.total)
    ]);
  });

  return rows;
}

function buildTransactionsSheet() {
  const rows = [[
    "Source",
    "Date",
    "Type",
    "Title",
    "Order",
    "Amount",
    "Fees & Taxes",
    "Net",
    "Info"
  ]];

  state.transactions.forEach((tx) => {
    rows.push([
      tx.source || "",
      tx.date,
      tx.type,
      tx.title,
      tx.orderId ? `#${tx.orderId}` : "",
      roundMoney(tx.amount),
      roundMoney(tx.feesTaxes),
      roundMoney(tx.net),
      tx.info
    ]);
  });

  return rows;
}

function buildCostsSheet() {
  const rows = [
    ["Cost Input", "Value", "Applies To", "Notes"],
    ["Default product cost", roundMoney(toNumber(state.costs.defaultProductCost)), "Every order unless overridden", ""],
    ["Default shipping label cost", roundMoney(toNumber(state.costs.defaultShippingCost)), "Every order unless overridden", ""],
    ["Default packaging cost", roundMoney(toNumber(state.costs.defaultPackagingCost)), "Every order unless overridden", ""],
    ["Default other cost", roundMoney(toNumber(state.costs.defaultOtherCost)), "Every order unless overridden", ""],
    ["Fixed shop costs", roundMoney(toNumber(state.costs.fixedCosts)), "Whole report", ""],
    [],
    ["Product Rule", "Unit Cost", "Detected Orders", "Notes"]
  ];

  state.productTitles.forEach(({ title, count }) => {
    rows.push([
      title,
      roundMoney(toNumber(state.costs.productCosts[title])),
      count,
      state.costs.productCosts[title] === "" || state.costs.productCosts[title] == null ? "Using default product cost" : "Product-specific cost"
    ]);
  });

  rows.push([]);
  rows.push(["Order Override", "Value", "Field", "Notes"]);
  state.orders.forEach((order) => {
    const overrides = state.costs.orderOverrides[order.id] || {};
    Object.entries(overrides).forEach(([field, value]) => {
      if (String(value).trim() === "") return;
      rows.push([`#${order.id}`, field === "qty" ? toNumber(value) : roundMoney(toNumber(value)), field, order.productTitle]);
    });
  });

  return rows;
}

function appendSheet(workbook, name, rows, columns) {
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet["!cols"] = columns;
  applyWorkbookFormats(sheet);
  XLSX.utils.book_append_sheet(workbook, sheet, name);
}

function applyWorkbookFormats(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = sheet[address];
      if (!cell || typeof cell.v !== "number") continue;

      const header = String(sheet[XLSX.utils.encode_cell({ r: 0, c: col })]?.v || "");
      const label = String(sheet[XLSX.utils.encode_cell({ r: row, c: 0 })]?.v || "");
      if (/margin/i.test(header) || /margin/i.test(label)) {
        cell.z = "0.0%";
      } else if (/qty|orders|transactions/i.test(header) || /orders|transactions/i.test(label)) {
        cell.z = "#,##0";
      } else {
        cell.z = "$#,##0.00";
      }
    }
  }
}

function dedupeTransactions(transactions) {
  const seenRows = new Set();
  const seenSourcesByExactRow = new Map();
  const unique = [];

  transactions.forEach((tx) => {
    const exactRowKey = [
      tx.date,
      tx.type,
      tx.title,
      tx.info,
      tx.currency,
      tx.amount,
      tx.feesTaxes,
      tx.net,
      tx.taxDetails
    ].join("|");
    const sourceRowKey = `${exactRowKey}|${tx.source}|${tx.sourceRow}`;
    const seenSources = seenSourcesByExactRow.get(exactRowKey) || new Set();

    if (seenRows.has(sourceRowKey)) return;
    if (seenSources.size && !seenSources.has(tx.source)) return;

    seenRows.add(sourceRowKey);
    seenSources.add(tx.source);
    seenSourcesByExactRow.set(exactRowKey, seenSources);
    unique.push(tx);
  });

  return {
    transactions: unique,
    duplicatesSkipped: transactions.length - unique.length
  };
}

function formatFileTitle(files, loadedSources) {
  if (files.length === 1 && loadedSources.length === 1) {
    return loadedSources[0].label;
  }

  if (files.length === 1) {
    return `${loadedSources.length} worksheets loaded`;
  }

  return `${files.length} files loaded`;
}

function formatFileMeta(sources, rowCount, orderCount, duplicatesSkipped) {
  const sourceLabel = `${sources.length} source${sources.length === 1 ? "" : "s"}`;
  const duplicateLabel = duplicatesSkipped ? `, ${duplicatesSkipped} duplicate row${duplicatesSkipped === 1 ? "" : "s"} skipped` : "";
  return `${sourceLabel}, ${rowCount} rows, ${orderCount} order${orderCount === 1 ? "" : "s"}${duplicateLabel}`;
}

function populateDefaultInputs() {
  document.querySelectorAll("[data-cost-key]").forEach((input) => {
    input.value = state.costs[input.dataset.costKey] || "";
  });
}

function loadCosts() {
  try {
    const stored = JSON.parse(localStorage.getItem("etsyProfitClarityCosts") || "{}");
    state.costs = { ...state.costs, ...stored };
    state.costs.productCosts = stored.productCosts || {};
    state.costs.orderOverrides = stored.orderOverrides || {};
  } catch {
    saveCosts();
  }
}

function saveCosts() {
  localStorage.setItem("etsyProfitClarityCosts", JSON.stringify(state.costs));
}

function parseMoney(value) {
  const text = String(value ?? "").trim();
  if (!text || text === "--") return 0;
  const negative = text.includes("-") || /^\(.+\)$/.test(text);
  const numeric = Number(text.replace(/[^0-9.]/g, ""));
  if (!Number.isFinite(numeric)) return 0;
  return negative ? -numeric : numeric;
}

function extractMoneyFromText(text) {
  const match = String(text).match(/-?\$?[\d,]+(?:\.\d{2})?/);
  return match ? parseMoney(match[0]) : 0;
}

function extractOrderId(text) {
  const match = String(text).match(/Order #(\d+)/i);
  return match ? match[1] : "";
}

function extractProductTitle(title) {
  const clean = String(title || "").trim();
  if (!/^transaction fee:/i.test(clean)) return "";
  const product = clean.replace(/^transaction fee:\s*/i, "").trim();
  if (/^shipping$/i.test(product)) return "";
  return product || "";
}

function parseDateValue(dateText) {
  const value = Date.parse(dateText);
  return Number.isFinite(value) ? value : 0;
}

function valueOrFallback(value, fallback) {
  return String(value ?? "").trim() === "" ? toNumber(fallback) : toNumber(value);
}

function toNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : 0;
}

function formatMoney(value) {
  return money.format(Number.isFinite(value) ? value : 0);
}

function formatPlain(value) {
  return Number.isFinite(value) && value > 0 ? value.toFixed(2) : "";
}

function roundMoney(value) {
  return Math.round((Number.isFinite(value) ? value : 0) * 100) / 100;
}

function formatPercent(value) {
  return `${((Number.isFinite(value) ? value : 0) * 100).toFixed(1)}%`;
}

function typeClass(type) {
  return String(type || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "-");
}

function csvCell(value) {
  if (value == null) return "";
  const text = String(value);
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, "\"\"")}"` : text;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeAttr(value) {
  return escapeHtml(value).replace(/`/g, "&#096;");
}

function refreshIcons() {
  if (window.lucide) {
    window.lucide.createIcons();
  }
}
