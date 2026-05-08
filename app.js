const PDFJS_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.mjs";
const PDFJS_WORKER_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.mjs";

const ACCOUNT_NAMES = {
  "1000": "資産合計",
  "1100": "流動資産",
  "1101": "現金及び預金",
  "1111": "現金",
  "1113": "普通預金",
  "1114": "定期預金",
  "1115": "その他預金",
  "1120": "売上債権",
  "1121": "売掛金",
  "1122": "受取手形",
  "1124": "電子記録債権",
  "2000": "負債・純資産合計",
  "2100": "流動負債",
  "2112": "買掛金",
  "2113": "支払手形",
  "2114": "電子記録債務",
  "2117": "短期借入金",
  "2125": "一年以内返済長期借入金",
  "2200": "固定負債",
  "2211": "長期借入金",
  "2212": "役員借入金",
  "2213": "金融機関借入金",
  "3000": "負債合計",
  "3100": "純資産",
  "4000": "売上高合計",
  "4111": "売上高",
  "4115": "売上値引戻り高",
  "5000": "売上原価",
  "5100": "期首棚卸高",
  "5200": "仕入高",
  "5400": "製造原価",
  "6000": "販売費及び一般管理費"
};

const CATEGORY_LABELS = {
  cash: "現預金",
  ar: "売掛金",
  bill: "受取手形",
  eReceivable: "電子記録債権",
  loan: "借入金",
  sales: "売上",
  cost: "原価",
  fixed: "固定費",
  other: "その他"
};

const els = {
  pdfInput: document.querySelector("#pdfInput"),
  rawTextInput: document.querySelector("#rawTextInput"),
  parseTextButton: document.querySelector("#parseTextButton"),
  sampleButton: document.querySelector("#sampleButton"),
  exportButton: document.querySelector("#exportButton"),
  recalculateButton: document.querySelector("#recalculateButton"),
  importStatus: document.querySelector("#importStatus"),
  monthMetric: document.querySelector("#monthMetric"),
  cashMetric: document.querySelector("#cashMetric"),
  receivableMetric: document.querySelector("#receivableMetric"),
  loanMetric: document.querySelector("#loanMetric"),
  basisText: document.querySelector("#basisText"),
  rowsTable: document.querySelector("#rowsTable tbody"),
  rowCount: document.querySelector("#rowCount"),
  cashflowHead: document.querySelector("#cashflowTable thead"),
  cashflowBody: document.querySelector("#cashflowTable tbody"),
  emptyTableTemplate: document.querySelector("#emptyTableTemplate")
};

const inputIds = [
  "cashCodes",
  "arCodes",
  "billCodes",
  "eReceivableCodes",
  "loanCodes",
  "salesForecast",
  "collectCurrent",
  "collectNext",
  "collectAfterNext",
  "billMaturity",
  "eReceivableMaturity",
  "costRate",
  "fixedCosts",
  "loanRepayment",
  "openingCash",
  "forecastMonths"
];

const state = {
  months: [],
  rows: [],
  forecastRows: [],
  manual: {
    otherIn: [],
    otherOut: [],
    loanRepayment: []
  }
};

let pdfjsPromise = null;

function formatYen(value) {
  const amount = Math.round(Number(value) || 0);
  return amount.toLocaleString("ja-JP");
}

function parseAmount(value) {
  if (value == null || value === "") return 0;
  const text = String(value).replace(/[,\s円]/g, "").replace(/△/g, "-").replace(/▲/g, "-");
  return Number(text) || 0;
}

function average(values) {
  const nums = values.filter((value) => Number.isFinite(value));
  if (!nums.length) return 0;
  return nums.reduce((sum, value) => sum + value, 0) / nums.length;
}

function last(values, fallback = 0) {
  return values.length ? values[values.length - 1] : fallback;
}

function parseCodes(value) {
  return String(value || "")
    .split(/[,\s、]+/)
    .map((code) => code.trim())
    .filter(Boolean);
}

function normalizeText(text) {
  return String(text || "")
    .replace(/\u0000/g, "")
    .replace(/[　\t]+/g, " ")
    .replace(/\r/g, "\n")
    .replace(/([0-9]{4})\s*年\s*([0-9]{1,2})\s*月/g, "$1/$2")
    .replace(/([0-9]{4})[^0-9\n]{1,8}([0-9]{2})[^0-9\n]{0,4}/g, "$1/$2");
}

function monthKey(year, month) {
  return `${year}/${String(month).padStart(2, "0")}`;
}

function nextMonth(key, offset = 1) {
  const [year, month] = key.split("/").map(Number);
  const date = new Date(year, month - 1 + offset, 1);
  return monthKey(date.getFullYear(), date.getMonth() + 1);
}

function extractMonths(text) {
  const found = [];
  const seen = new Set();
  const patterns = [
    /(20\d{2})[\/年\-. ]{1,4}(0?[1-9]|1[0-2])\s*月?/g,
    /(20\d{2})\D{0,8}(0[1-9]|1[0-2])\D/g
  ];

  patterns.forEach((pattern) => {
    let match;
    while ((match = pattern.exec(text)) !== null) {
      const key = monthKey(Number(match[1]), Number(match[2]));
      if (!seen.has(key)) {
        seen.add(key);
        found.push(key);
      }
    }
  });

  const ordered = found
    .filter((key) => Number(key.slice(0, 4)) >= 2020)
    .sort((a, b) => new Date(`${a}/01`) - new Date(`${b}/01`));

  return ordered.slice(-12);
}

function splitRows(text) {
  const rowPattern = /\((?:△)?([0-9]{4})\)/g;
  const matches = [...text.matchAll(rowPattern)];
  return matches.map((match, index) => {
    const start = match.index || 0;
    const end = index + 1 < matches.length ? matches[index + 1].index || text.length : text.length;
    return {
      code: match[1],
      raw: text.slice(start, end)
    };
  });
}

function extractLabel(segment, code) {
  const afterCode = segment.replace(new RegExp(`^.*?\\(${code}\\)`), "");
  const beforeAmount = afterCode.split(/[-△▲]?\d{1,3}(?:,\d{3})+/)[0] || "";
  const cleaned = beforeAmount
    .replace(/[0-9.,%()\/:*-]/g, "")
    .replace(/\s+/g, " ")
    .trim();
  return ACCOUNT_NAMES[code] || cleaned || `科目 ${code}`;
}

function extractAmounts(segment, monthCount) {
  const spaced = segment
    .replace(/(\d{1,3}\.\d)\s*(\d{1,3},\d{3})/g, "$1 $2")
    .replace(/(\d{1,3},\d{3})(\d{1,3}\.\d)/g, "$1 $2");
  const tokens = [...spaced.matchAll(/[△▲-]?\d{1,3}(?:,\d{3})+/g)].map((match) => parseAmount(match[0]));
  if (tokens.length <= monthCount) return tokens;
  return tokens.slice(-monthCount);
}

function parseBalanceText(rawText) {
  const text = normalizeText(rawText);
  let months = extractMonths(text);
  const segments = splitRows(text);

  if (!months.length && segments.length) {
    const maxAmountCount = Math.max(...segments.map((segment) => extractAmounts(segment.raw, 24).length), 0);
    const count = Math.min(Math.max(maxAmountCount, 6), 12);
    const today = new Date();
    months = Array.from({ length: count }, (_, index) => {
      const date = new Date(today.getFullYear(), today.getMonth() - count + index + 1, 1);
      return monthKey(date.getFullYear(), date.getMonth() + 1);
    });
  }

  const monthCount = months.length || 12;
  const rows = segments
    .map((segment) => {
      const values = extractAmounts(segment.raw, monthCount);
      return {
        code: segment.code,
        name: extractLabel(segment.raw, segment.code),
        values,
        latest: last(values, 0)
      };
    })
    .filter((row) => row.values.length >= Math.min(3, monthCount));

  return { months, rows: dedupeRows(rows) };
}

function dedupeRows(rows) {
  const byCode = new Map();
  rows.forEach((row) => {
    const current = byCode.get(row.code);
    if (!current || row.values.length > current.values.length) byCode.set(row.code, row);
  });
  return [...byCode.values()].sort((a, b) => a.code.localeCompare(b.code, "ja"));
}

function getInputs() {
  return Object.fromEntries(inputIds.map((id) => [id, document.querySelector(`#${id}`).value]));
}

function findRows(codes) {
  const set = new Set(codes);
  return state.rows.filter((row) => set.has(row.code));
}

function pickSeries(codes, preferAggregate = true) {
  const rows = findRows(codes);
  if (!rows.length) return Array(state.months.length).fill(0);
  if (preferAggregate) {
    const exact = rows.find((row) => row.code.endsWith("00") || row.code.endsWith("01"));
    if (exact) return alignSeries(exact.values);
  }
  return sumSeries(rows.map((row) => row.values));
}

function alignSeries(values) {
  const length = state.months.length;
  if (values.length === length) return [...values];
  if (values.length > length) return values.slice(-length);
  return Array(length - values.length).fill(0).concat(values);
}

function sumSeries(seriesList) {
  const length = state.months.length;
  return Array.from({ length }, (_, index) =>
    seriesList.reduce((sum, values) => sum + (alignSeries(values)[index] || 0), 0)
  );
}

function classifyRow(row) {
  const inputs = getInputs();
  const codeGroups = {
    cash: parseCodes(inputs.cashCodes),
    ar: parseCodes(inputs.arCodes),
    bill: parseCodes(inputs.billCodes),
    eReceivable: parseCodes(inputs.eReceivableCodes),
    loan: parseCodes(inputs.loanCodes),
    sales: ["4000", "4111"],
    cost: ["5000", "5200", "5400"],
    fixed: ["6000"]
  };
  const hit = Object.entries(codeGroups).find(([, codes]) => codes.includes(row.code));
  if (hit) return CATEGORY_LABELS[hit[0]];
  if (/借入/.test(row.name)) return CATEGORY_LABELS.loan;
  return CATEGORY_LABELS.other;
}

function seedAssumptions() {
  const inputs = getInputs();
  const cash = pickSeries(parseCodes(inputs.cashCodes));
  const ar = pickSeries(parseCodes(inputs.arCodes), false);
  const bills = pickSeries(parseCodes(inputs.billCodes), false);
  const eReceivables = pickSeries(parseCodes(inputs.eReceivableCodes), false);
  const loans = pickSeries(parseCodes(inputs.loanCodes), false);
  const sales = pickSeries(["4111", "4000"]);
  const costs = pickSeries(["5000", "5200", "5400"]);
  const fixed = pickSeries(["6000"]);

  const latestCash = last(cash);
  const recentSales = average(sales.slice(-3));
  const recentCosts = average(costs.slice(-3));
  const recentFixed = average(fixed.slice(-3));
  const costRate = recentSales > 0 && recentCosts > 0 ? Math.round((recentCosts / recentSales) * 100) : Number(inputs.costRate);
  const recentLoanRepayments = loans.slice(1).map((value, index) => Math.max(0, loans[index] - value));
  const loanRepayment = average(recentLoanRepayments.slice(-3));

  setInputIfEmpty("openingCash", latestCash);
  setInputIfEmpty("salesForecast", recentSales);
  setInputIfEmpty("fixedCosts", recentFixed);
  setInputIfEmpty("loanRepayment", loanRepayment);
  if (Number.isFinite(costRate) && costRate > 0) document.querySelector("#costRate").value = Math.min(Math.round(costRate), 150);

  return { cash, ar, bills, eReceivables, loans, sales };
}

function setInputIfEmpty(id, value) {
  const input = document.querySelector(`#${id}`);
  if (!input.value && Number.isFinite(value)) input.value = Math.round(value);
}

function buildForecast() {
  if (!state.months.length) {
    renderEmptyTable();
    return [];
  }

  const inputs = getInputs();
  const basis = seedAssumptions();
  const months = Array.from({ length: Number(inputs.forecastMonths) || 6 }, (_, index) => nextMonth(last(state.months), index + 1));
  const salesForecast = parseAmount(document.querySelector("#salesForecast").value);
  const costRate = Number(document.querySelector("#costRate").value || 0) / 100;
  const fixedCosts = parseAmount(document.querySelector("#fixedCosts").value);
  const openingCash = parseAmount(document.querySelector("#openingCash").value);
  const baseLoanRepayment = parseAmount(document.querySelector("#loanRepayment").value);

  ensureManualArrays(months.length, baseLoanRepayment);

  const currentPct = Number(inputs.collectCurrent || 0) / 100;
  const nextPct = Number(inputs.collectNext || 0) / 100;
  const afterNextPct = Number(inputs.collectAfterNext || 0) / 100;
  const billMaturity = Math.max(1, Number(inputs.billMaturity || 1));
  const eReceivableMaturity = Math.max(1, Number(inputs.eReceivableMaturity || 1));

  const openingAr = last(basis.ar);
  const openingBills = last(basis.bills);
  const openingEReceivables = last(basis.eReceivables);
  const existingBillCollections = spreadCollection(openingBills, billMaturity, months.length);
  const existingEReceivableCollections = spreadCollection(openingEReceivables, eReceivableMaturity, months.length);

  let cash = openingCash;
  const rows = months.map((month, index) => {
    const currentSalesCollection = salesForecast * currentPct;
    const priorSalesCollection = index >= 1 ? salesForecast * nextPct : openingAr;
    const twoMonthSalesCollection = index >= 2 ? salesForecast * afterNextPct : 0;
    const salesCollection = currentSalesCollection + priorSalesCollection + twoMonthSalesCollection;
    const billCollection = existingBillCollections[index] || 0;
    const eReceivableCollection = existingEReceivableCollections[index] || 0;
    const otherIn = state.manual.otherIn[index] || 0;
    const purchasePayment = salesForecast * costRate;
    const loanRepayment = state.manual.loanRepayment[index] ?? baseLoanRepayment;
    const otherOut = state.manual.otherOut[index] || 0;
    const totalIn = salesCollection + billCollection + eReceivableCollection + otherIn;
    const totalOut = purchasePayment + fixedCosts + loanRepayment + otherOut;
    const opening = cash;
    cash = opening + totalIn - totalOut;

    return {
      month,
      opening,
      salesCollection,
      billCollection,
      eReceivableCollection,
      otherIn,
      totalIn,
      purchasePayment,
      fixedCosts,
      loanRepayment,
      otherOut,
      totalOut,
      ending: cash
    };
  });

  state.forecastRows = rows;
  renderForecast(rows);
  return rows;
}

function spreadCollection(amount, months, length) {
  if (!amount || !months) return Array(length).fill(0);
  const perMonth = amount / months;
  return Array.from({ length }, (_, index) => (index < months ? perMonth : 0));
}

function ensureManualArrays(length, baseLoanRepayment) {
  ["otherIn", "otherOut", "loanRepayment"].forEach((key) => {
    state.manual[key] = Array.from({ length }, (_, index) => {
      const existing = state.manual[key][index];
      if (Number.isFinite(existing)) return existing;
      return key === "loanRepayment" ? baseLoanRepayment : 0;
    });
  });
}

function renderForecast(rows) {
  els.cashflowHead.innerHTML = `<tr><th>項目</th>${rows.map((row) => `<th>${row.month}</th>`).join("")}</tr>`;
  const lines = [
    { label: "月初資金", key: "opening" },
    { label: "入金", section: true },
    { label: "売上回収", key: "salesCollection" },
    { label: "受取手形期日入金", key: "billCollection" },
    { label: "電子記録債権期日入金", key: "eReceivableCollection" },
    { label: "その他入金", key: "otherIn", editable: true },
    { label: "入金計", key: "totalIn" },
    { label: "出金", section: true },
    { label: "仕入・外注支払", key: "purchasePayment" },
    { label: "固定費支払", key: "fixedCosts" },
    { label: "借入返済", key: "loanRepayment", editable: true },
    { label: "その他出金", key: "otherOut", editable: true },
    { label: "出金計", key: "totalOut" },
    { label: "月末資金", key: "ending", cashEnd: true }
  ];

  els.cashflowBody.innerHTML = lines
    .map((line) => {
      if (line.section) return `<tr class="row-section"><td colspan="${rows.length + 1}">${line.label}</td></tr>`;
      const cells = rows
        .map((row, index) => renderForecastCell(line, row, index))
        .join("");
      const className = line.cashEnd ? "cash-end" : "";
      return `<tr class="${className}"><td>${line.label}</td>${cells}</tr>`;
    })
    .join("");

  document.querySelectorAll("[data-manual-key]").forEach((input) => {
    input.addEventListener("change", () => {
      const key = input.dataset.manualKey;
      const index = Number(input.dataset.index);
      state.manual[key][index] = parseAmount(input.value);
      buildForecast();
    });
  });

  els.basisText.textContent = `${last(state.months)}実績を起点に、${rows.length}か月分を作成`;
}

function renderForecastCell(line, row, index) {
  const value = row[line.key] || 0;
  if (line.editable) {
    return `<td class="editable-cell"><input type="number" step="10000" value="${Math.round(value)}" data-manual-key="${line.key}" data-index="${index}"></td>`;
  }
  const className = value < 0 ? "negative" : "";
  return `<td class="${className}">${formatYen(value)}</td>`;
}

function renderEmptyTable() {
  els.cashflowHead.innerHTML = "";
  els.cashflowBody.innerHTML = els.emptyTableTemplate.innerHTML;
}

function renderMetrics() {
  const inputs = getInputs();
  const cash = pickSeries(parseCodes(inputs.cashCodes));
  const ar = pickSeries(parseCodes(inputs.arCodes), false);
  const bills = pickSeries(parseCodes(inputs.billCodes), false);
  const eReceivables = pickSeries(parseCodes(inputs.eReceivableCodes), false);
  const loans = pickSeries(parseCodes(inputs.loanCodes), false);

  els.monthMetric.textContent = last(state.months, "-");
  els.cashMetric.textContent = state.months.length ? `${formatYen(last(cash))}円` : "-";
  els.receivableMetric.textContent = state.months.length ? `${formatYen(last(ar) + last(bills) + last(eReceivables))}円` : "-";
  els.loanMetric.textContent = state.months.length ? `${formatYen(last(loans))}円` : "-";
}

function renderRows() {
  els.rowCount.textContent = `${state.rows.length}件`;
  els.rowsTable.innerHTML = state.rows
    .map(
      (row) => `<tr>
        <td>${row.code}</td>
        <td>${row.name}</td>
        <td>${formatYen(row.latest)}</td>
        <td>${classifyRow(row)}</td>
      </tr>`
    )
    .join("");
}

function acceptParsedData(parsed, sourceLabel) {
  state.months = parsed.months;
  state.rows = parsed.rows;
  state.manual = { otherIn: [], otherOut: [], loanRepayment: [] };

  renderRows();
  renderMetrics();
  buildForecast();

  els.importStatus.textContent = `${sourceLabel}読込`;
  els.importStatus.classList.add("ready");
}

async function loadPdfJs() {
  if (!pdfjsPromise) {
    pdfjsPromise = import(PDFJS_URL).then((pdfjs) => {
      pdfjs.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
      return pdfjs;
    });
  }
  return pdfjsPromise;
}

async function extractPdfText(file) {
  const pdfjs = await loadPdfJs();
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjs.getDocument({ data: buffer }).promise;
  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    pages.push(textItemsToLines(content.items));
  }

  return pages.join("\n");
}

function textItemsToLines(items) {
  const lines = [];
  const sorted = [...items].sort((a, b) => {
    const ay = Math.round(a.transform[5]);
    const by = Math.round(b.transform[5]);
    if (Math.abs(by - ay) > 2) return by - ay;
    return a.transform[4] - b.transform[4];
  });

  sorted.forEach((item) => {
    const y = Math.round(item.transform[5]);
    let line = lines.find((entry) => Math.abs(entry.y - y) <= 2);
    if (!line) {
      line = { y, items: [] };
      lines.push(line);
    }
    line.items.push(item);
  });

  return lines
    .map((line) => line.items.sort((a, b) => a.transform[4] - b.transform[4]).map((item) => item.str).join(" "))
    .join("\n");
}

function sampleText() {
  const months = ["2025/07", "2025/08", "2025/09", "2025/10", "2025/11", "2025/12", "2026/01", "2026/02", "2026/03", "2026/04", "2026/05"];
  const rows = [
    ["1101", "現金及び預金", [5800000, 6420000, 6100000, 7200000, 6650000, 7900000, 7600000, 8120000, 7350000, 6900000, 7450000]],
    ["1121", "売掛金", [9200000, 8700000, 9600000, 10400000, 10100000, 9900000, 11200000, 10800000, 11600000, 12100000, 11800000]],
    ["1122", "受取手形", [2200000, 2100000, 2400000, 2500000, 2300000, 2600000, 2500000, 2400000, 2300000, 2200000, 2100000]],
    ["1124", "電子記録債権", [1400000, 1500000, 1450000, 1600000, 1580000, 1700000, 1650000, 1800000, 1750000, 1680000, 1720000]],
    ["2211", "長期借入金", [9800000, 9600000, 9400000, 9200000, 9000000, 8800000, 8600000, 8400000, 8200000, 8000000, 7800000]],
    ["4111", "売上高", [9800000, 9400000, 10100000, 10700000, 9900000, 11500000, 10300000, 10900000, 11100000, 11800000, 11200000]],
    ["5000", "売上原価", [6100000, 5850000, 6320000, 6600000, 6200000, 7100000, 6500000, 6800000, 6950000, 7350000, 7000000]],
    ["6000", "販売費及び一般管理費", [2700000, 2750000, 2680000, 2800000, 2760000, 2900000, 2820000, 2850000, 2880000, 2920000, 2860000]]
  ];
  return `${months.join(" ")}\n${rows.map(([code, name, values]) => `(${code}) ${name} ${values.map(formatYen).join(" ")}`).join("\n")}`;
}

function exportCsv() {
  if (!state.forecastRows.length) return;
  const header = ["項目", ...state.forecastRows.map((row) => row.month)];
  const lines = [
    ["月初資金", "opening"],
    ["売上回収", "salesCollection"],
    ["受取手形期日入金", "billCollection"],
    ["電子記録債権期日入金", "eReceivableCollection"],
    ["その他入金", "otherIn"],
    ["入金計", "totalIn"],
    ["仕入・外注支払", "purchasePayment"],
    ["固定費支払", "fixedCosts"],
    ["借入返済", "loanRepayment"],
    ["その他出金", "otherOut"],
    ["出金計", "totalOut"],
    ["月末資金", "ending"]
  ];
  const csv = [header, ...lines.map(([label, key]) => [label, ...state.forecastRows.map((row) => Math.round(row[key] || 0))])]
    .map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(","))
    .join("\n");
  const blob = new Blob([`\uFEFF${csv}`], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `資金繰り表_${new Date().toISOString().slice(0, 10)}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

els.pdfInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  els.importStatus.textContent = "読込中";
  els.importStatus.classList.remove("ready");
  try {
    const text = await extractPdfText(file);
    els.rawTextInput.value = text;
    acceptParsedData(parseBalanceText(text), "PDF");
  } catch (error) {
    els.importStatus.textContent = "要確認";
    console.error(error);
    alert("PDF読込に失敗しました。テキスト読込を使ってください。");
  }
});

els.parseTextButton.addEventListener("click", () => {
  acceptParsedData(parseBalanceText(els.rawTextInput.value), "テキスト");
});

els.sampleButton.addEventListener("click", () => {
  els.rawTextInput.value = sampleText();
  acceptParsedData(parseBalanceText(els.rawTextInput.value), "サンプル");
});

els.recalculateButton.addEventListener("click", () => {
  renderRows();
  renderMetrics();
  buildForecast();
});

els.exportButton.addEventListener("click", exportCsv);

inputIds.forEach((id) => {
  document.querySelector(`#${id}`).addEventListener("change", () => {
    renderRows();
    renderMetrics();
    buildForecast();
  });
});

renderEmptyTable();
