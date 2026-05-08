const PDFJS_SOURCES = [
  {
    module: "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.mjs",
    worker: "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.mjs",
    cMap: "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/cmaps/",
    fonts: "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/standard_fonts/"
  },
  {
    module: "https://unpkg.com/pdfjs-dist@4.10.38/build/pdf.mjs",
    worker: "https://unpkg.com/pdfjs-dist@4.10.38/build/pdf.worker.mjs",
    cMap: "https://unpkg.com/pdfjs-dist@4.10.38/cmaps/",
    fonts: "https://unpkg.com/pdfjs-dist@4.10.38/standard_fonts/"
  }
];

const ACCOUNT_NAMES = {
  "1000": "資産合計",
  "1100": "流動資産",
  "1101": "現金及び預金",
  "1111": "現金",
  "1113": "普通預金",
  "1114": "定期預金",
  "1115": "その他預金",
  "1120": "売上債権",
  "1121": "受取手形",
  "1122": "売掛金",
  "1124": "電子記録債権",
  "2000": "負債・純資産合計",
  "2100": "流動負債",
  "2112": "買掛金",
  "2113": "支払手形",
  "2114": "電子記録債務",
  "2117": "預り金",
  "2125": "未払法人税等",
  "2200": "固定負債",
  "2211": "社債",
  "2212": "長期借入金",
  "2213": "役員借入金",
  "3000": "負債合計",
  "3100": "純資産",
  "4000": "売上高合計",
  "4111": "売上高",
  "4115": "売上値引戻り高",
  "5000": "売上原価",
  "5100": "期首棚卸高",
  "5200": "当期売上原価",
  "5400": "製造原価",
  "6100": "販売費及び一般管理費計",
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
  dropZone: document.querySelector(".drop-zone"),
  importMessage: document.querySelector("#importMessage"),
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
  evidenceSummary: document.querySelector("#evidenceSummary"),
  presetEvidenceList: document.querySelector("#presetEvidenceList"),
  billDetailsBody: document.querySelector("#billDetailsBody"),
  eReceivableDetailsBody: document.querySelector("#eReceivableDetailsBody"),
  addBillDetail: document.querySelector("#addBillDetail"),
  addEReceivableDetail: document.querySelector("#addEReceivableDetail"),
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
  "collectBill",
  "collectEReceivable",
  "billMaturity",
  "eReceivableMaturity",
  "costRate",
  "fixedCosts",
  "loanRepayment",
  "openingCash",
  "forecastMonths"
];

const moneyInputIds = ["salesForecast", "fixedCosts", "loanRepayment", "openingCash"];

const state = {
  months: [],
  rows: [],
  forecastRows: [],
  sources: [],
  presetEvidence: [],
  billDetails: [],
  eReceivableDetails: [],
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

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function parseAmount(value) {
  if (value == null || value === "") return 0;
  const text = String(value).replace(/[,\s円]/g, "").replace(/△/g, "-").replace(/▲/g, "-");
  return Number(text) || 0;
}

function formatMoneyInput(input) {
  const hadMinus = /^-/.test(String(input.value).trim());
  const digits = String(input.value).replace(/[^\d]/g, "");
  if (!digits) {
    input.value = "";
    return 0;
  }
  const value = Number(`${hadMinus ? "-" : ""}${digits}`);
  input.value = formatYen(value);
  return value;
}

function setMoneyInputValue(id, value) {
  const input = document.querySelector(`#${id}`);
  if (!input || !Number.isFinite(value)) return;
  input.value = formatYen(value);
}

function getCollectionMode() {
  return document.querySelector('input[name="collectionMode"]:checked')?.value || "simple";
}

function createReceivableDetail(overrides = {}) {
  return {
    id: `detail-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    amount: 0,
    issueDate: "",
    dueDate: "",
    memo: "",
    ...overrides
  };
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
    .replace(/([△▲-]?\d{1,4}\.\d)\s*(\d{1,3},\d{3}(?:,\d{3})*)/g, "$1 $2")
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
        latest: last(values, 0),
        statement: inferStatementKind(segment.code, extractLabel(segment.raw, segment.code)),
        rowType: inferRowType(segment.code, extractLabel(segment.raw, segment.code)),
        sourceName: "テキスト"
      };
    })
    .filter((row) => row.values.length >= Math.min(3, monthCount));

  return { months, rows: dedupeRows(rows) };
}

function decodeCsvBuffer(buffer) {
  const utf8 = new TextDecoder("utf-8", { fatal: false }).decode(buffer);
  if (!utf8.includes("�")) return utf8;
  try {
    return new TextDecoder("shift_jis", { fatal: false }).decode(buffer);
  } catch {
    return utf8;
  }
}

function parseCsvText(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (char === '"' && inQuotes && next === '"') {
      cell += '"';
      index += 1;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === "," && !inQuotes) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell);
      if (row.some((value) => value !== "")) rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  if (cell || row.length) {
    row.push(cell);
    if (row.some((value) => value !== "")) rows.push(row);
  }

  return rows;
}

function parseBalanceCsvText(rawText, sourceName = "CSV") {
  const table = parseCsvText(rawText.replace(/^\uFEFF/, ""));
  const headerIndex = table.findIndex((row) => row.some((cell) => cell.trim() === "勘定科目コード"));
  if (headerIndex < 0) throw new Error(`${sourceName} に「勘定科目コード」の列が見つかりません。`);

  const header = table[headerIndex].map((cell) => cell.trim());
  const codeIndex = header.findIndex((cell) => cell === "勘定科目コード");
  const nameIndex = header.findIndex((cell) => cell === "勘定科目名");
  const monthColumns = header
    .map((cell, index) => ({ cell, index }))
    .filter(({ cell }) => /^20\d{2}\/(0?[1-9]|1[0-2])$/.test(cell))
    .map(({ cell, index }) => ({ month: normalizeMonthCell(cell), index }));

  if (codeIndex < 0 || nameIndex < 0 || monthColumns.length < 3) {
    throw new Error(`${sourceName} の列構成を確認してください。月次列が不足しています。`);
  }

  const months = monthColumns.map((column) => column.month);
  const rows = table.slice(headerIndex + 1)
    .map((line) => {
      const code = String(line[codeIndex] || "").trim();
      const name = String(line[nameIndex] || "").trim();
      if (!/^\d{4}$/.test(code) || !name) return null;
      const values = monthColumns.map((column) => parseAmount(line[column.index] || 0));
      const statement = inferStatementKind(code, name);
      const rowType = inferRowType(code, name);
      return {
        code,
        name,
        values,
        latest: last(values, 0),
        statement,
        rowType,
        sourceName
      };
    })
    .filter(Boolean);

  return {
    months,
    rows,
    source: {
      name: sourceName,
      kind: summarizeKinds(rows),
      rowCount: rows.length
    }
  };
}

function normalizeMonthCell(cell) {
  const match = String(cell).match(/^(20\d{2})\/(0?[1-9]|1[0-2])$/);
  return match ? monthKey(Number(match[1]), Number(match[2])) : cell;
}

function inferStatementKind(code, name = "") {
  if (/CR|ＣＲ|資金繰|資金収支|現金収支|キャッシュ|収支/.test(name)) return "CR";
  const numericCode = Number(code);
  if (numericCode >= 5400 && numericCode <= 5500) return "CR";
  if (numericCode >= 4000 && numericCode < 9000) return "P/L";
  if (numericCode >= 1000 && numericCode < 4000) return "B/S";
  if (numericCode >= 9000) return "B/S";
  return "その他";
}

function inferRowType(code, name = "") {
  const aggregatePatterns = /(計|合計|小計|総利益|営業利益|経常利益|当期.*利益|税引|純売上高|負債・純資産|株主資本|当座資産|棚卸資産|流動|固定|純資産の部)/;
  return aggregatePatterns.test(name) || ["4000", "5000", "5200", "5400", "5500", "6000", "7000", "7200", "8000", "9000"].includes(code)
    ? "集計行"
    : "科目残高";
}

function summarizeKinds(rows) {
  const kinds = [...new Set(rows.map((row) => row.statement).filter(Boolean))];
  return kinds.join("・") || "CSV";
}

function mergeParsedSources(parsedSources) {
  const validSources = parsedSources.filter((source) => source && source.rows.length && source.months.length);
  if (!validSources.length) return { months: [], rows: [], sources: [] };

  const months = chooseMonths(validSources.map((source) => source.months));
  const rows = validSources.flatMap((source) =>
    source.rows.map((row) => ({
      ...row,
      values: alignValuesToMonths(source.months, row.values, months),
      latest: last(alignValuesToMonths(source.months, row.values, months), 0),
      sourceName: row.sourceName || source.source?.name || "取込データ",
      statement: row.statement || inferStatementKind(row.code, row.name),
      rowType: row.rowType || inferRowType(row.code, row.name)
    }))
  );

  return {
    months,
    rows: dedupeRows(rows),
    sources: uniqueSources(validSources.map((source) => source.source).filter(Boolean))
  };
}

function chooseMonths(monthGroups) {
  return [...monthGroups].sort((a, b) => b.length - a.length)[0] || [];
}

function alignValuesToMonths(sourceMonths, values, targetMonths) {
  return targetMonths.map((month) => {
    const index = sourceMonths.indexOf(month);
    return index >= 0 ? values[index] || 0 : 0;
  });
}

function uniqueSources(sources) {
  const seen = new Set();
  return sources.filter((source) => {
    const key = `${source.kind}:${source.rowCount}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function parseBalancePdfPages(pageItems) {
  const pageModels = pageItems.map((items, index) => {
    const lines = groupPdfLines(items, index + 1);
    return {
      pageNumber: index + 1,
      lines,
      monthAnchors: extractMonthAnchorsFromLines(lines),
      text: lines.map((line) => line.text).join("\n")
    };
  });
  const fallbackText = pageModels.map((page) => page.text).join("\n");
  const bestAnchors = chooseBestMonthAnchors(pageModels.map((page) => page.monthAnchors));

  if (bestAnchors.length < 3) {
    return { months: [], rows: [], text: fallbackText };
  }

  const months = bestAnchors.map((anchor) => anchor.month);
  const rows = [];
  pageModels.forEach((page) => {
    const anchors = page.monthAnchors.length >= 3 ? page.monthAnchors : bestAnchors;
    page.lines.forEach((line) => {
      const row = parsePdfAccountLine(line, anchors);
      if (row) rows.push({ ...row, sourceName: `PDF ${page.pageNumber}頁` });
    });
  });

  return {
    months,
    rows: dedupeRows(rows).filter((row) => row.values.length >= Math.min(3, months.length)),
    text: fallbackText,
    source: {
      name: "PDF",
      kind: summarizeKinds(rows),
      rowCount: rows.length
    }
  };
}

function groupPdfLines(items, pageNumber) {
  const normalized = items
    .filter((item) => item && typeof item.str === "string" && item.str.trim())
    .map((item) => {
      const x = item.transform[4];
      const y = item.transform[5];
      const width = item.width || Math.max(item.str.length * 4, 1);
      return {
        str: item.str,
        x,
        y,
        x2: x + width,
        width,
        pageNumber
      };
    })
    .sort((a, b) => {
      if (Math.abs(b.y - a.y) > 2) return b.y - a.y;
      return a.x - b.x;
    });

  const lines = [];
  normalized.forEach((item) => {
    let line = lines.find((entry) => Math.abs(entry.y - item.y) <= 2.5);
    if (!line) {
      line = { pageNumber, y: item.y, items: [] };
      lines.push(line);
    }
    line.items.push(item);
  });

  return lines.map((line) => {
    const ordered = line.items.sort((a, b) => a.x - b.x);
    return {
      ...line,
      items: ordered,
      text: buildPdfLineText(ordered)
    };
  });
}

function buildPdfLineText(items) {
  return items.reduce((text, item, index) => {
    const prev = items[index - 1];
    const gap = prev ? item.x - prev.x2 : 0;
    const spacer = gap > 3 ? " " : "";
    return `${text}${spacer}${item.str}`;
  }, "");
}

function chunkPdfLine(items, gapLimit = 5) {
  const chunks = [];
  items.forEach((item) => {
    const prev = chunks[chunks.length - 1];
    if (!prev || item.x - prev.x2 > gapLimit) {
      chunks.push({
        text: item.str,
        x: item.x,
        x2: item.x2
      });
    } else {
      prev.text += item.str;
      prev.x2 = Math.max(prev.x2, item.x2);
    }
  });
  return chunks;
}

function extractMonthAnchorsFromLines(lines) {
  let best = [];
  lines.forEach((line) => {
    const anchors = [];
    chunkPdfLine(line.items, 18).forEach((chunk) => {
      const compact = chunk.text.replace(/\s+/g, "");
      const matches = [...compact.matchAll(/(20\d{2})(?:[\/年.-]?)(0?[1-9]|1[0-2])月?/g)];
      matches.forEach((match) => {
        const ratio = compact.length ? match.index / compact.length : 0.5;
        anchors.push({
          month: monthKey(Number(match[1]), Number(match[2])),
          x: chunk.x + (chunk.x2 - chunk.x) * ratio + (chunk.x2 - chunk.x) / Math.max(matches.length * 2, 2)
        });
      });
    });
    const unique = dedupeMonthAnchors(anchors);
    if (unique.length > best.length) best = unique;
  });
  return best.sort((a, b) => a.x - b.x);
}

function dedupeMonthAnchors(anchors) {
  const byMonth = new Map();
  anchors.forEach((anchor) => {
    if (!byMonth.has(anchor.month)) byMonth.set(anchor.month, anchor);
  });
  return [...byMonth.values()].sort((a, b) => a.x - b.x);
}

function chooseBestMonthAnchors(anchorGroups) {
  const groups = anchorGroups
    .map(dedupeMonthAnchors)
    .filter((anchors) => anchors.length >= 3)
    .sort((a, b) => b.length - a.length);
  return groups[0] || [];
}

function parsePdfAccountLine(line, monthAnchors) {
  const codeMatch = line.text.match(/\((?:△|▲)?\s*(\d{4})\)/);
  if (!codeMatch) return null;
  const code = codeMatch[1];
  const chunks = chunkPdfLine(line.items, 4);
  const codeChunk = chunks.find((chunk) => chunk.text.includes(code));
  const codeX = codeChunk ? codeChunk.x2 : 0;
  const values = extractValuesByMonthColumns(chunks, monthAnchors, codeX);
  const numericChunks = chunks.filter((chunk) => chunk.x2 > codeX && /\d/.test(chunk.text));
  const fallbackValues = numericChunks.length >= monthAnchors.length && values.every((value) => Number.isFinite(value))
    ? values
    : extractAmounts(line.text, monthAnchors.length);

  if (fallbackValues.length !== monthAnchors.length) return null;

  const name = extractPdfAccountName(line.text, code);
  return {
    code,
    name,
    values: fallbackValues,
    latest: last(fallbackValues, 0),
    statement: inferStatementKind(code, name),
    rowType: inferRowType(code, name)
  };
}

function extractPdfAccountName(text, code) {
  const afterCode = text.replace(new RegExp(`^.*?\\(${code}\\)`), "");
  const beforeValue = afterCode.split(/[△▲-]?\d{1,4}\.\d|[△▲-]?\d{1,3}(?:,\d{3})+/)[0] || "";
  const cleaned = beforeValue.replace(/[0-9.,%()\/:*-]/g, "").replace(/\s+/g, " ").trim();
  return ACCOUNT_NAMES[code] || cleaned || `科目 ${code}`;
}

function extractValuesByMonthColumns(chunks, monthAnchors, codeX) {
  return monthAnchors.map((anchor, index) => {
    const prev = monthAnchors[index - 1];
    const next = monthAnchors[index + 1];
    const left = prev ? (prev.x + anchor.x) / 2 : anchor.x - ((next ? next.x - anchor.x : 48) / 2);
    const right = next ? (anchor.x + next.x) / 2 : anchor.x + ((prev ? anchor.x - prev.x : 48) / 2);
    const cellText = chunks
      .filter((chunk) => chunk.x2 > codeX && chunk.x < right && chunk.x2 > left)
      .map((chunk) => chunk.text)
      .join(" ");
    return parseBalanceCellValue(cellText);
  });
}

function parseBalanceCellValue(text) {
  const normalized = String(text || "")
    .replace(/([△▲-]?\d{1,4}\.\d)\s*(\d{1,3},\d{3}(?:,\d{3})*)/g, "$1 $2")
    .replace(/[−－]/g, "-");
  const amountMatches = [...normalized.matchAll(/[△▲-]?\d{1,3}(?:,\d{3})+/g)];
  if (amountMatches.length) return parseAmount(amountMatches[amountMatches.length - 1][0]);

  const withoutRates = normalized.replace(/[△▲-]?\d+\.\d+/g, " ");
  const integerMatches = [...withoutRates.matchAll(/[△▲-]?\d+/g)];
  if (!integerMatches.length) return NaN;
  return parseAmount(integerMatches[integerMatches.length - 1][0]);
}

function dedupeRows(rows) {
  const byCode = new Map();
  rows.forEach((row) => {
    const key = `${row.statement || inferStatementKind(row.code, row.name)}:${row.code}`;
    const current = byCode.get(key);
    if (!current || row.values.length > current.values.length) byCode.set(key, row);
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
  return pickSeriesInfo(codes, preferAggregate).values;
}

function pickSeriesInfo(codes, preferAggregate = true) {
  const rows = findRows(codes);
  if (!rows.length) return { values: Array(state.months.length).fill(0), rows: [] };
  const orderedRows = codes.flatMap((code) => rows.filter((row) => row.code === code));
  if (preferAggregate) {
    const aggregate = orderedRows.find((row) => row.rowType === "集計行") || orderedRows[0];
    if (aggregate) return { values: alignSeries(aggregate.values), rows: [aggregate] };
  }
  return { values: sumSeries(orderedRows.map((row) => row.values)), rows: orderedRows };
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

function getBasisFromRows() {
  const inputs = getInputs();
  const cashInfo = pickSeriesInfo(parseCodes(inputs.cashCodes));
  const arInfo = pickSeriesInfo(parseCodes(inputs.arCodes), false);
  const billsInfo = pickSeriesInfo(parseCodes(inputs.billCodes), false);
  const eReceivablesInfo = pickSeriesInfo(parseCodes(inputs.eReceivableCodes), false);
  const loansInfo = pickSeriesInfo(parseCodes(inputs.loanCodes), false);
  const salesInfo = pickSeriesInfo(["4000", "4111"]);
  const costsInfo = pickSeriesInfo(["5200", "5100", "5500", "5400"]);
  const fixedInfo = pickSeriesInfo(["6100", "6000"]);

  return {
    cash: cashInfo.values,
    ar: arInfo.values,
    bills: billsInfo.values,
    eReceivables: eReceivablesInfo.values,
    loans: loansInfo.values,
    sales: salesInfo.values,
    costs: costsInfo.values,
    fixed: fixedInfo.values,
    cashInfo,
    salesInfo,
    costsInfo,
    fixedInfo,
    loansInfo
  };
}

function calculatePdfDefaults(basis) {
  const latestCash = seriesHasData(basis.cash) ? last(basis.cash) : NaN;
  const recentSales = seriesHasData(basis.sales) ? recentFlowAverage(basis.sales) : NaN;
  const recentCosts = seriesHasData(basis.costs) ? recentFlowAverage(basis.costs) : NaN;
  const recentFixed = seriesHasData(basis.fixed) ? recentFlowAverage(basis.fixed) : NaN;
  const costRate = recentSales > 0 && recentCosts > 0 ? Math.round((recentCosts / recentSales) * 100) : Number(document.querySelector("#costRate").value || 0);
  const recentLoanRepayments = basis.loans.slice(1).map((value, index) => Math.max(0, basis.loans[index] - value));
  const loanRepayment = recentLoanRepayments.some((value) => value > 0) ? average(recentLoanRepayments.slice(-3)) : NaN;

  return {
    openingCash: latestCash,
    salesForecast: recentSales,
    fixedCosts: recentFixed,
    loanRepayment,
    costRate
  };
}

function seriesHasData(values) {
  return values.some((value) => Number.isFinite(value) && value !== 0);
}

function recentFlowAverage(values) {
  const nums = values.filter((value) => Number.isFinite(value));
  if (nums.length < 2) return last(nums, NaN);
  const diffs = nums.slice(1).map((value, index) => value - nums[index]);
  const nonDecreasingRatio = diffs.filter((value) => value >= 0).length / diffs.length;
  const recentDiffs = diffs.slice(-3).filter((value) => value >= 0);
  if (nonDecreasingRatio >= 0.75 && recentDiffs.length) {
    return average(recentDiffs);
  }
  return average(nums.slice(-3));
}

function applyPdfDefaults(basis, force = false) {
  const defaults = calculatePdfDefaults(basis);
  state.presetEvidence = buildPresetEvidence(basis, defaults);
  setInputIfEmpty("openingCash", defaults.openingCash, force);
  setInputIfEmpty("salesForecast", defaults.salesForecast, force);
  setInputIfEmpty("fixedCosts", defaults.fixedCosts, force);
  setInputIfEmpty("loanRepayment", defaults.loanRepayment, force);
  if ((force || !document.querySelector("#costRate").value) && Number.isFinite(defaults.costRate) && defaults.costRate > 0) {
    document.querySelector("#costRate").value = Math.min(Math.round(defaults.costRate), 150);
  }
  renderPresetEvidence();
}

function buildPresetEvidence(basis, defaults) {
  const lastMonth = last(state.months, "-");
  const recentMonths = state.months.slice(-3);
  const evidence = [];
  if (Number.isFinite(defaults.openingCash)) {
    evidence.push({
      label: "開始資金",
      value: defaults.openingCash,
      method: `最終月 ${lastMonth} の現預金残高`,
      source: describeSourceRows(basis.cashInfo.rows),
      values: [{ month: lastMonth, value: defaults.openingCash }]
    });
  }
  if (Number.isFinite(defaults.salesForecast)) {
    evidence.push({
      label: "月商見込",
      value: defaults.salesForecast,
      method: "売上の直近3か月平均",
      source: describeSourceRows(basis.salesInfo.rows),
      values: recentValuePairs(basis.sales, recentMonths)
    });
  }
  if (Number.isFinite(defaults.fixedCosts)) {
    evidence.push({
      label: "固定費/月",
      value: defaults.fixedCosts,
      method: "販売費及び一般管理費の直近3か月平均",
      source: describeSourceRows(basis.fixedInfo.rows),
      values: recentValuePairs(basis.fixed, recentMonths)
    });
  }
  if (Number.isFinite(defaults.loanRepayment)) {
    evidence.push({
      label: "借入返済/月",
      value: defaults.loanRepayment,
      method: "借入金残高の減少額の直近3か月平均",
      source: describeSourceRows(basis.loansInfo.rows),
      values: recentLoanDecreasePairs(basis.loans, recentMonths)
    });
  }
  if (Number.isFinite(defaults.costRate) && defaults.costRate > 0) {
    evidence.push({
      label: "仕入・外注率",
      value: defaults.costRate,
      suffix: "%",
      method: "売上原価 ÷ 売上",
      source: `${describeSourceRows(basis.costsInfo.rows)} / ${describeSourceRows(basis.salesInfo.rows)}`,
      values: [
        { month: "売上原価平均", value: average(basis.costs.slice(-3)) },
        { month: "売上平均", value: average(basis.sales.slice(-3)) }
      ]
    });
  }
  return evidence;
}

function describeSourceRows(rows) {
  if (!rows.length) return "該当科目なし";
  return rows.map((row) => `${row.statement} ${row.code} ${row.name}（${row.sourceName || "取込データ"}）`).join(" + ");
}

function recentValuePairs(values, months) {
  return months.map((month, index) => ({
    month,
    value: values[values.length - months.length + index] || 0
  }));
}

function recentLoanDecreasePairs(values, months) {
  const decreases = values.slice(1).map((value, index) => Math.max(0, values[index] - value));
  const recent = decreases.slice(-3);
  return months.slice(-recent.length).map((month, index) => ({ month, value: recent[index] || 0 }));
}

function renderPresetEvidence() {
  if (!state.presetEvidence.length) {
    els.evidenceSummary.textContent = "根拠に使える読取値がありません";
    els.presetEvidenceList.innerHTML = "";
    return;
  }
  els.evidenceSummary.textContent = `${state.presetEvidence.length}項目を自動設定`;
  els.presetEvidenceList.innerHTML = state.presetEvidence.map((item) => {
    const valueText = item.suffix === "%" ? `${formatYen(item.value)}%` : `${formatYen(item.value)}円`;
    const valueLines = item.values
      .map((entry) => `${entry.month}: ${formatYen(entry.value)}${item.suffix === "%" ? "" : "円"}`)
      .join(" / ");
    return `<div class="evidence-item">
      <strong>${escapeHtml(item.label)}: ${escapeHtml(valueText)}</strong>
      <p>根拠: ${escapeHtml(item.method)}</p>
      <p>参照: ${escapeHtml(item.source)}</p>
      <p>値: ${escapeHtml(valueLines)}</p>
    </div>`;
  }).join("");
}

function setInputIfEmpty(id, value, force = false) {
  const input = document.querySelector(`#${id}`);
  if (force && moneyInputIds.includes(id) && !Number.isFinite(value)) {
    input.value = "";
    return;
  }
  if ((!input.value || force) && Number.isFinite(value)) {
    if (moneyInputIds.includes(id)) {
      setMoneyInputValue(id, Math.round(value));
    } else {
      input.value = Math.round(value);
    }
  }
}

function buildDueDateCollections(details, months) {
  return months.map((month) =>
    details.reduce((sum, detail) => {
      if (!detail.dueDate || !detail.amount) return sum;
      return dateToMonthKey(detail.dueDate) === month ? sum + Number(detail.amount || 0) : sum;
    }, 0)
  );
}

function dateToMonthKey(dateValue) {
  const match = String(dateValue || "").match(/^(\d{4})-(\d{2})-\d{2}$/);
  return match ? `${match[1]}/${match[2]}` : "";
}

function buildForecast() {
  if (!state.months.length) {
    renderEmptyTable();
    return [];
  }

  const basis = getBasisFromRows();
  applyPdfDefaults(basis, false);
  const inputs = getInputs();
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
  const billPct = Number(inputs.collectBill || 0) / 100;
  const eReceivablePct = Number(inputs.collectEReceivable || 0) / 100;
  const billMaturity = Math.max(1, Number(inputs.billMaturity || 1));
  const eReceivableMaturity = Math.max(1, Number(inputs.eReceivableMaturity || 1));
  const collectionMode = getCollectionMode();

  const openingAr = last(basis.ar);
  const openingBills = last(basis.bills);
  const openingEReceivables = last(basis.eReceivables);
  const existingBillCollections = collectionMode === "detail"
    ? buildDueDateCollections(state.billDetails, months)
    : spreadCollection(openingBills, billMaturity, months.length);
  const existingEReceivableCollections = collectionMode === "detail"
    ? buildDueDateCollections(state.eReceivableDetails, months)
    : spreadCollection(openingEReceivables, eReceivableMaturity, months.length);

  let cash = openingCash;
  const rows = months.map((month, index) => {
    const currentSalesCollection = salesForecast * currentPct;
    const priorSalesCollection = index >= 1 ? salesForecast * nextPct : openingAr;
    const twoMonthSalesCollection = index >= 2 ? salesForecast * afterNextPct : 0;
    const salesCollection = currentSalesCollection + priorSalesCollection + twoMonthSalesCollection;
    const futureBillCollection = index >= billMaturity ? salesForecast * billPct : 0;
    const futureEReceivableCollection = index >= eReceivableMaturity ? salesForecast * eReceivablePct : 0;
    const billCollection = (existingBillCollections[index] || 0) + futureBillCollection;
    const eReceivableCollection = (existingEReceivableCollections[index] || 0) + futureEReceivableCollection;
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
    input.addEventListener("input", () => {
      const key = input.dataset.manualKey;
      const index = Number(input.dataset.index);
      state.manual[key][index] = formatMoneyInput(input);
    });
    input.addEventListener("change", () => {
      buildForecast();
    });
  });

  const modeText = getCollectionMode() === "detail" ? "手形・電債は明細期日で反映" : "手形・電債は簡易配分で反映";
  els.basisText.textContent = `${last(state.months)}実績を起点に、${rows.length}か月分を作成（${modeText}）`;
}

function renderForecastCell(line, row, index) {
  const value = row[line.key] || 0;
  if (line.editable) {
    return `<td class="editable-cell"><input class="money-input" type="text" inputmode="numeric" value="${formatYen(value)}" data-manual-key="${line.key}" data-index="${index}"></td>`;
  }
  const className = value < 0 ? "negative" : "";
  return `<td class="${className}">${formatYen(value)}</td>`;
}

function renderReceivableDetails() {
  renderDetailRows("bill", state.billDetails, els.billDetailsBody);
  renderDetailRows("eReceivable", state.eReceivableDetails, els.eReceivableDetailsBody);
  bindDetailInputs();
}

function renderDetailRows(kind, details, tbody) {
  if (!tbody) return;
  tbody.innerHTML = details.map((detail) => `<tr>
    <td><input class="money-input detail-amount" type="text" inputmode="numeric" value="${detail.amount ? formatYen(detail.amount) : ""}" data-detail-kind="${kind}" data-detail-id="${detail.id}" data-detail-field="amount"></td>
    <td><input type="date" value="${escapeHtml(detail.issueDate)}" data-detail-kind="${kind}" data-detail-id="${detail.id}" data-detail-field="issueDate"></td>
    <td><input type="date" value="${escapeHtml(detail.dueDate)}" data-detail-kind="${kind}" data-detail-id="${detail.id}" data-detail-field="dueDate"></td>
    <td><input class="detail-memo" type="text" value="${escapeHtml(detail.memo)}" data-detail-kind="${kind}" data-detail-id="${detail.id}" data-detail-field="memo"></td>
    <td><button class="remove-detail" type="button" data-remove-detail-kind="${kind}" data-remove-detail-id="${detail.id}">削除</button></td>
  </tr>`).join("");
}

function bindDetailInputs() {
  document.querySelectorAll("[data-detail-field]").forEach((input) => {
    input.addEventListener("input", () => {
      updateDetailFromInput(input);
      if (input.dataset.detailField === "amount") formatMoneyInput(input);
      if (getCollectionMode() === "detail") buildForecast();
    });
    input.addEventListener("change", () => {
      updateDetailFromInput(input);
      buildForecast();
    });
  });
  document.querySelectorAll("[data-remove-detail-id]").forEach((button) => {
    button.addEventListener("click", () => {
      removeReceivableDetail(button.dataset.removeDetailKind, button.dataset.removeDetailId);
    });
  });
}

function updateDetailFromInput(input) {
  const detail = findReceivableDetail(input.dataset.detailKind, input.dataset.detailId);
  if (!detail) return;
  const field = input.dataset.detailField;
  detail[field] = field === "amount" ? parseAmount(input.value) : input.value;
}

function findReceivableDetail(kind, id) {
  return getReceivableDetailList(kind).find((detail) => detail.id === id);
}

function getReceivableDetailList(kind) {
  return kind === "bill" ? state.billDetails : state.eReceivableDetails;
}

function addReceivableDetail(kind) {
  getReceivableDetailList(kind).push(createReceivableDetail());
  renderReceivableDetails();
}

function removeReceivableDetail(kind, id) {
  const list = getReceivableDetailList(kind);
  const index = list.findIndex((detail) => detail.id === id);
  if (index >= 0) list.splice(index, 1);
  renderReceivableDetails();
  buildForecast();
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
      (row) => `<tr class="${row.rowType === "集計行" ? "aggregate-row" : ""}">
        <td>${escapeHtml(row.code)}</td>
        <td>${escapeHtml(row.name)}</td>
        <td>${formatYen(row.latest)}</td>
        <td>${escapeHtml(row.statement || "-")}</td>
        <td>${escapeHtml(row.rowType || inferRowType(row.code, row.name))}</td>
        <td>${escapeHtml(classifyRow(row))}</td>
      </tr>`
    )
    .join("");
}

function acceptParsedData(parsed, sourceLabel) {
  if (!parsed.rows.length || !parsed.months.length) {
    const rowText = parsed.rows.length ? `${parsed.rows.length}行` : "科目行なし";
    const monthText = parsed.months.length ? `${parsed.months.length}か月` : "月列なし";
    throw new Error(`読み取り結果が不足しています（${rowText}、${monthText}）。`);
  }

  state.months = parsed.months;
  state.rows = parsed.rows;
  state.sources = parsed.sources || (parsed.source ? [parsed.source] : []);
  state.manual = { otherIn: [], otherOut: [], loanRepayment: [] };

  applyPdfDefaults(getBasisFromRows(), true);
  renderRows();
  renderMetrics();
  buildForecast();

  els.importStatus.textContent = `${sourceLabel}読込`;
  els.importStatus.classList.add("ready");
  const sourceText = state.sources.length
    ? `（${state.sources.map((source) => `${source.kind}: ${source.rowCount}行`).join(" / ")}）`
    : "";
  setImportMessage(`${parsed.months.length}か月、${parsed.rows.length}科目を読み取りました。${sourceText}`, false);
}

async function loadPdfJs() {
  if (!pdfjsPromise) {
    pdfjsPromise = (async () => {
      let lastError;
      for (const source of PDFJS_SOURCES) {
        try {
          const pdfjs = await import(source.module);
          pdfjs.GlobalWorkerOptions.workerSrc = source.worker;
          return { pdfjs, source };
        } catch (error) {
          lastError = error;
        }
      }
      throw new Error(`PDF.jsを読み込めませんでした。ネットワーク接続を確認してください。${lastError ? ` (${lastError.message})` : ""}`);
    })().catch((error) => {
      pdfjsPromise = null;
      throw error;
    });
  }
  return pdfjsPromise;
}

async function extractPdfData(file) {
  const { pdfjs, source } = await loadPdfJs();
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjs.getDocument({
    data: buffer,
    cMapUrl: source.cMap,
    cMapPacked: true,
    standardFontDataUrl: source.fonts,
    useSystemFonts: true
  }).promise;
  const pages = [];
  const pageItems = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent({ includeMarkedContent: false });
    pageItems.push(content.items);
    pages.push(textItemsToLines(content.items));
  }

  const text = pages.join("\n");
  const parsed = parseBalancePdfPages(pageItems);
  return {
    text: parsed.text || text,
    parsed
  };
}

function textItemsToLines(items) {
  const lines = [];
  const usableItems = items.filter((item) => item && typeof item.str === "string" && item.str.trim());
  const sorted = [...usableItems].sort((a, b) => {
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
    .map((line) => {
      const ordered = line.items.sort((a, b) => a.transform[4] - b.transform[4]);
      return ordered.reduce((text, item, index) => {
        const prev = ordered[index - 1];
        const gap = prev ? item.transform[4] - (prev.transform[4] + (prev.width || 0)) : 0;
        const spacer = gap > 4 ? " " : "";
        return `${text}${spacer}${item.str}`;
      }, "");
    })
    .join("\n");
}

function setImportMessage(message, isError = false) {
  els.importMessage.textContent = message;
  els.importMessage.classList.toggle("error", isError);
}

async function processInputFiles(files) {
  const fileList = [...files].filter(Boolean);
  if (!fileList.length) return;

  els.importStatus.textContent = "読込中";
  els.importStatus.classList.remove("ready");
  setImportMessage("PDF/CSVから月次残高を読み取っています。", false);

  try {
    const parsedSources = [];
    const textPreviews = [];
    for (const file of fileList) {
      if (isPdfFile(file)) {
        const result = await extractPdfData(file);
        textPreviews.push(`--- ${file.name} ---\n${result.text}`);
        parsedSources.push({ ...result.parsed, source: { ...(result.parsed.source || {}), name: file.name } });
      } else if (isCsvFile(file)) {
        const text = decodeCsvBuffer(await file.arrayBuffer());
        textPreviews.push(`--- ${file.name} ---\n${text}`);
        parsedSources.push(parseBalanceCsvText(text, file.name));
      }
    }

    if (!parsedSources.length) {
      throw new Error("PDFまたはCSVファイルを指定してください。");
    }

    els.rawTextInput.value = textPreviews.join("\n\n");
    acceptParsedData(mergeParsedSources(parsedSources), fileList.some(isCsvFile) ? "CSV" : "PDF");
  } catch (error) {
    els.importStatus.textContent = "要確認";
    els.importStatus.classList.remove("ready");
    console.error(error);
    setImportMessage(error.message || "読込に失敗しました。ファイル形式を確認してください。", true);
  }
}

function isPdfFile(file) {
  return file.type === "application/pdf" || /\.pdf$/i.test(file.name);
}

function isCsvFile(file) {
  return file.type === "text/csv" || /\.csv$/i.test(file.name);
}

function sampleText() {
  const months = ["2025/07", "2025/08", "2025/09", "2025/10", "2025/11", "2025/12", "2026/01", "2026/02", "2026/03", "2026/04", "2026/05"];
  const rows = [
    ["1101", "現金及び預金", [5800000, 6420000, 6100000, 7200000, 6650000, 7900000, 7600000, 8120000, 7350000, 6900000, 7450000]],
    ["1121", "受取手形", [2200000, 2100000, 2400000, 2500000, 2300000, 2600000, 2500000, 2400000, 2300000, 2200000, 2100000]],
    ["1122", "売掛金", [9200000, 8700000, 9600000, 10400000, 10100000, 9900000, 11200000, 10800000, 11600000, 12100000, 11800000]],
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
  await processInputFiles(event.target.files);
});

["dragenter", "dragover"].forEach((eventName) => {
  els.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.dropZone.classList.add("drag-over");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  els.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.dropZone.classList.remove("drag-over");
  });
});

els.dropZone.addEventListener("drop", async (event) => {
  await processInputFiles(event.dataTransfer.files);
});

els.parseTextButton.addEventListener("click", () => {
  try {
    acceptParsedData(parseBalanceText(els.rawTextInput.value), "テキスト");
  } catch (error) {
    els.importStatus.textContent = "要確認";
    els.importStatus.classList.remove("ready");
    setImportMessage(error.message, true);
  }
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

els.addBillDetail.addEventListener("click", () => addReceivableDetail("bill"));
els.addEReceivableDetail.addEventListener("click", () => addReceivableDetail("eReceivable"));

document.querySelectorAll('input[name="collectionMode"]').forEach((input) => {
  input.addEventListener("change", buildForecast);
});

inputIds.forEach((id) => {
  const input = document.querySelector(`#${id}`);
  if (moneyInputIds.includes(id)) {
    input.addEventListener("input", () => {
      formatMoneyInput(input);
      buildForecast();
    });
  }
  input.addEventListener("change", () => {
    renderRows();
    renderMetrics();
    buildForecast();
  });
});

state.billDetails.push(createReceivableDetail());
state.eReceivableDetails.push(createReceivableDetail());
renderReceivableDetails();
renderEmptyTable();
