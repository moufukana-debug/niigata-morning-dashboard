"use strict";

const STAFF_NAME = "西村";
const MEMO_KEY = "niigataMorning.memo";
const TASK_KEY_PREFIX = "niigataMorning.tasks.";
const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];
const DEFAULT_TASKS = ["自分の勤務を確認", "今日の団体・活動を確認", "特記事項・要確認を確認"];
const UNKNOWN = "要確認";

const state = {
  selectedDate: new Date(),
  pdfData: null,
  workbookRows: [],
  work: null,
  events: [],
  attention: [],
  confirmations: ["資料未読み込みのため要確認"],
  excelLog: createEmptyExcelLog(),
  excelDebug: createEmptyExcelDebug(),
};

const elements = {
  appError: document.querySelector("#app-error"),
  appErrorDetail: document.querySelector("#app-error-detail"),
  todayLabel: document.querySelector("#today-label"),
  dateInput: document.querySelector("#date-input"),
  pdfInput: document.querySelector("#pdf-input"),
  excelInput: document.querySelector("#excel-input"),
  pdfStatus: document.querySelector("#pdf-status"),
  excelStatus: document.querySelector("#excel-status"),
  workSummary: document.querySelector("#work-summary"),
  attentionList: document.querySelector("#attention-list"),
  confirmationList: document.querySelector("#confirmation-list"),
  taskList: document.querySelector("#task-list"),
  eventsList: document.querySelector("#events-list"),
  excelLog: document.querySelector("#excel-log"),
  excelDebug: document.querySelector("#excel-debug"),
  basicInfo: document.querySelector("#basic-info"),
  memo: document.querySelector("#memo"),
  memoStatus: document.querySelector("#memo-status"),
};

window.addEventListener("error", (event) => {
  showAppError("JavaScriptの実行中にエラーが発生しました。", event.error || event.message);
});

window.addEventListener("unhandledrejection", (event) => {
  showAppError("非同期処理の読み込みに失敗しました。", event.reason);
});

try {
  init();
} catch (error) {
  showAppError("アプリの初期表示に失敗しました。", error);
}

function init() {
  requireElements(["todayLabel", "dateInput", "pdfInput", "excelInput", "pdfStatus", "excelStatus", "workSummary", "attentionList", "confirmationList", "taskList", "eventsList", "excelLog", "excelDebug", "basicInfo", "memo", "memoStatus"]);
  elements.dateInput.value = toInputDate(state.selectedDate);
  elements.memo.value = getStorageItem(MEMO_KEY) || "";
  elements.dateInput.addEventListener("change", onDateChange);
  elements.pdfInput.addEventListener("change", onPdfSelected);
  elements.excelInput.addEventListener("change", onExcelSelected);
  elements.memo.addEventListener("input", saveMemo);
  render();
  window.NIIGATA_APP_READY = true;
}

function requireElements(keys) {
  const missing = keys.filter((key) => !elements[key]);
  if (missing.length > 0) throw new Error(`画面部品が見つかりません: ${missing.join(", ")}`);
}

function showAppError(message, error) {
  console.error(message, error);
  const detail = error instanceof Error ? error.message : String(error || "原因不明のエラー");
  if (window.NIIGATA_BOOT?.showError) {
    window.NIIGATA_BOOT.showError(message, detail);
    return;
  }
  if (elements.appError) elements.appError.hidden = false;
  if (elements.appErrorDetail) elements.appErrorDetail.textContent = `${message} ${detail}`;
}

function onDateChange(event) {
  const nextDate = new Date(`${event.target.value}T00:00:00`);
  if (Number.isNaN(nextDate.getTime())) return;
  state.selectedDate = nextDate;
  refreshDerivedData();
}

async function onPdfSelected(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  elements.pdfStatus.textContent = "PDFを解析中...";
  try {
    state.pdfData = await extractPdfData(file);
    elements.pdfStatus.textContent = `✅ ${file.name}（${state.pdfData.pages.length}ページ）`;
  } catch (error) {
    showAppError("PDFを読み込めませんでした。", error);
    state.pdfData = null;
    elements.pdfStatus.textContent = "PDFを読み込めませんでした（要確認）";
  }
  refreshDerivedData();
}

async function onExcelSelected(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  elements.excelStatus.textContent = "Excelを解析中...";
  state.excelDebug = createExcelDebugFromFile(file, { extension: getFileExtension(file), label: "未判定" });
  updateExcelDebugStage("ファイル選択イベントは動作しました", { selectionStatus: "ファイル選択イベントは動作しました" });

  try {
    const fileKind = getExcelFileKind(file);
    state.excelDebug.extension = fileKind.extension;
    state.workbookRows = await extractWorkbookRows(file, updateExcelDebugStage);
    state.excelDebug = state.workbookRows.debug ? { ...state.excelDebug, ...state.workbookRows.debug, stageLog: state.excelDebug.stageLog } : state.excelDebug;
    elements.excelStatus.textContent = `✅ ${file.name}（${fileKind.label} / ${state.workbookRows.length}行）`;
    refreshDerivedData();
    updateExcelDebugStage(state.events.length ? `解析結果${state.events.length}件` : "解析結果0件", {
      analysisStatus: state.events.length ? `団体抽出成功（${state.events.length}件）` : `団体抽出失敗（${UNKNOWN}）`,
    });
  } catch (error) {
    showAppError("Excelを読み込めませんでした。", error);
    state.workbookRows = [];
    state.excelLog = createEmptyExcelLog();
    state.excelDebug = error.debug ? { ...state.excelDebug, ...error.debug, stageLog: state.excelDebug?.stageLog || [] } : state.excelDebug || createEmptyExcelDebug();
    state.excelDebug.error = error instanceof Error ? error.message : String(error);
    updateExcelDebugStage(`失敗: ${state.excelDebug.error}`, { analysisStatus: `停止段階: ${state.excelDebug.stageLog?.at(-1) || "不明"}` });
    elements.excelStatus.textContent = formatExcelReadError(error);
    refreshDerivedData();
  }
}

async function extractPdfData(file) {
  const pdfjsLib = await import("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs");
  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    pages.push({ pageNumber, rows: groupTextItemsByRow(textContent.items) });
  }

  return { pages, text: pages.flatMap((page) => page.rows.map((row) => row.text)).join("\n") };
}

function groupTextItemsByRow(items) {
  const rows = new Map();
  items.forEach((item) => {
    const text = normalizeCell(item.str);
    if (!text) return;
    const y = Math.round(item.transform[5] / 3) * 3;
    const row = rows.get(y) || [];
    row.push({ x: item.transform[4], width: item.width || 0, text });
    rows.set(y, row);
  });

  return Array.from(rows.entries())
    .sort((a, b) => b[0] - a[0])
    .map(([y, rowItems]) => {
      const sortedItems = rowItems.sort((a, b) => a.x - b.x);
      return { y, items: sortedItems, text: sortedItems.map((item) => item.text).join(" ").replace(/\s+/g, " ").trim() };
    })
    .filter((row) => row.text);
}

async function extractWorkbookRows(file, onStage = () => {}) {
  const fileKind = getExcelFileKind(file);
  const debug = createExcelDebugFromFile(file, fileKind);
  onStage("XLSXライブラリ確認中", { ...debug });
  await waitForXlsx();

  try {
    debug.fileReadStatus = "読込中";
    onStage("ArrayBuffer取得中", { fileReadStatus: debug.fileReadStatus });
    const data = await file.arrayBuffer();
    debug.fileReadStatus = "成功";
    onStage("ArrayBuffer取得成功", { fileReadStatus: debug.fileReadStatus });
    debug.workbookStatus = "生成中";
    onStage("XLSX.read開始", { workbookStatus: debug.workbookStatus });
    const workbook = XLSX.read(data, { type: "array", cellDates: true, bookVBA: false });
    debug.workbookStatus = "成功";
    onStage("XLSX.read成功", { workbookStatus: debug.workbookStatus });
    if (!workbook.SheetNames?.length) throw new ExcelParseError("シートが見つかりませんでした", debug);

    const sheets = workbook.SheetNames.map((sheetName) => sheetToFilledSheet(sheetName, workbook.Sheets[sheetName]));
    const parsedRows = sheets.flatMap((sheet) => sheet.rows);
    parsedRows.sheets = sheets;
    parsedRows.usedMerges = sheets.some((sheet) => sheet.usedMerges);
    debug.sheetStatus = "成功";
    debug.sheetNames = workbook.SheetNames;
    onStage("シート名一覧取得成功", { sheetStatus: debug.sheetStatus, sheetNames: debug.sheetNames });
    debug.usedMerges = parsedRows.usedMerges;
    debug.sheetSummaries = sheets.map((sheet) => ({
      name: sheet.sheetName,
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      nonEmptyRows: sheet.rows.length,
      mergeCount: sheet.mergeCount,
    }));
    debug.preview = buildExcelPreview(sheets[0]);
    onStage("読み込めた中身のプレビュー作成", { ...debug });
    if (!parsedRows.some((row) => row.cells.length > 0)) throw new ExcelParseError("セルデータを読み取れませんでした", debug);
    debug.analysisStatus = "解析開始";
    onStage("解析開始", { ...debug });
    parsedRows.debug = debug;
    return parsedRows;
  } catch (error) {
    if (error instanceof ExcelFileTypeError || error instanceof ExcelParseError) throw error;
    if (debug.fileReadStatus !== "成功") throw new ExcelParseError(`ファイル読込失敗: ${error instanceof Error ? error.message : String(error)}`, debug);
    throw new ExcelParseError(error instanceof Error ? error.message : String(error), debug);
  }
}

function sheetToFilledSheet(sheetName, sheet) {
  if (!sheet?.["!ref"]) return { sheetName, rows: [], grid: [], usedMerges: false, mergeCount: 0, rowCount: 0, columnCount: 0 };
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const filled = new Map();
  const original = new Set();

  for (let row = range.s.r; row <= range.e.r; row += 1) {
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: column });
      const value = normalizeCell(sheet[address]?.v ?? sheet[address]?.w ?? "");
      if (!value) continue;
      filled.set(`${row}:${column}`, value);
      original.add(`${row}:${column}`);
    }
  }

  (sheet["!merges"] || []).forEach((merge) => {
    const topLeftKey = `${merge.s.r}:${merge.s.c}`;
    const topLeftValue = filled.get(topLeftKey);
    if (!topLeftValue) return;
    for (let row = merge.s.r; row <= merge.e.r; row += 1) {
      for (let column = merge.s.c; column <= merge.e.c; column += 1) {
        const key = `${row}:${column}`;
        if (!filled.has(key)) filled.set(key, topLeftValue);
      }
    }
  });

  const rows = [];
  const grid = [];
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    const cells = [];
    grid[row] = [];
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const key = `${row}:${column}`;
      const value = filled.get(key) || "";
      grid[row][column] = value;
      if (!value) continue;
      cells.push({ value, columnIndex: column, rowIndex: row, fromMerge: !original.has(key) });
    }
    if (cells.length) rows.push({ sheetName, rowIndex: row, cells });
  }
  return {
    sheetName,
    rows,
    grid,
    usedMerges: Boolean(sheet["!merges"]?.length),
    mergeCount: sheet["!merges"]?.length || 0,
    rowCount: range.e.r - range.s.r + 1,
    columnCount: range.e.c - range.s.c + 1,
  };
}

function createEmptyExcelDebug() {
  return {
    fileName: "未選択",
    extension: "-",
    sizeLabel: "-",
    fileReadStatus: "未実行",
    workbookStatus: "未実行",
    sheetStatus: "未実行",
    sheetNames: [],
    usedMerges: false,
    sheetSummaries: [],
    preview: null,
    analysisStatus: "未実行",
    error: "",
    selectionStatus: "未選択",
    stageLog: [],
  };
}

function createExcelDebugFromFile(file, fileKind) {
  return {
    ...createEmptyExcelDebug(),
    fileName: file.name,
    extension: fileKind.extension,
    sizeLabel: formatFileSize(file.size),
    fileReadStatus: "未実行",
    workbookStatus: "未実行",
    sheetStatus: "未実行",
    selectionStatus: "ファイル選択イベントは動作しました",
  };
}

function buildExcelPreview(sheet) {
  if (!sheet) return null;
  const previewRows = [];
  for (let rowIndex = 0; rowIndex < sheet.grid.length && previewRows.length < 20; rowIndex += 1) {
    const row = sheet.grid[rowIndex] || [];
    const values = row.map((value) => normalizeCell(value));
    const nonEmptyValues = values.filter(Boolean);
    if (!nonEmptyValues.length) continue;
    previewRows.push({ rowNumber: rowIndex + 1, nonEmptyCount: nonEmptyValues.length, values: nonEmptyValues.slice(0, 12) });
  }
  return { sheetName: sheet.sheetName, mergeCount: sheet.mergeCount, rows: previewRows };
}

function formatFileSize(size) {
  if (!Number.isFinite(size)) return "不明";
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / 1024 / 1024).toFixed(2)} MB`;
}
class ExcelFileTypeError extends Error {
  constructor(message) {
    super(message);
    this.name = "ExcelFileTypeError";
  }
}

class ExcelParseError extends Error {
  constructor(message, debug = null) {
    super(message);
    this.name = "ExcelParseError";
    this.debug = debug;
  }
}

function getFileExtension(file) {
  return file.name.split(".").pop()?.toLowerCase() || "";
}

function updateExcelDebugStage(stage, patch = {}) {
  const current = state.excelDebug || createEmptyExcelDebug();
  state.excelDebug = {
    ...current,
    ...patch,
    stageLog: [...(current.stageLog || []), stage],
  };
  renderExcelDebug();
}

function getExcelFileKind(file) {
  const extension = getFileExtension(file);
  const supported = {
    xlsx: "Excel（.xlsx）",
    xls: "Excel 97-2003（.xls）",
    xlsm: "Excelマクロ有効ブック（.xlsm / マクロは実行しません）",
  };
  if (!supported[extension]) {
    throw new ExcelFileTypeError("対応しているExcel形式は .xlsx / .xls / .xlsm です");
  }
  return { extension, label: supported[extension] };
}

function formatExcelReadError(error) {
  if (error instanceof ExcelFileTypeError) return `ファイル形式エラー：${error.message}（要確認）`;
  if (error instanceof ExcelParseError) return `中身の解析エラー：${error.message}（要確認）`;
  return "Excelを読み込めませんでした（要確認）";
}

function waitForXlsx() {
  return new Promise((resolve, reject) => {
    let tries = 0;
    const timer = setInterval(() => {
      tries += 1;
      if (window.XLSX) {
        clearInterval(timer);
        resolve();
      }
      if (tries > 50) {
        clearInterval(timer);
        reject(new Error("SheetJSの読み込みに失敗しました"));
      }
    }, 100);
  });
}

function refreshDerivedData() {
  state.work = parseWorkFromPdf(state.pdfData, state.selectedDate);
  state.attention = parseSpecialNotesFromPdf(state.pdfData, state.selectedDate);
  const excelResult = parseEventsFromWorkbook(state.workbookRows, state.selectedDate);
  state.events = excelResult.events;
  state.excelLog = excelResult.log;
  state.excelDebug.analysisStatus = state.workbookRows.length ? (state.events.length ? `団体抽出成功（${state.events.length}件）` : `団体抽出失敗（${UNKNOWN}）`) : state.excelDebug.analysisStatus;
  state.confirmations = buildConfirmations();
  render();
}

function parseWorkFromPdf(pdfData, date) {
  if (!pdfData) return null;
  const candidates = [];

  pdfData.pages.forEach((page) => {
    const staffRows = page.rows.filter((row) => row.text.includes(STAFF_NAME));
    const dateRows = page.rows.filter((row) => row.items.some((item) => isDateHeaderText(item.text, date)));

    staffRows.forEach((staffRow) => {
      dateRows.forEach((dateRow) => {
        const dateCell = findTargetDateCell(dateRow, date);
        if (!dateCell) return;
        const dateCells = collectDateCells(dateRow);
        const bounds = estimateColumnBounds(dateCells, dateCell);
        const value = extractTextInBounds(staffRow.items, bounds, [STAFF_NAME]);
        const verticalDistance = Math.abs(staffRow.y - dateRow.y);
        candidates.push({ value, staffRow, dateRow, bounds, verticalDistance, pageNumber: page.pageNumber });
      });
    });
  });

  const best = candidates
    .filter((candidate) => candidate.value)
    .sort((a, b) => a.verticalDistance - b.verticalDistance)[0];

  if (best) return buildWorkResult(best.value, `PDF座標抽出 ${best.pageNumber}ページ`);

  const fallback = parseWorkFromPdfText(pdfData.text, date);
  return fallback || { status: UNKNOWN, division: "西村行または日付列を検出できません", time: UNKNOWN, source: "PDF抽出不可", confidence: "low" };
}

function parseWorkFromPdfText(text, date) {
  if (!text) return null;
  const lines = text.split(/\n+/).map((line) => line.trim()).filter(Boolean);
  const staffLine = lines.find((line) => line.includes(STAFF_NAME));
  if (!staffLine) return null;

  const dateTokens = buildDateTokens(date);
  const dateLine = lines.find((line) => dateTokens.some((token) => line.includes(token)) && /\d/.test(line));
  const staffCells = splitTableLikeLine(staffLine);
  const staffIndex = staffCells.findIndex((cell) => cell.includes(STAFF_NAME));
  const dateCells = dateLine ? splitTableLikeLine(dateLine) : [];
  const dateIndex = findDateCellIndex(dateCells, date);
  const offsetIndex = dateIndex >= 0 && staffIndex >= 0 ? staffIndex + dateIndex : -1;
  const rawValue = staffCells[offsetIndex] || findNearbyDateValue(lines, staffLine, date) || UNKNOWN;
  return buildWorkResult(rawValue, dateIndex >= 0 ? "PDFテキスト抽出" : "PDF簡易抽出（要確認）");
}

function buildWorkResult(rawValue, source) {
  const cleaned = cleanWorkValue(rawValue);
  return {
    status: cleaned || UNKNOWN,
    division: classifyWork(cleaned),
    time: inferWorkTime(cleaned),
    source,
    confidence: cleaned && cleaned !== UNKNOWN && !source.includes("要確認") ? "medium" : "low",
  };
}

function parseSpecialNotesFromPdf(pdfData, date) {
  if (!pdfData) return [];
  const dateTokens = buildDateTokens(date, { includeBareDay: false });
  const notes = [];

  pdfData.pages.forEach((page) => {
    const dateRows = page.rows.filter((row) => row.items.some((item) => isDateHeaderText(item.text, date)));
    const noteRows = page.rows.filter((row) => /特記|備考|注意|連絡|変更|確認/.test(row.text));

    noteRows.forEach((row) => {
      const direct = dateTokens.some((token) => row.text.includes(token));
      const matchedDateRow = dateRows[0];
      const dateCell = matchedDateRow ? findTargetDateCell(matchedDateRow, date) : null;
      const bounds = matchedDateRow && dateCell ? estimateColumnBounds(collectDateCells(matchedDateRow), dateCell) : null;
      const boundedValue = bounds ? extractTextInBounds(row.items, bounds, []) : "";
      const value = direct ? row.text : boundedValue;
      if (value) notes.push(`PDF特記事項: ${value}`);
    });
  });

  return uniqueStrings(notes).slice(0, 5);
}

function parseEventsFromWorkbook(rows, date) {
  const log = createExcelLog(rows);
  if (!rows.length) return { events: [], log };

  const dateCandidates = findExcelDateCandidates(rows, date);
  const blockCandidates = findGroupBlockCandidates(rows, dateCandidates);
  log.dateCandidates = dateCandidates.length;
  log.groupCandidates = blockCandidates.length;

  const blockEvents = parseGroupBlockEvents(rows, dateCandidates, blockCandidates);
  const fallbackEvents = blockEvents.length ? [] : parseFallbackBlockCandidates(blockCandidates);
  const unique = new Map();

  [...blockEvents, ...fallbackEvents].forEach((event) => {
    if (!hasEventContent(event)) return;
    const key = [event.group, event.time, event.activity, event.place, event.staff, event.source].join("|");
    if (!unique.has(key)) unique.set(key, event);
  });

  const events = Array.from(unique.values()).sort((a, b) => eventRank(a) - eventRank(b)).slice(0, 12);
  log.adopted = events.length;
  log.exact = events.filter((event) => event.matchLevel === "exact").length;
  log.candidates = events.filter((event) => event.matchLevel === "candidate").length;
  log.needConfirmation = events.filter((event) => event.matchLevel === "confirm").length;
  log.mode = log.exact ? "完全一致" : log.candidates ? "候補あり" : log.needConfirmation ? "要確認" : "未検出";
  return { events, log };
}

function createEmptyExcelLog() {
  return { sheets: [], usedMerges: false, dateCandidates: 0, groupCandidates: 0, adopted: 0, exact: 0, candidates: 0, needConfirmation: 0, mode: "未読み込み" };
}

function createExcelLog(rows) {
  const log = createEmptyExcelLog();
  log.sheets = rows.sheets?.map((sheet) => sheet.sheetName) || Array.from(new Set(rows.map((row) => row.sheetName)));
  log.usedMerges = Boolean(rows.usedMerges || rows.sheets?.some((sheet) => sheet.usedMerges));
  log.mode = rows.length ? "解析中" : "未読み込み";
  return log;
}

function findExcelDateCandidates(rows, date) {
  const candidates = [];
  const seen = new Set();
  rows.forEach((row) => {
    row.cells.forEach((cell) => {
      const match = getDateMatchLevel(cell.value, date);
      if (!match) return;
      const key = `${row.sheetName}:${row.rowIndex}:${cell.columnIndex}:${match.level}`;
      if (seen.has(key)) return;
      seen.add(key);
      candidates.push({ ...cell, sheetName: row.sheetName, rowIndex: row.rowIndex, matchLevel: match.level, label: match.label, evidence: `セル日付: ${cell.value}` });
    });

    const rowMatch = getDateMatchLevel(row.cells.map((cell) => cell.value).join(" "), date);
    if (rowMatch) {
      const firstColumn = row.cells[0]?.columnIndex ?? 0;
      const key = `${row.sheetName}:${row.rowIndex}:${firstColumn}:row:${rowMatch.level}`;
      if (!seen.has(key)) {
        seen.add(key);
        candidates.push({ value: rowMatch.label, sheetName: row.sheetName, rowIndex: row.rowIndex, columnIndex: firstColumn, matchLevel: rowMatch.level, label: rowMatch.label, evidence: "行内の日付分割候補" });
      }
    }
  });
  return candidates;
}

function findGroupBlockCandidates(rows, dateCandidates) {
  const bySheet = new Map();
  rows.forEach((row) => {
    if (!bySheet.has(row.sheetName)) bySheet.set(row.sheetName, []);
    bySheet.get(row.sheetName).push(row);
  });

  const blocks = [];
  bySheet.forEach((sheetRows, sheetName) => {
    const sortedRows = sheetRows.slice().sort((a, b) => a.rowIndex - b.rowIndex);
    const groupRows = sortedRows.filter((row) => row.cells.some((cell) => isGroupLikeCell(cell.value)));
    groupRows.forEach((row) => {
      const groupCell = row.cells.find((cell) => isGroupLikeCell(cell.value));
      if (!groupCell) return;
      const start = Math.max(sortedRows[0].rowIndex, row.rowIndex - 2);
      const end = Math.min(sortedRows.at(-1).rowIndex, row.rowIndex + 8);
      const blockRows = sortedRows.filter((candidate) => candidate.rowIndex >= start && candidate.rowIndex <= end);
      const nearestDate = findNearestDateCandidate(dateCandidates, groupCell, sheetName);
      blocks.push({ sheetName, groupCell, startRow: start, endRow: end, rows: blockRows, nearestDate });
    });
  });
  return dedupeBlocks(blocks);
}

function dedupeBlocks(blocks) {
  const seen = new Set();
  return blocks.filter((block) => {
    const key = `${block.sheetName}:${block.groupCell.value}:${block.startRow}:${block.endRow}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function findNearestDateCandidate(dateCandidates, cell, sheetName) {
  return dateCandidates
    .filter((dateCell) => dateCell.sheetName === sheetName)
    .map((dateCell) => ({
      dateCell,
      distance: Math.abs(dateCell.rowIndex - cell.rowIndex) + Math.min(Math.abs(dateCell.columnIndex - cell.columnIndex), 8),
      rowDistance: Math.abs(dateCell.rowIndex - cell.rowIndex),
    }))
    .filter(({ distance, rowDistance }) => distance <= 22 || rowDistance <= 14)
    .sort((a, b) => a.distance - b.distance)[0]?.dateCell || null;
}

function parseGroupBlockEvents(rows, dateCandidates, blockCandidates) {
  return blockCandidates
    .map((block) => blockToEvent(block, dateCandidates))
    .filter(hasEventContent);
}

function blockToEvent(block) {
  const values = block.rows.flatMap((row) => row.cells.map((cell) => cell.value));
  const rowValues = block.rows.map((row) => row.cells.map((cell) => cell.value).join(" "));
  const nearestDate = block.nearestDate;
  const dateEvidence = nearestDate ? `${nearestDate.evidence || "日付候補"}: ${nearestDate.value}` : "日付一致なし";
  const rawMatch = nearestDate?.matchLevel || "confirm";
  const event = eventFromValues(values, `${block.sheetName} ${block.startRow + 1}-${block.endRow + 1}行目`, rawMatch, dateEvidence);
  event.group = block.groupCell.value || event.group;
  event.staff = findFirstByPredicate(rowValues, isStaffLikeText) || UNKNOWN;
  event.note = buildBlockNote(event.note, rowValues);
  const requiredMissing = [event.group, event.time, event.activity, event.place].includes(UNKNOWN);
  if (!nearestDate) event.matchLevel = "confirm";
  else if (nearestDate.matchLevel === "exact" && !requiredMissing) event.matchLevel = "exact";
  else event.matchLevel = "candidate";
  if (event.matchLevel !== "exact" && !event.evidence.includes(UNKNOWN)) event.evidence = `${event.evidence}（${UNKNOWN}）`;
  return event;
}

function parseFallbackBlockCandidates(blockCandidates) {
  return blockCandidates.slice(0, 3).map((block) => {
    const event = blockToEvent({ ...block, nearestDate: null });
    event.matchLevel = "confirm";
    event.evidence = `日付一致なし。団体ブロック候補のみ（${UNKNOWN}）`;
    return event;
  });
}

function eventFromValues(values, source, matchLevel = "candidate", evidence = "候補") {
  const meaningful = uniqueStrings(values).filter((value) => !isMetadataOnly(value));
  if (!meaningful.length) return { group: UNKNOWN, time: UNKNOWN, activity: UNKNOWN, place: UNKNOWN, staff: UNKNOWN, note: UNKNOWN, source, matchLevel: "confirm", evidence };

  const time = meaningful.find(isTimeLikeCell) || UNKNOWN;
  const place = meaningful.find(isPlaceLikeCell) || UNKNOWN;
  const activity = meaningful.find((value) => isActivityLikeCell(value) && value !== place) || UNKNOWN;
  const staff = meaningful.find(isStaffLikeText) || UNKNOWN;
  const group = meaningful.find((value) => value !== time && value !== place && value !== activity && value !== staff && isGroupLikeCell(value)) || UNKNOWN;
  const note = meaningful.filter((value) => ![time, place, activity, place, group, staff].includes(value)).join(" / ") || UNKNOWN;
  const needsConfirmation = [group, time, activity, place].includes(UNKNOWN) || matchLevel !== "exact";
  return { group, time, activity, place, staff, note, source, matchLevel: needsConfirmation ? (matchLevel === "exact" ? "candidate" : matchLevel) : "exact", evidence };
}

function buildBlockNote(existingNote, rowValues) {
  const notes = rowValues.filter((value) => /備考|注意|持参|雨天|変更|確認/.test(value));
  return uniqueStrings([existingNote, ...notes].filter((value) => value && value !== UNKNOWN)).join(" / ") || UNKNOWN;
}

function findFirstByPredicate(values, predicate) {
  return values.map((value) => normalizeCell(value)).find(predicate) || "";
}

function hasEventContent(event) {
  return [event.group, event.time, event.activity, event.place, event.staff, event.note].some((value) => value && value !== UNKNOWN);
}

function eventRank(event) {
  return { exact: 0, candidate: 1, confirm: 2 }[event.matchLevel] ?? 3;
}
function buildConfirmations() {
  const confirmations = [];
  if (!state.pdfData) confirmations.push("勤務予定表PDFが未読み込みです");
  if (!state.workbookRows.length) confirmations.push("調整プログラムExcelが未読み込みです");
  if (!state.work || state.work.status === UNKNOWN || state.work.confidence === "low") confirmations.push("西村さんの勤務は原本で要確認");
  if (state.pdfData && !state.attention.length) confirmations.push("今日の特記事項を自動検出できませんでした");
  if (state.events.some((event) => [event.group, event.time, event.activity, event.place, event.note].includes(UNKNOWN))) confirmations.push("団体・行事に不足項目があります");
  if (state.workbookRows.length && !state.events.length) confirmations.push("今日の日付に一致する団体・行事を検出できませんでした");
  return confirmations.length ? confirmations.map((item) => `${item}（${UNKNOWN}）`) : ["自動抽出で不足は見つかりませんでした。最終確認は原本で行ってください"];
}

function render() {
  elements.todayLabel.textContent = `${formatJapaneseDate(state.selectedDate)} ／ 西村さん`;
  renderWork();
  renderAttention();
  renderConfirmations();
  renderTasks();
  renderEvents();
  renderExcelLog();
  renderExcelDebug();
  renderBasicInfo();
}

function renderWork() {
  if (!state.work) {
    elements.workSummary.className = "summary-box empty-state";
    elements.workSummary.textContent = "PDFを読み込むと、西村さんの勤務情報を表示します。";
    return;
  }
  elements.workSummary.className = "summary-box";
  elements.workSummary.innerHTML = `
    <div class="work-main">${escapeHtml(state.work.status)}</div>
    <div class="work-sub">勤務区分：${escapeHtml(state.work.division)}</div>
    <div class="work-sub">勤務時間：${escapeHtml(state.work.time)}</div>
    <div class="work-sub">抽出方法：${escapeHtml(state.work.source)}</div>
  `;
}

function renderAttention() {
  const items = state.attention.length ? state.attention : ["今日の特記事項は未検出です。原本確認が必要です。"];
  elements.attentionList.innerHTML = items.map((item) => `<li>${escapeHtml(item)}</li>`).join("");
}

function renderConfirmations() {
  elements.confirmationList.innerHTML = state.confirmations.map((item) => `<li>${escapeHtml(item)}</li>`).join("");
}

function renderTasks() {
  const key = `${TASK_KEY_PREFIX}${toInputDate(state.selectedDate)}`;
  const saved = getStorageJson(key, {});
  elements.taskList.innerHTML = DEFAULT_TASKS.map((task, index) => `
    <label class="task-item">
      <input type="checkbox" data-task-index="${index}" ${saved[index] ? "checked" : ""} />
      <span>${escapeHtml(task)}</span>
    </label>
  `).join("");
  elements.taskList.querySelectorAll("input").forEach((input) => {
    input.addEventListener("change", () => {
      const next = getStorageJson(key, {});
      next[input.dataset.taskIndex] = input.checked;
      setStorageItem(key, JSON.stringify(next));
    });
  });
}

function renderEvents() {
  if (!state.workbookRows.length) {
    elements.eventsList.className = "events-list empty-state";
    elements.eventsList.textContent = "Excelを読み込むと、今日の団体・活動を表示します。";
    return;
  }
  if (!state.events.length) {
    elements.eventsList.className = "events-list empty-state";
    elements.eventsList.textContent = "今日の日付に一致する団体・行事を検出できません。原本で要確認。";
    return;
  }
  elements.eventsList.className = "events-list";
  elements.eventsList.innerHTML = state.events.map((event) => `
    <article class="event-card ${escapeHtml(event.matchLevel || "candidate")}">
      <div class="event-title-row">
        <h3>${escapeHtml(event.group)}</h3>
        <span class="match-label ${escapeHtml(event.matchLevel || "candidate")}">${escapeHtml(matchLevelLabel(event.matchLevel))}</span>
      </div>
      <div class="event-meta">
        <span><strong>時間</strong>：${escapeHtml(event.time)}</span>
        <span><strong>活動</strong>：${escapeHtml(event.activity)}</span>
        <span><strong>場所</strong>：${escapeHtml(event.place)}</span>
        <span><strong>担当</strong>：${escapeHtml(event.staff || UNKNOWN)}</span>
        <span><strong>備考</strong>：${escapeHtml(event.note)}</span>
        <span><strong>根拠</strong>：${escapeHtml(event.evidence || UNKNOWN)}</span>
        <span><strong>抽出元</strong>：${escapeHtml(event.source)}</span>
      </div>
    </article>
  `).join("");
}

function renderExcelLog() {
  const log = state.excelLog || createEmptyExcelLog();
  elements.excelLog.innerHTML = `
    <strong>解析ログ</strong>
    <span>シート: ${escapeHtml(log.sheets.length ? log.sheets.join(" / ") : "未読み込み")}</span>
    <span>結合セル補完: ${escapeHtml(log.usedMerges ? "あり" : "なし")}</span>
    <span>日付候補: ${escapeHtml(log.dateCandidates)}件</span>
    <span>団体候補: ${escapeHtml(log.groupCandidates)}件</span>
    <span>採用: ${escapeHtml(log.adopted)}件（完全一致 ${escapeHtml(log.exact)} / 候補あり ${escapeHtml(log.candidates)} / 要確認 ${escapeHtml(log.needConfirmation)}）</span>
    <span>状態: ${escapeHtml(log.mode)}</span>
  `;
}

function renderExcelDebug() {
  const debug = state.excelDebug || createEmptyExcelDebug();
  const sheetSummary = debug.sheetSummaries.length
    ? debug.sheetSummaries.map((sheet) => `<li>${escapeHtml(sheet.name)}：${escapeHtml(sheet.rowCount)}行 × ${escapeHtml(sheet.columnCount)}列 / 非空行 ${escapeHtml(sheet.nonEmptyRows)} / 結合 ${escapeHtml(sheet.mergeCount)}</li>`).join("")
    : "<li>シート未検出</li>";
  const preview = debug.preview
    ? `
      <div class="preview-box">
        <strong>先頭シートプレビュー：${escapeHtml(debug.preview.sheetName)}（結合セル ${escapeHtml(debug.preview.mergeCount)}件）</strong>
        ${debug.preview.rows.map((row) => `<div class="preview-row"><b>${escapeHtml(row.rowNumber)}行目</b> 非空${escapeHtml(row.nonEmptyCount)}：${escapeHtml(row.values.join(" / "))}</div>`).join("") || "<div>表示できるセルがありません</div>"}
      </div>
    `
    : "<div class=\"preview-box\">プレビュー未作成</div>";

  elements.excelDebug.innerHTML = `
    <strong>読込デバッグ</strong>
    <span>${escapeHtml(debug.selectionStatus)}</span>
    <span>ファイル名: ${escapeHtml(debug.fileName)}</span>
    <span>拡張子: ${escapeHtml(debug.extension)} / サイズ: ${escapeHtml(debug.sizeLabel)}</span>
    <span>ファイル読込: ${escapeHtml(debug.fileReadStatus)} / Workbook生成: ${escapeHtml(debug.workbookStatus)} / シート検出: ${escapeHtml(debug.sheetStatus)}</span>
    <span>シート名: ${escapeHtml(debug.sheetNames.length ? debug.sheetNames.join(" / ") : "未検出")}</span>
    <span>結合セル補完: ${escapeHtml(debug.usedMerges ? "あり" : "なし")}</span>
    <span>解析段階: ${escapeHtml(debug.analysisStatus)}</span>
    ${debug.error ? `<span>エラー: ${escapeHtml(debug.error)}</span>` : ""}
    <span>段階ログ: ${escapeHtml((debug.stageLog || []).join(" → ") || "未開始")}</span>
    <ul>${sheetSummary}</ul>
    ${preview}
  `;
}

function matchLevelLabel(level) {
  if (level === "exact") return "完全一致";
  if (level === "candidate") return "候補あり・要確認";
  return "要確認";
}

function renderBasicInfo() {
  const rows = [
    ["勤務", state.work?.status || `PDF未読み込み（${UNKNOWN}）`],
    ["勤務区分", state.work?.division || UNKNOWN],
    ["特記事項", state.attention[0] || UNKNOWN],
    ["団体件数", state.events.length ? `${state.events.length}件` : UNKNOWN],
  ];
  elements.basicInfo.innerHTML = rows.map(([label, value]) => `<div><dt>${escapeHtml(label)}</dt><dd>${escapeHtml(value)}</dd></div>`).join("");
}

function splitTableLikeLine(line) {
  return line.split(/\s{1,}|\||,|、/).map((cell) => cell.trim()).filter(Boolean);
}

function findDateCellIndex(cells, date) {
  const tokens = buildDateTokens(date);
  return cells.findIndex((cell) => tokens.some((token) => cell === token || cell.includes(token)));
}

function findNearbyDateValue(lines, staffLine, date) {
  const staffLineIndex = lines.indexOf(staffLine);
  const day = String(date.getDate());
  const candidate = lines.slice(Math.max(0, staffLineIndex - 2), staffLineIndex + 3).find((line) => line.includes(day) && line.includes(STAFF_NAME));
  if (!candidate) return "";
  const match = candidate.match(new RegExp(`${day}[^\\n]{0,12}(${STAFF_NAME})?\\s*([^\\s]{1,8})`));
  return match?.[2] || "";
}

function findTargetDateCell(row, date) {
  return row.items.find((item) => isDateHeaderText(item.text, date));
}

function collectDateCells(row) {
  return row.items.filter((item) => /^(\d{1,2}|\d{1,2}日|\d{1,2}\/\d{1,2}|\d{1,2}月\d{1,2}日)$/.test(item.text));
}

function estimateColumnBounds(dateCells, targetCell) {
  const sorted = dateCells.length ? dateCells.slice().sort((a, b) => a.x - b.x) : [targetCell];
  const index = sorted.indexOf(targetCell);
  const previous = sorted[index - 1];
  const next = sorted[index + 1];
  return {
    left: previous ? (previous.x + targetCell.x) / 2 : targetCell.x - 12,
    right: next ? (targetCell.x + next.x) / 2 : targetCell.x + Math.max(18, targetCell.width + 12),
  };
}

function extractTextInBounds(items, bounds, excludeTexts) {
  return items
    .filter((item) => item.x >= bounds.left && item.x < bounds.right)
    .map((item) => item.text)
    .filter((text) => !excludeTexts.some((exclude) => text.includes(exclude)))
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

function isDateHeaderText(value, date) {
  const text = normalizeCell(value).replace(/[（）()]/g, "");
  const month = date.getMonth() + 1;
  const day = date.getDate();
  return text === String(day) || text === `${day}日` || text === `${month}/${day}` || text === `${month}月${day}日`;
}

function getDateMatchLevel(value, date) {
  const text = normalizeCell(value).replace(/[（）()\s]/g, "");
  if (!text) return null;
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const weekday = WEEKDAYS[date.getDay()];
  const exactTokens = [`${date.getFullYear()}/${month}/${day}`, `${date.getFullYear()}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`, `${month}/${day}`, `${month}月${day}日`, `${day}日`];
  if (exactTokens.some((token) => text === token || text.includes(token))) return { level: "exact", label: text };
  if (text === String(day)) return { level: "candidate", label: text };
  if (text === weekday || text.includes(`(${weekday})`) || text.includes(`（${weekday}）`)) return { level: "confirm", label: text };
  return null;
}

function isTimeLikeCell(value) {
  return /\d{1,2}[:：]\d{2}|\d{1,2}時|午前|午後|終日|AM|PM|集合|解散/i.test(normalizeCell(value));
}

function isPlaceLikeCell(value) {
  return /研修室|体育館|食堂|広場|キャンプ|野外|自然|ホール|室|場|館|棟|ロビー|講堂|グラウンド|駐車場|広間|交流/.test(normalizeCell(value));
}

function isActivityLikeCell(value) {
  return /活動|研修|入所|退所|体験|講座|説明|式|会|炊飯|登山|散策|クラフト|オリエン|宿泊|朝食|昼食|夕食|清掃|準備|片付|受付|キャンプ/.test(normalizeCell(value));
}

function isStaffLikeText(value) {
  const text = normalizeCell(value);
  return /担当|職員|係|指導|所員|スタッフ/.test(text) && text.length <= 40;
}

function isGroupLikeCell(value) {
  const text = normalizeCell(value);
  if (!text || text.length < 2 || text.length > 40) return false;
  if (isTimeLikeCell(text) || isPlaceLikeCell(text) || isAnyDateLikeCell(text) || /備考|注意|担当|場所|活動|時間|予定|プログラム|午前|午後/.test(text)) return false;
  return /学校|小|中|高|大学|園|団|会|クラブ|協会|連盟|市|町|村|利用|研修|子ども|こども|少年|自然|センター|教室|講座|チーム|会議|保育|幼稚/.test(text) || /[一-龠ぁ-んァ-ヶ]{3,}/.test(text);
}

function isExcelDateCell(value, date) {
  return Boolean(getDateMatchLevel(value, date));
}

function isAnyDateLikeCell(value) {
  return /\d{1,2}月\d{1,2}日|\d{4}[/-]\d{1,2}[/-]\d{1,2}|\d{1,2}[/-]\d{1,2}|^\d{1,2}日$|^\d{1,2}$|^[日月火水木金土]$/.test(normalizeCell(value));
}

function isMetadataOnly(value) {
  return isAnyDateLikeCell(value) || /曜日|団体名|活動内容|場所|時間|備考|予定|プログラム/.test(value);
}

function cleanWorkValue(value) {
  return normalizeCell(value).replace(STAFF_NAME, "").replace(/[：:]/g, "").trim() || UNKNOWN;
}

function classifyWork(value) {
  if (!value || value === UNKNOWN) return UNKNOWN;
  if (/休|休日|週休|年休|代休/.test(value)) return "休日";
  if (/日直|宿直|直/.test(value)) return "日直・当番系";
  return "勤務記号";
}

function inferWorkTime(value) {
  const text = String(value || "");
  const timeMatch = text.match(/\d{1,2}[:：]\d{2}\s*[〜~\-－]\s*\d{1,2}[:：]\d{2}/);
  if (timeMatch) return timeMatch[0].replace("：", ":");
  if (/休|休日|週休|年休|代休/.test(text)) return "休日";
  return `資料の記号を確認（${UNKNOWN}）`;
}

function saveMemo() {
  setStorageItem(MEMO_KEY, elements.memo.value);
  elements.memoStatus.textContent = "保存しました。";
  clearTimeout(saveMemo.timer);
  saveMemo.timer = setTimeout(() => {
    elements.memoStatus.textContent = "入力すると自動保存します。";
  }, 1400);
}

function buildDateTokens(date, options = {}) {
  const { includeBareDay = true } = options;
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const reiwa = year - 2018;
  const tokens = [
    `${year}/${month}/${day}`,
    `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`,
    `${month}/${day}`,
    `${month}月${day}日`,
    `${day}日`,
    `R${reiwa}.${String(month).padStart(2, "0")}.${String(day).padStart(2, "0")}`,
  ];
  if (includeBareDay) tokens.push(`${day}`);
  return tokens;
}

function getStorageItem(key) {
  try {
    return window.localStorage.getItem(key);
  } catch (error) {
    showAppError("ブラウザの保存領域を読み取れませんでした。", error);
    return null;
  }
}

function setStorageItem(key, value) {
  try {
    window.localStorage.setItem(key, value);
  } catch (error) {
    showAppError("ブラウザの保存領域に保存できませんでした。", error);
  }
}

function getStorageJson(key, fallback) {
  const raw = getStorageItem(key);
  if (!raw) return fallback;
  try {
    return JSON.parse(raw);
  } catch (error) {
    showAppError("保存済みチェック項目の読み込みに失敗しました。", error);
    return fallback;
  }
}

function normalizeCell(value) {
  if (value instanceof Date) return `${value.getFullYear()}/${value.getMonth() + 1}/${value.getDate()}`;
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function uniqueStrings(values) {
  return Array.from(new Set(values.map((value) => normalizeCell(value)).filter(Boolean)));
}

function formatJapaneseDate(date) {
  return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日（${WEEKDAYS[date.getDay()]}）`;
}

function toInputDate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[char]));
}
