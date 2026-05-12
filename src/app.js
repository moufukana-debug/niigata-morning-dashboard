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
  requireElements(["todayLabel", "dateInput", "pdfInput", "excelInput", "pdfStatus", "excelStatus", "workSummary", "attentionList", "confirmationList", "taskList", "eventsList", "basicInfo", "memo", "memoStatus"]);
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
  try {
    state.workbookRows = await extractWorkbookRows(file);
    elements.excelStatus.textContent = `✅ ${file.name}（${state.workbookRows.length}行）`;
  } catch (error) {
    showAppError("Excelを読み込めませんでした。", error);
    state.workbookRows = [];
    elements.excelStatus.textContent = "Excelを読み込めませんでした（要確認）";
  }
  refreshDerivedData();
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

async function extractWorkbookRows(file) {
  await waitForXlsx();
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellDates: true });
  return workbook.SheetNames.flatMap((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false, defval: "" });
    return rows.map((cells, rowIndex) => ({
      sheetName,
      rowIndex,
      cells: cells.map((value, columnIndex) => ({ value: normalizeCell(value), columnIndex })).filter((cell) => cell.value),
    }));
  });
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
  state.events = parseEventsFromWorkbook(state.workbookRows, state.selectedDate);
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
  if (!rows.length) return [];
  const directDateEvents = parseDateSectionEvents(rows, date);
  const columnEvents = parseDateColumnEvents(rows, date);
  const merged = [...directDateEvents, ...columnEvents];
  const unique = new Map();
  merged.forEach((event) => {
    const key = [event.group, event.time, event.activity, event.place, event.source].join("|");
    if (!unique.has(key)) unique.set(key, event);
  });
  return Array.from(unique.values()).slice(0, 10);
}

function parseDateSectionEvents(rows, date) {
  const events = [];
  rows.forEach((row) => {
    const dateCells = row.cells.filter((cell) => isExcelDateCell(cell.value, date));
    if (!dateCells.length) return;
    const sameSheetRows = rows.filter((candidate) => candidate.sheetName === row.sheetName && candidate.rowIndex > row.rowIndex && candidate.rowIndex <= row.rowIndex + 12);
    const block = [];
    for (const candidate of sameSheetRows) {
      if (candidate.cells.some((cell) => isAnyDateLikeCell(cell.value)) && candidate.rowIndex !== row.rowIndex + 1) break;
      block.push(candidate);
    }
    block.forEach((candidate) => events.push(rowToEvent(candidate, `${candidate.sheetName} ${candidate.rowIndex + 1}行目`)));
  });
  return events.filter(hasEventContent);
}

function parseDateColumnEvents(rows, date) {
  const events = [];
  rows.forEach((headerRow) => {
    const dateCell = headerRow.cells.find((cell) => isExcelDateCell(cell.value, date));
    if (!dateCell) return;
    const sameSheetRows = rows.filter((candidate) => candidate.sheetName === headerRow.sheetName && candidate.rowIndex > headerRow.rowIndex && candidate.rowIndex <= headerRow.rowIndex + 30);
    sameSheetRows.forEach((row) => {
      const target = row.cells.find((cell) => cell.columnIndex === dateCell.columnIndex);
      if (!target?.value) return;
      const leftLabels = row.cells.filter((cell) => cell.columnIndex < dateCell.columnIndex).map((cell) => cell.value);
      const event = eventFromValues([...leftLabels, target.value], `${row.sheetName} ${row.rowIndex + 1}行目`);
      if (hasEventContent(event)) events.push(event);
    });
  });
  return events;
}

function rowToEvent(row, source) {
  return eventFromValues(row.cells.map((cell) => cell.value), source);
}

function eventFromValues(values, source) {
  const meaningful = values.map((value) => normalizeCell(value)).filter(Boolean).filter((value) => !isMetadataOnly(value));
  if (!meaningful.length) return { group: UNKNOWN, time: UNKNOWN, activity: UNKNOWN, place: UNKNOWN, note: UNKNOWN, source };

  const time = meaningful.find((value) => /\d{1,2}[:：]\d{2}|\d{1,2}時|午前|午後|終日|AM|PM/i.test(value)) || UNKNOWN;
  const place = meaningful.find((value) => /研修室|体育館|食堂|広場|キャンプ|野外|自然|ホール|室|場|館|棟|ロビー|講堂|グラウンド/.test(value)) || UNKNOWN;
  const activity = meaningful.find((value) => /活動|研修|入所|退所|体験|講座|説明|式|会|炊飯|登山|散策|クラフト|オリエン|宿泊|朝食|昼食|夕食/.test(value) && value !== place) || UNKNOWN;
  const group = meaningful.find((value) => value !== time && value !== place && value !== activity && !/備考|注意|担当|場所|活動|時間/.test(value) && value.length >= 2) || UNKNOWN;
  const note = meaningful.filter((value) => ![time, place, activity, group].includes(value)).join(" / ") || UNKNOWN;
  return { group, time, activity, place, note, source };
}

function hasEventContent(event) {
  return [event.group, event.time, event.activity, event.place].some((value) => value && value !== UNKNOWN);
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
    <article class="event-card">
      <h3>${escapeHtml(event.group)}</h3>
      <div class="event-meta">
        <span><strong>時間</strong>：${escapeHtml(event.time)}</span>
        <span><strong>活動</strong>：${escapeHtml(event.activity)}</span>
        <span><strong>場所</strong>：${escapeHtml(event.place)}</span>
        <span><strong>備考</strong>：${escapeHtml(event.note)}</span>
        <span><strong>抽出元</strong>：${escapeHtml(event.source)}</span>
      </div>
    </article>
  `).join("");
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

function isExcelDateCell(value, date) {
  const text = normalizeCell(value);
  return buildDateTokens(date, { includeBareDay: false }).some((token) => text === token || text.includes(token));
}

function isAnyDateLikeCell(value) {
  return /\d{1,2}月\d{1,2}日|\d{4}[/-]\d{1,2}[/-]\d{1,2}|\d{1,2}[/-]\d{1,2}|^\d{1,2}日$/.test(normalizeCell(value));
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
