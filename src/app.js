"use strict";

const STAFF_NAME = "西村";
const MEMO_KEY = "niigataMorning.memo";
const TASK_KEY_PREFIX = "niigataMorning.tasks.";
const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];
const DEFAULT_TASKS = ["自分の勤務を確認", "今日の団体・活動を確認", "特記事項・要確認を確認"];

const state = {
  selectedDate: new Date(),
  pdfText: "",
  workbookRows: [],
  work: null,
  events: [],
  warnings: ["資料未読み込みのため要確認"],
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
  requireElements(["todayLabel", "dateInput", "pdfInput", "excelInput", "pdfStatus", "excelStatus", "workSummary", "attentionList", "taskList", "eventsList", "basicInfo", "memo", "memoStatus"]);
  elements.dateInput.value = toInputDate(state.selectedDate);
  elements.memo.value = getStorageItem(MEMO_KEY) || "";
  elements.dateInput.addEventListener("change", onDateChange);
  elements.pdfInput.addEventListener("change", onPdfSelected);
  elements.excelInput.addEventListener("change", onExcelSelected);
  elements.memo.addEventListener("input", saveMemo);
  render();
}

function requireElements(keys) {
  const missing = keys.filter((key) => !elements[key]);
  if (missing.length > 0) {
    throw new Error(`画面部品が見つかりません: ${missing.join(", ")}`);
  }
}

function showAppError(message, error) {
  console.error(message, error);
  const detail = error instanceof Error ? error.message : String(error || "原因不明のエラー");
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
  elements.pdfStatus.textContent = "読み込み中...";
  try {
    state.pdfText = await extractPdfText(file);
    elements.pdfStatus.textContent = `${file.name} を読み込みました`;
  } catch (error) {
    showAppError("PDFを読み込めませんでした。", error);
    state.pdfText = "";
    elements.pdfStatus.textContent = "PDFを読み込めませんでした（要確認）";
  }
  refreshDerivedData();
}

async function onExcelSelected(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  elements.excelStatus.textContent = "読み込み中...";
  try {
    state.workbookRows = await extractWorkbookRows(file);
    elements.excelStatus.textContent = `${file.name} を読み込みました`;
  } catch (error) {
    showAppError("Excelを読み込めませんでした。", error);
    state.workbookRows = [];
    elements.excelStatus.textContent = "Excelを読み込めませんでした（要確認）";
  }
  refreshDerivedData();
}

async function extractPdfText(file) {
  const pdfjsLib = await import("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs");
  pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const pages = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    const lines = groupTextItemsByLine(textContent.items);
    pages.push(lines.join("\n"));
  }

  return pages.join("\n");
}

function groupTextItemsByLine(items) {
  const rows = new Map();
  items.forEach((item) => {
    const y = Math.round(item.transform[5] / 4) * 4;
    const row = rows.get(y) || [];
    row.push({ x: item.transform[4], text: item.str });
    rows.set(y, row);
  });

  return Array.from(rows.entries())
    .sort((a, b) => b[0] - a[0])
    .map(([, row]) => row.sort((a, b) => a.x - b.x).map((item) => item.text).join(" ").replace(/\s+/g, " ").trim())
    .filter(Boolean);
}

async function extractWorkbookRows(file) {
  await waitForXlsx();
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array", cellDates: true });
  return workbook.SheetNames.flatMap((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false, defval: "" });
    return rows.map((cells, rowIndex) => ({ sheetName, rowIndex, cells: cells.map(normalizeCell) }));
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
  state.work = parseWorkFromPdf(state.pdfText, state.selectedDate);
  state.events = parseEventsFromWorkbook(state.workbookRows, state.selectedDate);
  state.warnings = buildWarnings();
  render();
}

function parseWorkFromPdf(text, date) {
  if (!text) return null;
  const lines = text.split(/\n+/).map((line) => line.trim()).filter(Boolean);
  const staffLine = lines.find((line) => line.includes(STAFF_NAME));
  const dateTokens = buildDateTokens(date);
  const dateLine = lines.find((line) => dateTokens.some((token) => line.includes(token)) && /\d/.test(line));

  if (!staffLine) {
    return { status: "要確認", division: "西村行を検出できません", time: "要確認", source: "PDF簡易抽出" };
  }

  const staffCells = splitTableLikeLine(staffLine);
  const staffIndex = staffCells.findIndex((cell) => cell.includes(STAFF_NAME));
  const dateCells = dateLine ? splitTableLikeLine(dateLine) : [];
  const dateIndex = findDateCellIndex(dateCells, date);
  const offsetIndex = dateIndex >= 0 && staffIndex >= 0 ? staffIndex + dateIndex : -1;
  const rawValue = staffCells[offsetIndex] || findNearbyDateValue(lines, staffLine, date) || "要確認";
  const cleaned = cleanWorkValue(rawValue);

  return {
    status: cleaned || "要確認",
    division: classifyWork(cleaned),
    time: inferWorkTime(cleaned),
    source: dateIndex >= 0 ? "PDF日付列優先" : "PDF簡易抽出（要確認）",
    raw: staffLine,
  };
}

function splitTableLikeLine(line) {
  return line.split(/\s{1,}|\||,|、/).map((cell) => cell.trim()).filter(Boolean);
}

function findDateCellIndex(cells, date) {
  const day = String(date.getDate());
  const tokens = buildDateTokens(date);
  return cells.findIndex((cell) => cell === day || tokens.some((token) => cell.includes(token)));
}

function findNearbyDateValue(lines, staffLine, date) {
  const staffLineIndex = lines.indexOf(staffLine);
  const day = String(date.getDate());
  const candidate = lines.slice(Math.max(0, staffLineIndex - 2), staffLineIndex + 3).find((line) => line.includes(day) && line.includes(STAFF_NAME));
  if (!candidate) return "";
  const match = candidate.match(new RegExp(`${day}[^\\n]{0,12}(${STAFF_NAME})?\\s*([^\\s]{1,8})`));
  return match?.[2] || "";
}

function cleanWorkValue(value) {
  return String(value).replace(STAFF_NAME, "").replace(/[：:]/g, "").trim();
}

function classifyWork(value) {
  if (!value || value === "要確認") return "要確認";
  if (/休|休日|週休|年休|代休/.test(value)) return "休日";
  if (/日直|宿直|直/.test(value)) return "日直・当番系";
  return "勤務記号";
}

function inferWorkTime(value) {
  const text = String(value || "");
  const timeMatch = text.match(/\d{1,2}[:：]\d{2}\s*[〜~\-－]\s*\d{1,2}[:：]\d{2}/);
  if (timeMatch) return timeMatch[0].replace("：", ":");
  if (/休|休日|週休|年休|代休/.test(text)) return "休日";
  return "資料の記号を確認（要確認）";
}

function parseEventsFromWorkbook(rows, date) {
  if (!rows.length) return [];
  const dateTokens = buildDateTokens(date, { includeBareDay: false });
  const matchedRows = rows.filter(({ cells }) => cells.some((cell) => dateTokens.some((token) => String(cell).includes(token))));
  const contextRows = matchedRows.flatMap((row) => rows.filter((candidate) => candidate.sheetName === row.sheetName && candidate.rowIndex >= row.rowIndex && candidate.rowIndex <= row.rowIndex + 8));
  const uniqueRows = Array.from(new Map(contextRows.map((row) => [`${row.sheetName}-${row.rowIndex}`, row])).values());

  return uniqueRows
    .map((row) => rowToEvent(row, dateTokens))
    .filter((event) => event.group || event.activity || event.time || event.place)
    .slice(0, 8);
}

function rowToEvent(row, dateTokens) {
  const cells = row.cells.map((cell) => String(cell).trim()).filter(Boolean);
  const meaningful = cells.filter((cell) => !dateTokens.some((token) => cell.includes(token)));
  const time = meaningful.find((cell) => /\d{1,2}[:：]\d{2}|午前|午後|終日/.test(cell)) || "要確認";
  const place = meaningful.find((cell) => /研修室|体育館|食堂|広場|キャンプ|野外|自然|ホール|室|場|館/.test(cell)) || "要確認";
  const activity = meaningful.find((cell) => /活動|研修|入所|退所|体験|講座|説明|式|会|炊飯|登山|散策/.test(cell) && cell !== place) || "要確認";
  const group = meaningful.find((cell) => cell !== time && cell !== place && cell !== activity && cell.length >= 2) || "要確認";
  const note = meaningful.filter((cell) => ![time, place, activity, group].includes(cell)).join(" / ") || "要確認";
  return { group, time, activity, place, note, source: `${row.sheetName} ${row.rowIndex + 1}行目` };
}

function buildWarnings() {
  const warnings = [];
  if (!state.pdfText) warnings.push("勤務予定表PDFが未読み込みです（要確認）");
  if (!state.workbookRows.length) warnings.push("調整プログラムExcelが未読み込みです（要確認）");
  if (state.work?.status === "要確認" || state.work?.division?.includes("検出できません")) warnings.push("西村さんの勤務セルを確定できません（要確認）");
  if (state.work?.source?.includes("簡易抽出")) warnings.push("PDF表構造のため簡易抽出です。勤務記号は原本確認してください");
  if (state.events.some((event) => [event.group, event.time, event.activity, event.place, event.note].includes("要確認"))) warnings.push("調整プログラムに不足または判定できない項目があります（要確認）");
  if (!state.events.length && state.workbookRows.length) warnings.push("今日の日付に一致する団体・行事を検出できません（要確認）");
  return warnings.length ? warnings : ["現時点で自動検出された注意点はありません。最終確認は原本で行ってください"];
}

function render() {
  const dateText = formatJapaneseDate(state.selectedDate);
  elements.todayLabel.textContent = `${dateText} ／ 西村さん`;
  renderWork();
  renderWarnings();
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
    <div class="work-sub">抽出：${escapeHtml(state.work.source)}</div>
  `;
}

function renderWarnings() {
  elements.attentionList.innerHTML = state.warnings.map((warning) => `<li>${escapeHtml(warning)}</li>`).join("");
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
      </div>
    </article>
  `).join("");
}

function renderBasicInfo() {
  const rows = [
    ["勤務", state.work?.status || "PDF未読み込み（要確認）"],
    ["勤務区分", state.work?.division || "要確認"],
    ["食事情報", "資料から自動判定していません（要確認）"],
    ["注意事項", state.warnings[0] || "要確認"],
  ];
  elements.basicInfo.innerHTML = rows.map(([label, value]) => `<div><dt>${escapeHtml(label)}</dt><dd>${escapeHtml(value)}</dd></div>`).join("");
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

function formatJapaneseDate(date) {
  return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日（${WEEKDAYS[date.getDay()]}）`;
}

function toInputDate(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[char]));
}
