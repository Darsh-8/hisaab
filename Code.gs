// ===== Settings =====
// Your dashboard password. Open the app with  .../exec?key=THIS_VALUE  (letters and numbers only)
const DASHBOARD_KEY = "change-me-to-something-long-and-random";
const SHEET_NAME = "Transactions";

// Optional: public link to your icon, ending in .png (for example a GitHub "Raw" link). Leave "" for none.
const FAVICON_URL = "";

// Header text in row 1  ->  field name. Columns are found by header, so hidden or moved columns don't matter.
const COLUMNS = [
  ["date", "date"], ["month", "month"], ["time", "time"], ["category", "category"],
  ["amount", "amount"], ["payment mode", "paymentMode"], ["remarks", "remarks"],
  ["device name", "deviceName"], ["type", "type"], ["person", "person"]
];
const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];


// ===== iPhone Shortcut sends entries here =====
function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  appendEntry_(data);
  return ContentService.createTextOutput("Success");
}


// ===== Opens the app =====
function doGet(e) {
  const key = (e && e.parameter && e.parameter.key) || "";
  if (key !== DASHBOARD_KEY) {
    return HtmlService.createHtmlOutput(
      "<p style='font-family:sans-serif;padding:24px'>This page needs a key. Add <b>?key=YOUR_KEY</b> to the end of the URL.</p>"
    );
  }
  const page = HtmlService.createTemplateFromFile("Dashboard");
  page.key = key;
  const out = page.evaluate()
    .setTitle("Hisaab")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, viewport-fit=cover")
    .addMetaTag("apple-mobile-web-app-capable", "yes");
  if (FAVICON_URL) out.setFaviconUrl(FAVICON_URL);
  return out;
}


// ===== Called by the app: read all entries =====
function getData(key) {
  checkKey_(key);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const sheet = getSheet_();
  const col = columnMap_(sheet);
  const values = sheet.getDataRange().getValues();
  const get = (r, name) => (col[name] !== undefined && col[name] < r.length) ? r[col[name]] : "";

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r.every(c => c === "")) continue;
    const date = toIsoDate_(get(r, "date"), tz);
    const amount = Number(String(get(r, "amount")).replace(/[^0-9.\-]/g, ""));
    if (!date || !isFinite(amount) || amount === 0) continue;
    rows.push({
      date: date,
      time: toTime_(get(r, "time"), tz),
      category: String(get(r, "category")).trim(),
      amount: amount,
      mode: String(get(r, "paymentMode")).trim(),
      remarks: String(get(r, "remarks")).trim(),
      type: /^credit/i.test(String(get(r, "type")).trim()) ? "Credit" : "Debit",
      person: String(get(r, "person")).trim()
    });
  }
  return rows;
}


// ===== Called by the app: add an entry from the web =====
function addEntry(key, e) {
  checkKey_(key);
  const amount = Number(e.amount);
  if (!(amount > 0)) throw new Error("Amount must be more than zero");
  const p = String(e.date).split("-");                     // "2026-09-27"
  if (p.length !== 3) throw new Error("Invalid date");

  appendEntry_({
    date: p[2] + "/" + p[1] + "/" + p[0],                  // same dd/mm/yyyy format as the Shortcut
    month: MONTH_NAMES[Number(p[1]) - 1],
    time: e.time || "",
    category: e.category || "",
    amount: amount,
    paymentMode: e.paymentMode || "",
    remarks: e.remarks || "",
    deviceName: "Web",
    type: e.type === "Credit" ? "Credit" : "Debit",
    person: e.person || ""
  });
  return true;
}


// ----- helpers -----
function appendEntry_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const col = columnMap_(sheet);
    const width = Math.max(sheet.getLastColumn(), 1);
    const values = sheet.getRange(1, 1, sheet.getMaxRows(), width).getValues();

    let lastDataRow = 1;
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i].some(cell => cell !== "")) { lastDataRow = i + 1; break; }
    }
    const row = lastDataRow + 1;
    if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 50);

    COLUMNS.forEach(([, field]) => {
      if (col[field] === undefined) return;
      let v = data[field];
      if (field === "type") v = v || "Debit";
      sheet.getRange(row, col[field] + 1).setValue(v === undefined || v === null ? "" : v);
    });
  } finally {
    lock.releaseLock();
  }
}

function columnMap_(sheet) {
  const width = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(h => String(h).trim().toLowerCase());
  const map = {};
  COLUMNS.forEach(([header, field], i) => {
    const idx = headers.indexOf(header);
    map[field] = idx >= 0 ? idx : i;                        // fall back to the original position
  });
  return map;
}

function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("No sheet named " + SHEET_NAME);
  return sheet;
}

function checkKey_(key) {
  if (key !== DASHBOARD_KEY) throw new Error("Invalid key");
}

function toIsoDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, "yyyy-MM-dd");
  const m = String(v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + "-" + m[2].padStart(2, "0") + "-" + m[1].padStart(2, "0");
  return null;
}

function toTime_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, "HH:mm");
  return String(v || "").trim();
}
