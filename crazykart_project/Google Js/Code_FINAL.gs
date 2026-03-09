// ════════════════════════════════════════════════════════════════════
// CrazyKart — Google Apps Script  ✅ FINAL VERSION
//
// HOW TO DEPLOY (follow exactly):
//   1. Open Google Sheet "CrazyKart live Sales"
//   2. Click Extensions → Apps Script
//   3. Delete ALL existing code in the editor
//   4. Paste this entire file
//   5. Click Save (💾 icon)
//   6. Click Run → testSheetAccess → check Logs shows ✅
//   7. Click Deploy → New Deployment
//        Type: Web App
//        Execute as: Me (sunilkulali064@gmail.com)
//        Who has access: Anyone
//   8. Copy the /exec URL → paste into index.html APPS_SCRIPT_URL
//   9. Open the /exec URL in browser → should return {"status":"ok"}
// ════════════════════════════════════════════════════════════════════

const SHEET_NAME      = "Sheet1";
const ADMIN_EMAIL     = "sunilkulali064@gmail.com";
const MIN_DAILY_SALES = 5000;

// ── Get sheet safely (works when opened via Extensions → Apps Script) ──
function getSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      const sh = ss.getSheetByName(SHEET_NAME);
      if (sh) return sh;
      return ss.getSheets()[0]; // fallback to first tab
    }
  } catch(e) {
    Logger.log("getActiveSpreadsheet failed: " + e);
  }
  // Fallback for standalone scripts
  try {
    const ID = "19BkYvDJL3t2EADjRRX8RFFqWjiVLL30RldWSOZ3lAE8";
    const ss2 = SpreadsheetApp.openById(ID);
    return ss2.getSheetByName(SHEET_NAME) || ss2.getSheets()[0];
  } catch(e2) {
    Logger.log("openById fallback failed: " + e2);
    return null;
  }
}

// ════════════════════════════════════════════════════════════════════
// doGet — MAIN ENTRY POINT (website sends data here as URL params)
// ════════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const p = e.parameter;

    // Health check — open URL in browser to test
    if (!p || !p.retailer_name) {
      const sheet = getSheet();
      return makeResponse({
        status:      "ok",
        service:     "CrazyKart API",
        sheet_name:  sheet ? sheet.getName() : "NOT FOUND",
        sheet_found: sheet !== null,
        total_rows:  sheet ? sheet.getLastRow() : 0,
        message:     "API running correctly"
      });
    }

    // Parse all fields from URL params
    const data = {
      retailer_name:  (p.retailer_name  || "").trim(),
      store_code:     (p.store_code     || "").trim(),
      brand:          (p.brand          || "").trim(),
      style_code:     (p.style_code     || "").trim(),
      category_group: (p.category_group || "").trim(),
      sub_category:   (p.sub_category   || "").trim(),
      size:           (p.size           || "").trim(),
      colour:         (p.colour         || "").trim(),
      season:         (p.season         || "").trim(),
      slevees:        (p.slevees        || "").trim(),
      core:           (p.core           || "").trim(),
      fit:            (p.fit            || "").trim(),
      mrp:            parseFloat(p.mrp)      || 0,
      qty:            parseInt(p.qty)        || 1,
      discount:       parseFloat(p.discount) || 0,
    };

    const v = validatePayload(data);
    if (!v.ok) return makeResponse({ status: "error", message: v.msg });

    if (checkDuplicate(data)) {
      return makeResponse({ status: "duplicate", message: "Similar entry in last 3 min" });
    }

    const row = buildRow(data);
    appendToSheet(row);
    standardizeLastRow();

    try { checkHighMRP(data); } catch(ex) {}

    Logger.log("✅ " + data.retailer_name + " | Store " + data.store_code +
               " | " + data.style_code + " | ₹" + (data.mrp * data.qty));

    return makeResponse({ status: "success", bill_no: row[2] });

  } catch (err) {
    Logger.log("❌ " + err.toString());
    return makeResponse({ status: "error", message: err.toString() });
  }
}

// doPost for compatibility
function doPost(e) {
  try {
    let p = {};
    if (e.parameter && e.parameter.retailer_name) {
      p = e.parameter;
    } else if (e.postData && e.postData.contents) {
      p = JSON.parse(e.postData.contents);
    }
    return doGet({ parameter: p });
  } catch(err) {
    return makeResponse({ status: "error", message: err.toString() });
  }
}

function makeResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════
// VALIDATION
// ════════════════════════════════════════════════════════════════════
function validatePayload(d) {
  if (!d.retailer_name) return { ok: false, msg: "retailer_name is required" };
  if (!d.store_code)    return { ok: false, msg: "store_code is required" };
  if (!d.brand)         return { ok: false, msg: "brand is required" };
  if (!d.style_code)    return { ok: false, msg: "style_code is required" };
  if (!d.mrp || d.mrp < 1 || d.mrp > 100000)
    return { ok: false, msg: "MRP must be 1–100000" };
  if (!d.qty || d.qty < 1 || d.qty > 999)
    return { ok: false, msg: "Qty must be 1–999" };
  return { ok: true };
}

// ════════════════════════════════════════════════════════════════════
// DUPLICATE CHECK — same store + MRP + qty within 3 minutes
// ════════════════════════════════════════════════════════════════════
function checkDuplicate(data) {
  const sheet = getSheet();
  if (!sheet) return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const cutoff = new Date(Date.now() - 3 * 60 * 1000);
  const checkN = Math.min(20, lastRow - 1);
  const rows   = sheet.getRange(lastRow - checkN + 1, 1, checkN, 23).getValues();

  for (let row of rows) {
    const ts = row[22] instanceof Date ? row[22] : new Date(row[22]);
    if (
      String(row[1]).trim() === String(data.store_code).trim() &&
      parseFloat(row[17])   === data.mrp &&
      parseInt(row[18])     === data.qty &&
      ts > cutoff
    ) return true;
  }
  return false;
}

// ════════════════════════════════════════════════════════════════════
// BUILD ROW — all 23 columns matching Sheet1 exactly
//
// Col 1:  RETAILER_NAME      Col 2:  Store Code
// Col 3:  Bill_No            Col 4:  DATE
// Col 5:  WEEK               Col 6:  FY
// Col 7:  StockNo            Col 8:  BRAND
// Col 9:  STYLE_Code         Col 10: CATEGORY GROUP
// Col 11: SUB CATEGORY       Col 12: SIZE
// Col 13: COLOUR             Col 14: SEASON
// Col 15: slevees            Col 16: CORE/NON CORE
// Col 17: fit                Col 18: MRP
// Col 19: Qty                Col 20: MRP_VALUE
// Col 21: Discount           Col 22: Bill_Val
// Col 23: Time stamp
// ════════════════════════════════════════════════════════════════════
function buildRow(d) {
  const now      = new Date();
  const mrp      = parseFloat(d.mrp);
  const qty      = parseInt(d.qty);
  const discount = parseFloat(d.discount) || 0;
  const mrpValue = mrp * qty;
  const billVal  = Math.max(0, mrpValue - discount);

  return [
    d.retailer_name,          // col 1:  RETAILER_NAME
    d.store_code,             // col 2:  Store Code
    "B-" + now.getTime(),    // col 3:  Bill_No  (auto)
    now,                      // col 4:  DATE
    getWeekNumber(now),       // col 5:  WEEK     (auto)
    getFiscalYear(now),       // col 6:  FY       (auto)
    "STK-" + d.style_code,   // col 7:  StockNo  (auto)
    d.brand,                  // col 8:  BRAND
    d.style_code,             // col 9:  STYLE_Code
    d.category_group,         // col 10: CATEGORY GROUP
    d.sub_category,           // col 11: SUB CATEGORY
    d.size,                   // col 12: SIZE
    d.colour,                 // col 13: COLOUR
    d.season,                 // col 14: SEASON
    d.slevees,                // col 15: slevees
    d.core,                   // col 16: CORE/NON CORE
    d.fit,                    // col 17: fit
    mrp,                      // col 18: MRP
    qty,                      // col 19: Qty
    mrpValue,                 // col 20: MRP_VALUE
    discount,                 // col 21: Discount
    billVal,                  // col 22: Bill_Val
    now,                      // col 23: Time stamp
  ];
}

// ════════════════════════════════════════════════════════════════════
// WRITE TO SHEET
// ════════════════════════════════════════════════════════════════════
function appendToSheet(rowData) {
  const sheet = getSheet();
  if (!sheet) throw new Error("Sheet not found. Check tab name.");

  sheet.appendRow(rowData);

  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 4).setNumberFormat("dd-mmm-yyyy");
  sheet.getRange(newRow, 23).setNumberFormat("dd-mmm-yyyy hh:mm");
  if (newRow % 2 === 0) {
    sheet.getRange(newRow, 1, 1, 23).setBackground("#F8F7FF");
  }
}

function standardizeLastRow() {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  const range   = sheet.getRange(lastRow, 1, 1, 23);
  const vals    = range.getValues()[0];

  vals[0] = toTitleCase(String(vals[0]).trim());   // Retailer → Title Case
  vals[7] = String(vals[7]).trim().toUpperCase();  // Brand    → UPPERCASE
  vals[8] = String(vals[8]).trim().toUpperCase();  // Style    → UPPERCASE

  range.setValues([vals]);
}

// ════════════════════════════════════════════════════════════════════
// SCHEDULED TRIGGERS
// Apps Script → ⏰ Triggers (clock icon, left sidebar)
//   nightlyCleanup     → Day timer  → 11pm–12am
//   checkDailyAlerts   → Day timer  → 8pm–9pm
//   weeklySummaryEmail → Week timer → Monday
// ════════════════════════════════════════════════════════════════════
function nightlyCleanup() {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, 23);
  const vals  = range.getValues();

  for (let i = 0; i < vals.length; i++) {
    let r = vals[i];
    r[0] = toTitleCase(String(r[0]).trim());
    r[7] = String(r[7]).trim().toUpperCase();
    if (!r[4] && r[3]) r[4] = getWeekNumber(new Date(r[3]));
    if (!r[5] && r[3]) r[5] = getFiscalYear(new Date(r[3]));
    if (parseFloat(r[17]) === 0) sheet.getRange(i + 2, 18).setBackground("#FECACA");
    vals[i] = r;
  }
  range.setValues(vals);
  Logger.log("Nightly cleanup done: " + (lastRow - 1) + " rows");
}

function checkDailyAlerts() {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const rows   = sheet.getRange(2, 1, lastRow - 1, 23).getValues();
  const totals = {};

  for (let r of rows) {
    const d = new Date(r[3]);
    d.setHours(0, 0, 0, 0);
    if (d.getTime() !== today.getTime()) continue;
    const key = r[0] + "|" + r[1];
    if (!totals[key]) totals[key] = { retailer: r[0], store: r[1], total: 0, count: 0 };
    totals[key].total += parseFloat(r[21]) || 0;
    totals[key].count++;
  }

  const under = Object.values(totals).filter(s => s.total < MIN_DAILY_SALES);
  if (under.length > 0) {
    let body = "CrazyKart Alert — " + today.toDateString() + "\n\n";
    under.forEach(s => {
      body += "• " + s.retailer + " Store " + s.store +
              ": ₹" + s.total.toLocaleString("en-IN") +
              " (" + s.count + " bills)\n";
    });
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: "[CrazyKart] Stores below ₹" + MIN_DAILY_SALES + " target",
      body: body
    });
  }
}

function weeklySummaryEmail() {
  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 7);

  const rows   = sheet.getRange(2, 1, lastRow - 1, 23).getValues();
  const brands = {}, stores = {};
  let total = 0;

  for (let r of rows) {
    if (new Date(r[3]) < cutoff) continue;
    const v = parseFloat(r[21]) || 0;
    brands[r[7]] = (brands[r[7]] || 0) + v;
    const k = r[0] + " | Store " + r[1];
    stores[k] = (stores[k] || 0) + v;
    total += v;
  }

  const top5 = Object.entries(stores).sort((a, b) => b[1] - a[1]).slice(0, 5);

  let html = "<h2 style='color:#1a1a2e'>CrazyKart Weekly Report</h2>";
  html += "<p><b>Total Revenue (7 days): ₹" + total.toLocaleString("en-IN") + "</b></p><hr/>";
  html += "<h3>Top 5 Stores</h3><ul>";
  top5.forEach(([s, v]) => { html += "<li>" + s + " — ₹" + v.toLocaleString("en-IN") + "</li>"; });
  html += "</ul><h3>By Brand</h3><ul>";
  Object.entries(brands).sort((a, b) => b[1] - a[1]).forEach(([b, v]) => {
    html += "<li>" + b + " — ₹" + v.toLocaleString("en-IN") + "</li>";
  });
  html += "</ul>";

  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: "[CrazyKart] Weekly Summary — ₹" + total.toLocaleString("en-IN"),
    htmlBody: html
  });
}

// ════════════════════════════════════════════════════════════════════
// UTILITY FUNCTIONS
// ════════════════════════════════════════════════════════════════════
function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function getFiscalYear(date) {
  const m = date.getMonth();
  const y = date.getFullYear();
  return m >= 3
    ? y + "-" + String(y + 1).slice(2)
    : (y - 1) + "-" + String(y).slice(2);
}

function toTitleCase(str) {
  return str.replace(/\w\S*/g, t => t.charAt(0).toUpperCase() + t.slice(1).toLowerCase());
}

function checkHighMRP(data) {
  if (parseFloat(data.mrp) > 50000) {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: "[CrazyKart] High MRP Alert",
      body: "Store: " + data.store_code + "\nStyle: " + data.style_code +
            "\nMRP: ₹" + data.mrp + "\nTotal: ₹" + (data.mrp * data.qty)
    });
  }
}

// ════════════════════════════════════════════════════════════════════
// TEST FUNCTION — run this FIRST before deploying
// Click the ▶ Run button with this function selected
// Check Execution Log (bottom panel) for ✅ messages
// ════════════════════════════════════════════════════════════════════
function testSheetAccess() {
  Logger.log("=== CrazyKart Sheet Access Test ===");
  const sheet = getSheet();
  if (!sheet) {
    Logger.log("❌ FAILED — Sheet not found!");
    Logger.log("Make sure tab name is exactly: Sheet1");
    return;
  }
  Logger.log("✅ Sheet found: " + sheet.getName());
  Logger.log("✅ Total rows: " + sheet.getLastRow());
  try {
    const name = SpreadsheetApp.getActiveSpreadsheet().getName();
    Logger.log("✅ Spreadsheet: " + name);
  } catch(e) {
    Logger.log("ℹ️  Standalone mode (no active spreadsheet)");
  }
  Logger.log("✅ All good — safe to deploy!");
}
