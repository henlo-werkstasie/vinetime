// ─────────────────────────────────────────────────────────────────────────────
// VineTime — Google Apps Script Backend
// Paste this entire file into: Extensions > Apps Script > Code.gs
// Deploy as Web App:  Execute as = Me,  Who has access = Anyone
// ─────────────────────────────────────────────────────────────────────────────

var RAW_SHEET = "Submissions";

// ── Entry points ─────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var raw = (e.parameter && e.parameter.data)
              ? e.parameter.data
              : (e.postData ? e.postData.contents : null);

    if (!raw) return respond({ ok: false, error: "No data received. Check the app is sending data correctly." });

    var payload = JSON.parse(raw);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (payload.action === "saveConfig") {
      saveConfig(ss, payload.staff, payload.rates);
      return respond({ ok: true });
    }
    if (payload.action === "addStaff") {
      addStaffMember(ss, payload.name, payload.role);
      return respond({ ok: true });
    }

    var saved = saveSubmission(ss, payload);
    rebuildPayrollTab(ss, payload.fortnightLabel, payload.rates, payload.permanentWorkers, payload.staff);
    return respond({ ok: true, saved: saved });
  } catch (err) {
    return respond({ ok: false, error: err.toString() });
  }
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (action === "submissions") return getSubmissions();
  if (action === "config")      return getConfig();
  if (action === "staff")       return getStaff();
  if (action === "rates")       return getRates();
  if (action === "estates")     return getEstates();
  return respond({ status: "VineTime API OK", timestamp: new Date().toISOString() });
}

function getSubmissions() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(RAW_SHEET);
    if (!sh || sh.getLastRow() < 2) return respond({ ok: true, submissions: [] });
    var data = sh.getDataRange().getValues();
    var hdr = data[0];
    var col = function(name) { return hdr.indexOf(name); };
    var tz = Session.getScriptTimeZone();
    function fmtDate(v) {
      if (!v) return "";
      if (v instanceof Date) return Utilities.formatDate(v, tz, "yyyy-MM-dd");
      return String(v).slice(0, 10);
    }
    function fmtTime(v) {
      if (!v) return "";
      if (v instanceof Date) return Utilities.formatDate(v, tz, "HH:mm");
      var s = String(v).slice(0, 5);
      return s;
    }
    var subs = data.slice(1).map(function(row) {
      var drv = String(row[col("Driver")] || "").trim();
      return {
        id:             String(row[col("ID")]),
        date:           fmtDate(row[col("Date")]),
        supervisor:     String(row[col("Supervisor")] || ""),
        estate:         String(row[col("Estate")] || ""),
        workType:       String(row[col("Work Type")] || ""),
        clockIn:        fmtTime(row[col("Clock In")]),
        clockOut:       fmtTime(row[col("Clock Out")]),
        lunchTaken:     row[col("Lunch Break")] === "Yes",
        hasDriver:      !!drv,
        driverName:     drv,
        driverClockIn:  fmtTime(row[col("Driver In")]),
        driverClockOut: fmtTime(row[col("Driver Out")]),
        driverDrops:    [],
        workers:        String(row[col("Workers")] || "").split(",").map(function(w){return w.trim();}).filter(Boolean),
        notes:          String(row[col("Notes")] || ""),
        submittedAt:    String(row[col("Submitted At")] || "")
      };
    });
    return respond({ ok: true, submissions: subs });
  } catch(err) {
    return respond({ ok: false, error: err.toString() });
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Staff tab — human-readable list, one name per row ────────────────────────

var STAFF_SHEET = "Staff";
var RATES_SHEET = "Rates";

function ensureStaffSheet(ss) {
  var sh = ss.getSheetByName(STAFF_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STAFF_SHEET);
    var hdr = [["Name", "Role"]];
    sh.getRange(1, 1, 1, 2).setValues(hdr)
      .setBackground("#7b1a3a").setFontColor("#ffffff").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 220);
    sh.setColumnWidth(2, 160);
  }
  return sh;
}

function ensureRatesSheet(ss) {
  var sh = ss.getSheetByName(RATES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RATES_SHEET);
    sh.getRange(1, 1, 1, 2).setValues([["Role", "Rate (R/hr)"]])
      .setBackground("#7b1a3a").setFontColor("#ffffff").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 180);
    sh.setColumnWidth(2, 120);
    // Seed default rates
    sh.getRange(2, 1, 4, 2).setValues([
      ["supervisor", 85],
      ["driver",     75],
      ["permanent",  55],
      ["casual",     45]
    ]);
  }
  return sh;
}

function getStaff() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(STAFF_SHEET);
    if (!sh || sh.getLastRow() < 2) return respond({ ok: true, staff: null });
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    var staff = { supervisors: [], drivers: [], permanent: [], casual: [] };
    rows.forEach(function(r) {
      var name = String(r[0]).trim();
      var role = String(r[1]).trim().toLowerCase();
      if (!name) return;
      if (role === "supervisor")  staff.supervisors.push(name);
      else if (role === "driver") staff.drivers.push(name);
      else if (role === "permanent") staff.permanent.push(name);
      else if (role === "casual")    staff.casual.push(name);
    });
    return respond({ ok: true, staff: staff });
  } catch(err) {
    return respond({ ok: false, error: err.toString() });
  }
}

function getRates() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(RATES_SHEET);
    if (!sh || sh.getLastRow() < 2) return respond({ ok: true, rates: null });
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    var rates = {};
    rows.forEach(function(r) {
      var role = String(r[0]).trim().toLowerCase();
      var rate = parseFloat(r[1]) || 0;
      if (role) rates[role] = rate;
    });
    return respond({ ok: true, rates: rates });
  } catch(err) {
    return respond({ ok: false, error: err.toString() });
  }
}

function addStaffMember(ss, name, role) {
  var sh = ensureStaffSheet(ss);
  // Check for duplicate
  if (sh.getLastRow() > 1) {
    var existing = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
      .map(function(r) { return String(r[0]).trim().toLowerCase(); });
    if (existing.indexOf(name.trim().toLowerCase()) >= 0) return;
  }
  sh.appendRow([name.trim(), role]);
}

// ── Config (staff + rates) — read/write a single JSON cell ───────────────────

var CONFIG_SHEET = "Config";

function saveConfig(ss, staff, rates) {
  var sh = ss.getSheetByName(CONFIG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG_SHEET);
    sh.getRange(1, 1).setValue("JSON").setFontWeight("bold")
      .setBackground("#7b1a3a").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 600);
  }
  var payload = JSON.stringify({ staff: staff, rates: rates, savedAt: new Date().toISOString() });
  var lastRow = sh.getLastRow();
  if (lastRow < 2) {
    sh.appendRow([payload]);
  } else {
    sh.getRange(2, 1).setValue(payload);
  }
}

function getConfig() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CONFIG_SHEET);
    if (!sh || sh.getLastRow() < 2) return respond({ ok: true, config: null });
    var raw = sh.getRange(2, 1).getValue();
    if (!raw) return respond({ ok: true, config: null });
    var config = JSON.parse(String(raw));
    return respond({ ok: true, config: config });
  } catch(err) {
    return respond({ ok: false, error: err.toString() });
  }
}

// ── Estates tab ──────────────────────────────────────────────────────────────

var ESTATES_SHEET = "Estates";

function ensureEstatesSheet(ss) {
  var sh = ss.getSheetByName(ESTATES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ESTATES_SHEET);
    sh.getRange(1, 1).setValue("Estate Name").setFontWeight("bold")
      .setBackground("#7b1a3a").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 260);
    // Seed defaults
    var defaults = ["Stellenbosch Ridge","Klein Constantia","De Morgenzon","Babylonstoren","Tokara","Spier"];
    sh.getRange(2, 1, defaults.length, 1).setValues(defaults.map(function(e){return [e];}));
  }
  return sh;
}

function getEstates() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(ESTATES_SHEET);
    if (!sh || sh.getLastRow() < 2) return respond({ ok: true, estates: null });
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    var estates = rows.map(function(r){return String(r[0]).trim();}).filter(Boolean);
    return respond({ ok: true, estates: estates });
  } catch(err) {
    return respond({ ok: false, error: err.toString() });
  }
}

// ── One-time migration: Config tab → Staff + Rates tabs ──────────────────────

function migrateConfigToTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSh = ss.getSheetByName("Config");
  if (!configSh || configSh.getLastRow() < 2) {
    Logger.log("No Config tab found or empty.");
    return;
  }
  var raw = configSh.getRange(2, 1).getValue();
  var config = JSON.parse(String(raw));

  // Build Staff tab
  var staffSh = ensureStaffSheet(ss);
  // Clear existing data rows
  if (staffSh.getLastRow() > 1) staffSh.getRange(2, 1, staffSh.getLastRow() - 1, 2).clearContent();
  var staffRows = [];
  (config.staff.supervisors || []).forEach(function(n) { staffRows.push([n, "supervisor"]); });
  (config.staff.drivers     || []).forEach(function(n) { staffRows.push([n, "driver"]); });
  (config.staff.permanent   || []).forEach(function(n) { staffRows.push([n, "permanent"]); });
  (config.staff.casual      || []).forEach(function(n) { staffRows.push([n, "casual"]); });
  if (staffRows.length) staffSh.getRange(2, 1, staffRows.length, 2).setValues(staffRows);

  // Build Rates tab
  var ratesSh = ensureRatesSheet(ss);
  if (ratesSh.getLastRow() > 1) ratesSh.getRange(2, 1, ratesSh.getLastRow() - 1, 2).clearContent();
  var r = config.rates || {};
  ratesSh.getRange(2, 1, 4, 2).setValues([
    ["supervisor", r.supervisor || 0],
    ["driver",     r.driver     || 0],
    ["permanent",  r.permanent  || 0],
    ["casual",     r.casual     || 0],
  ]);

  Logger.log("Migration complete. " + staffRows.length + " staff, 4 rates written.");
}

// ── Save raw submission row ───────────────────────────────────────────────────

function saveSubmission(ss, payload) {
  var s = payload.submission;
  var sh = ss.getSheetByName(RAW_SHEET);

  if (!sh) {
    sh = ss.insertSheet(RAW_SHEET, 0);
    var hdr = [
      "ID", "Date", "Fortnight", "Supervisor", "Estate", "Work Type",
      "Clock In", "Clock Out", "Lunch Break",
      "Workers", "Driver", "Driver In", "Driver Out", "Notes",
      "Rate Supervisor", "Rate Driver", "Rate Permanent", "Rate Casual",
      "Submitted At"
    ];
    sh.appendRow(hdr);
    sh.getRange(1, 1, 1, hdr.length)
      .setBackground("#7b1a3a").setFontColor("#ffffff").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 120);
    sh.setColumnWidth(10, 220);
  }

  // If this submission ID already exists, delete the old row so we can overwrite it
  var lastRow = sh.getLastRow();
  if (lastRow > 1) {
    var existingIds = sh.getRange(2, 1, lastRow - 1, 1).getValues().map(function(r) { return r[0]; });
    var existingIdx = existingIds.indexOf(s.id);
    if (existingIdx >= 0) sh.deleteRow(existingIdx + 2);
  }

  var r = payload.rates || {};
  sh.appendRow([
    s.id,
    s.date,
    payload.fortnightLabel,
    s.supervisor,
    s.estate,
    s.workType,
    s.clockIn,
    s.clockOut,
    s.lunchTaken === 1 ? "1h" : s.lunchTaken === 0.5 ? "30m" : "No",
    (s.workers || []).join(", "),
    s.driverName || "",
    s.driverClockIn || "",
    s.driverClockOut || "",
    s.notes || "",
    r.supervisor || 0,
    r.driver || 0,
    r.permanent || 0,
    r.casual || 0,
    new Date().toISOString()
  ]);

  // Alternate row shading
  var newRow = sh.getLastRow();
  if (newRow % 2 === 0) {
    sh.getRange(newRow, 1, 1, 19).setBackground("#f9f9f9");
  }

  return true;
}

// ── Hours calculation ─────────────────────────────────────────────────────────

function calcHours(clockIn, clockOut, lunchTaken) {
  if (!clockIn || !clockOut) return 0;
  // Sheet time cells come back as Date objects; extract H:M from them directly.
  // String cells like "06:00" are split normally.
  function toMins(t) {
    if (t instanceof Date) return t.getHours() * 60 + t.getMinutes();
    var p = String(t).split(":").map(Number);
    return p[0] * 60 + (p[1] || 0);
  }
  var mins = toMins(clockOut) - toMins(clockIn);
  var lunchDeduct = (lunchTaken === 1) ? 60 : (lunchTaken === 0.5 || lunchTaken === true || lunchTaken === "Yes") ? 30 : 0;
  mins -= lunchDeduct;
  return Math.max(0, Math.round(mins) / 60);
}

// ── Determine worker role ─────────────────────────────────────────────────────
// permanentWorkers: string[] from payload.permanentWorkers
// staff: full staff object from payload.staff (fallback)

function getRole(name, permanentWorkers, staff) {
  // Supervisors and drivers are identified by the submission fields directly.
  // Here we only need to distinguish permanent vs casual among general workers.
  var permList = permanentWorkers || (staff && staff.permanent) || [];
  var drvList  = (staff && staff.drivers) || [];
  var supList  = (staff && staff.supervisors) || [];
  if (supList.indexOf(name) >= 0) return "supervisor";
  if (drvList.indexOf(name) >= 0) return "driver";
  if (permList.indexOf(name) >= 0) return "permanent";
  return "casual";
}

// ── Format hours as "9h" or "8h 30m" ────────────────────────────────────────

function fmtHours(h) {
  if (!h || h <= 0) return "";
  var totalMins = Math.round(h * 60);
  var hrs = Math.floor(totalMins / 60);
  var mins = totalMins % 60;
  return mins > 0 ? (hrs + "h " + mins + "m") : (hrs + "h");
}

// ── Rebuild the formatted payroll tab for a fortnight ────────────────────────

function rebuildPayrollTab(ss, fortnightLabel, rates, permanentWorkers, staff) {
  try {
    _rebuildPayrollTab(ss, fortnightLabel, rates, permanentWorkers, staff);
  } catch(e) {
    Logger.log("rebuildPayrollTab ERROR: " + e.toString() + "\nStack: " + e.stack);
  }
}

function _rebuildPayrollTab(ss, fortnightLabel, rates, permanentWorkers, staff) {
  var rawSh = ss.getSheetByName(RAW_SHEET);
  if (!rawSh || rawSh.getLastRow() < 2) return;

  var data = rawSh.getDataRange().getValues();
  var hdr  = data[0];
  var col  = function(name) { return hdr.indexOf(name); };

  var rows = data.slice(1).filter(function(r) {
    return r[col("Fortnight")] === fortnightLabel;
  });
  if (!rows.length) return;

  var r = rates || {};
  var tz = Session.getScriptTimeZone();

  // ── Helper: normalise a date value to "YYYY-MM-DD" string ────────────────
  // Sheet cells with dates come back as JS Date objects, not strings.
  function toDateKey(val) {
    if (!val) return "";
    if (val instanceof Date) return Utilities.formatDate(val, tz, "yyyy-MM-dd");
    return String(val).slice(0, 10); // already a string like "2026-06-17"
  }

  // ── Collect all working dates (sorted) ───────────────────────────────────
  var dateSet = {};
  rows.forEach(function(row) {
    var dk = toDateKey(row[col("Date")]);
    if (dk) dateSet[dk] = true;
  });
  var dates = Object.keys(dateSet).sort();

  // ── Build per-person per-day hours map ───────────────────────────────────
  // people[name] = { role, rate, byDate: { "YYYY-MM-DD": hours }, totalHours }
  var people = {};

  function ensurePerson(name, role, rate) {
    if (!name || !String(name).trim()) return;
    name = String(name).trim();
    if (!people[name]) people[name] = { role: role, rate: rate, byDate: {}, totalHours: 0 };
  }

  function addHours(name, date, hours) {
    name = String(name).trim();
    date = String(date);
    people[name].byDate[date] = (people[name].byDate[date] || 0) + hours;
    people[name].totalHours  += hours;
  }

  rows.forEach(function(row) {
    var date     = toDateKey(row[col("Date")]);
    var clockIn  = row[col("Clock In")];
    var clockOut = row[col("Clock Out")];
    var lunchRaw = String(row[col("Lunch Break")] || "");
    var lunch    = lunchRaw === "1h" ? 1 : lunchRaw === "30m" ? 0.5 : 0;
    var hours    = calcHours(clockIn, clockOut, lunch);

    var sup = String(row[col("Supervisor")] || "").trim();
    if (sup) {
      ensurePerson(sup, "supervisor", r.supervisor || Number(row[col("Rate Supervisor")]) || 0);
      addHours(sup, date, hours);
    }

    var workerList = String(row[col("Workers")] || "").split(",")
      .map(function(w) { return w.trim(); }).filter(Boolean);
    workerList.forEach(function(w) {
      var role = (permanentWorkers || []).indexOf(w) >= 0 ? "permanent" : "casual";
      ensurePerson(w, role, r[role] || 0);
      addHours(w, date, hours);
    });

    var drv = String(row[col("Driver")] || "").trim();
    if (drv) {
      var dHours = calcHours(row[col("Driver In")], row[col("Driver Out")], false);
      ensurePerson(drv, "driver", r.driver || Number(row[col("Rate Driver")]) || 0);
      addHours(drv, date, dHours);
    }
  });

  // ── Sort people: supervisor → driver → permanent → casual ────────────────
  var roleOrder = { supervisor: 0, driver: 1, permanent: 2, casual: 3 };
  var names = Object.keys(people).sort(function(a, b) {
    return (roleOrder[people[a].role] || 0) - (roleOrder[people[b].role] || 0);
  });

  // ── Get or create the payroll tab ─────────────────────────────────────────
  var tabName = fortnightLabel.length > 50 ? fortnightLabel.slice(0, 50) : fortnightLabel;
  var tab = ss.getSheetByName(tabName);
  if (!tab) {
    tab = ss.insertSheet(tabName);
    var rawIdx = rawSh.getIndex();
    ss.setActiveSheet(tab);
    ss.moveActiveSheet(rawIdx + 1);
  }
  tab.clear();
  tab.clearFormats();

  // ── Colours ───────────────────────────────────────────────────────────────
  var BURG     = "#7b1a3a";
  var GOLD     = "#6b8e23";
  var LT_GREY  = "#f5f5f5";
  var ROLE_COL = { supervisor: BURG, driver: GOLD, permanent: "#4a7c1f", casual: "#7a8a6a" };
  var ROLE_LBL = { supervisor: "SUPERVISOR", driver: "DRIVER", permanent: "PERMANENT", casual: "CASUAL" };

  var W = 1 + dates.length + 2; // Name & Role col + date cols + Total Hours + Total Pay
  var row = 1;

  // ── Title ─────────────────────────────────────────────────────────────────
  tab.getRange(row, 1, 1, W).merge()
    .setValue("FORTNIGHTLY PAYROLL  —  " + fortnightLabel)
    .setBackground(BURG).setFontColor("#ffffff")
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  tab.setRowHeight(row, 36);
  row++;

  // Subtitle: generated timestamp + submission count + people count
  var grandTotal = 0;
  names.forEach(function(n) { grandTotal += people[n].totalHours * people[n].rate; });
  grandTotal = Math.round(grandTotal * 100) / 100;

  tab.getRange(row, 1, 1, W).merge()
    .setValue(
      "Generated: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy HH:mm") +
      "   |   " + rows.length + " submission(s)   |   " + names.length + " people   |   R " + grandTotal.toFixed(2) + " total"
    )
    .setBackground(BURG).setFontColor("#ffddaa")
    .setFontSize(9).setHorizontalAlignment("center");
  row++;

  // ── Rate badges row ───────────────────────────────────────────────────────
  row++;
  var rateBadges = [
    "Supervisor: R" + (r.supervisor || 0) + "/hr",
    "Driver: R"     + (r.driver     || 0) + "/hr",
    "Permanent: R"  + (r.permanent  || 0) + "/hr",
    "Casual: R"     + (r.casual     || 0) + "/hr"
  ].join("     ");
  tab.getRange(row, 1, 1, W).merge()
    .setValue(rateBadges)
    .setFontSize(9).setFontColor("#555555").setFontStyle("italic");
  row++;
  row++;

  // ── Column header row ─────────────────────────────────────────────────────
  var headerVals = [["Name & Role"]];
  dates.forEach(function(d) {
    var parts = d.split("-");
    var dt = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    headerVals[0].push(Utilities.formatDate(dt, tz, "EEE, d MMM"));
  });
  headerVals[0].push("Total Hours");
  headerVals[0].push("Total Pay (R)");

  var hdrRange = tab.getRange(row, 1, 1, W);
  hdrRange.setValues(headerVals)
    .setBackground(BURG).setFontColor("#ffffff")
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  tab.getRange(row, 1).setHorizontalAlignment("left");
  tab.setRowHeight(row, 28);
  row++;

  // ── Data rows ─────────────────────────────────────────────────────────────
  var dataStartRow = row;
  var colTotalHours = new Array(dates.length).fill(0); // per-date totals

  names.forEach(function(name, i) {
    var p = people[name];
    var rowVals = [name];
    dates.forEach(function(d, di) {
      var h = p.byDate[d] || 0;
      colTotalHours[di] += h;
      rowVals.push(h > 0 ? fmtHours(h) : "·");
    });
    rowVals.push(fmtHours(p.totalHours));
    rowVals.push(Math.round(p.totalHours * p.rate * 100) / 100);

    tab.getRange(row, 1, 1, W).setValues([rowVals])
      .setBackground(i % 2 === 0 ? "#ffffff" : LT_GREY)
      .setFontSize(10).setHorizontalAlignment("center");
    tab.getRange(row, 1).setHorizontalAlignment("left").setFontWeight("bold");

    // Role label in small text below name — write in next row, merge back up using rich text
    // Instead: colour the name cell by role
    tab.getRange(row, 1).setFontColor(ROLE_COL[p.role] || "#333333");

    // Role label — add as a note on the name cell
    tab.getRange(row, 1).setNote(ROLE_LBL[p.role] || "");

    row++;
  });

  // ── Totals row ────────────────────────────────────────────────────────────
  var totalsVals = ["TOTALS"];
  var totalAllHours = 0;
  colTotalHours.forEach(function(h) {
    totalsVals.push(h > 0 ? fmtHours(h) : "");
    totalAllHours += h;
  });
  totalsVals.push(fmtHours(totalAllHours));
  totalsVals.push(Math.round(grandTotal * 100) / 100);

  tab.getRange(row, 1, 1, W).setValues([totalsVals])
    .setBackground("#1a1a1a").setFontColor("#ffffff")
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center");
  tab.getRange(row, 1).setHorizontalAlignment("left");
  tab.setRowHeight(row, 30);

  // Format Total Pay column
  var payCol = W;
  var payRange = tab.getRange(dataStartRow, payCol, names.length + 1, 1);
  payRange.setNumberFormat("R #,##0.00");

  // ── Column widths ─────────────────────────────────────────────────────────
  tab.setColumnWidth(1, 200);
  for (var c = 2; c <= dates.length + 1; c++) tab.setColumnWidth(c, 80);
  tab.setColumnWidth(dates.length + 2, 90);
  tab.setColumnWidth(dates.length + 3, 110);

  // Freeze the header row and name column
  tab.setFrozenRows(dataStartRow - 1);
  tab.setFrozenColumns(1);
}
