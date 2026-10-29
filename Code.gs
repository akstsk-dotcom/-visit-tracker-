/* ==================================================================
   SSPL PR DESK — COMPLETE BACKEND
   Sani Steels Pvt Ltd · Corre Group

   This is the WHOLE backend. Select everything in Code.gs, delete it,
   and paste this in its place. Then:
       1. Save (Ctrl+S)
       2. Run  ->  diagnose        (approve permissions when asked)
       3. Read the log it prints. It tells you which sheets it found.
       4. Deploy > Manage deployments > pencil > New version > Deploy

   NEW IN THIS VERSION: Visits now store RoadKM and KmEstimated (Start Day / Return to SR / Day End
   rows and every visit carry the road km of their leg).
   EARLIER FIX: the tab named in SHEETS below now ALWAYS wins
   (if it has the right header row). Earlier, any other tab with a
   "Specifier" header and data could steal the bills.
   Nothing is ever deleted.
   ================================================================== */

/* ---------------- config ----------------
   These are YOUR actual tab names, confirmed from diagnose output.
   If you ever rename a tab in the sheet, change it here too.        */
var SHEETS = {
  specifiers:   "Master Data",        // your specifier master (headers on row 3)
  transactions: "Entry (Bill-wise)",  // your bill entries
  visits:       "Visits",             // created by this script
  brands:       "Brands"              // optional; brand names in column A
};

var DEFAULT_BRANDS = ["Jaquar", "Kohler", "GLO", "Supreme", "Varmora", "Johnson"];

var PHOTO_FOLDER = "SSPL PR Visit Photos";

var SPEC_HEADERS  = ["Name", "Firm / Studio", "Phone", "City", "Email", "Registered (Y/N)", "Registration Date"];
var TXN_HEADERS   = ["Date", "Bill No", "Branch", "Client / Ledger", "Specifier", "Brand", "Value", "Notes", "Logged At"];
var VISIT_HEADERS = ["VisitID", "Date", "Time", "Timestamp", "Executive", "Branch", "VisitType",
                     "Party", "Firm", "Contact", "Purpose", "Outcome", "Potential", "FollowUp", "Notes",
                     "Latitude", "Longitude", "AccuracyM", "Address", "Locality", "MapLink",
                     "PhotoURL", "Flags", "SyncedAt", "RoadKM", "KmEstimated"];

/* ==================================================================
   ENTRY POINTS
   ================================================================== */
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || "";
    if (action === "visits") return getVisits_(e);
    if (action === "report") return jsonOut_(buildReport_());
    return jsonOut_({ specifiers: listSpecifiers_(), brands: listBrands_() });
  } catch (err) {
    return jsonOut_({ error: String(err), specifiers: [], brands: [] });
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    switch (data.type) {
      case "specifier":        return addSpecifier_(data);
      case "specifier_update": return updateSpecifier_(data);
      case "specifier_delete": return deleteSpecifier_(data);
      case "specifiers_bulk":  return bulkSpecifiers_(data.items || []);
      case "transaction":      return addTransaction_(data);
      case "brand_add":        return addBrand_(data);
      case "brand_delete":     return deleteBrand_(data);
      case "visit":            return addVisit_(data);
      default: return jsonOut_({ ok: false, error: "Unknown entry type: " + data.type });
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ==================================================================
   SHEET PLUMBING — find by headers, create if missing
   ================================================================== */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function norm_(s) { return String(s == null ? "" : s).trim().toLowerCase(); }

/* Find the row that actually holds the column headers. Master sheets often
   carry a title and an instruction line ABOVE it, so scan the first rows
   rather than assuming row 1. Returns the row number, or 0 if not found. */
function headerRowOf_(sh, required) {
  if (!sh || sh.getLastColumn() < 1 || sh.getLastRow() < 1) return 0;
  var scan = Math.min(sh.getLastRow(), 10);
  var block = sh.getRange(1, 1, scan, sh.getLastColumn()).getValues();
  for (var r = 0; r < scan; r++) {
    var hdr = block[r].map(norm_);
    var all = true;
    for (var i = 0; i < required.length; i++) {
      var want = required[i], found = false;
      for (var j = 0; j < hdr.length; j++) {
        // whole-cell or word match, so an instruction sentence never counts
        if (hdr[j] === want || hdr[j].indexOf(want) === 0) { found = true; break; }
      }
      if (!found) { all = false; break; }
    }
    if (all) return r + 1;
  }
  return 0;
}

/* Find the sheet for a given purpose.
   1) The tab named in SHEETS ALWAYS wins, if it has the required header row.
   2) Only if that tab is missing (or has no such header) do we look through
      the other sheets by header, preferring the one that holds data.
   3) If nothing is found, create the tab. */
function findSheet_(configName, required, headers) {
  var ss = ss_();

  // 1) the tab you named always wins
  var named = ss.getSheetByName(configName);
  if (named && headerRowOf_(named, required) > 0) return named;

  // 2) otherwise scan every sheet
  var all = ss.getSheets();
  var cands = [];

  for (var i = 0; i < all.length; i++) {
    var hr = headerRowOf_(all[i], required);
    if (hr > 0) {
      cands.push({
        sh: all[i], hr: hr,
        rows: Math.max(0, all[i].getLastRow() - hr),
        named: all[i].getName() === configName
      });
    }
  }

  // 3) nothing matched — use / create the named tab
  if (!cands.length) {
    return named || createSheet_(configName, headers);
  }

  cands.sort(function (a, b) {
    if ((b.rows > 0 ? 1 : 0) !== (a.rows > 0 ? 1 : 0)) return (b.rows > 0 ? 1 : 0) - (a.rows > 0 ? 1 : 0);
    if (b.named !== a.named) return (b.named ? 1 : 0) - (a.named ? 1 : 0);
    return b.rows - a.rows;
  });
  return cands[0].sh;
}

function createSheet_(name, headers) {
  var sh = ss_().insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#0B2545");
  sh.setFrozenRows(1);
  return sh;
}

/* header name -> column index (0-based), read from the detected header row */
function hdrMap_(sh, hr) {
  var map = {};
  hr = hr || 1;
  if (sh.getLastColumn() < 1 || sh.getLastRow() < hr) return map;
  var hdr = sh.getRange(hr, 1, 1, sh.getLastColumn()).getValues()[0];
  for (var i = 0; i < hdr.length; i++) {
    var k = norm_(hdr[i]);
    if (k && map[k] === undefined) map[k] = i;
  }
  return map;
}

/* pick a column by any of several header fragments */
function col_(map, names) {
  for (var i = 0; i < names.length; i++) {
    var want = names[i];
    for (var k in map) if (k.indexOf(want) !== -1) return map[k];
  }
  return -1;
}

function cell_(row, idx) { return idx >= 0 && idx < row.length ? row[idx] : ""; }

/* Add a column to an existing sheet if it isn't there yet.
   Returns its 0-based index. Existing data and headers are untouched. */
function ensureColumn_(sh, hr, label) {
  var m = hdrMap_(sh, hr);
  var want = norm_(label);
  for (var k in m) if (k.indexOf(want) !== -1) return m[k];
  var idx = sh.getLastColumn();
  sh.getRange(hr, idx + 1).setValue(label).setFontWeight("bold");
  return idx;
}

function specSheet_()  { return findSheet_(SHEETS.specifiers,   ["name"],                SPEC_HEADERS); }
function txnSheet_()   { return findSheet_(SHEETS.transactions, ["specifier"],           TXN_HEADERS); }
function visitSheet_() { return findSheet_(SHEETS.visits,       ["visitid"],             VISIT_HEADERS); }

function specHR_(sh)  { return headerRowOf_(sh, ["name"]) || 1; }
function txnHR_(sh)   { return headerRowOf_(sh, ["specifier"]) || headerRowOf_(sh, ["name"]) || 1; }
function visitHR_(sh) { return headerRowOf_(sh, ["visitid"]) || 1; }

function specCols_(sh) {
  var m = hdrMap_(sh, specHR_(sh));
  return {
    name:  col_(m, ["name"]),
    firm:  col_(m, ["firm", "studio", "company"]),
    phone: col_(m, ["phone", "mobile", "contact"]),
    city:  col_(m, ["city", "town", "location"]),
    email: col_(m, ["email", "mail"]),
    reg:   col_(m, ["registered", "registration status"]),
    regdt: col_(m, ["registration date", "reg date"]),
    width: Math.max(sh.getLastColumn(), SPEC_HEADERS.length)
  };
}

function txnCols_(sh) {
  var m = hdrMap_(sh, txnHR_(sh));
  return {
    date:  col_(m, ["date"]),
    bill:  col_(m, ["bill", "invoice no", "inv no"]),
    branch: col_(m, ["branch", "showroom", "unit"]),
    client: col_(m, ["client", "ledger", "customer", "party"]),
    spec:  col_(m, ["specifier", "name"]),
    brand: col_(m, ["brand", "product"]),
    value: col_(m, ["value", "amount", "sales"]),
    notes: col_(m, ["note", "remark"]),
    width: Math.max(sh.getLastColumn(), TXN_HEADERS.length)
  };
}

/* ==================================================================
   FORMAT HELPERS
   ================================================================== */
function fmtDate_(v) {
  if (v === "" || v == null) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);   // dd/mm/yyyy
  if (m) return m[3] + "-" + pad_(m[2]) + "-" + pad_(m[1]);
  return s.slice(0, 10);
}
function fmtTime_(v) {
  if (v === "" || v == null) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "HH:mm");
  }
  return String(v).slice(0, 5);
}
function pad_(n) { return ("0" + n).slice(-2); }
function isYes_(v) {
  var s = norm_(v);
  return s === "y" || s === "yes" || s === "true" || v === true;
}
/* Stop Sheets turning text that starts with = + - @ (e.g. "-Ravi", "+91 98...")
   into a formula, which shows as #ERROR!. A leading apostrophe keeps it as text. */
function safe_(v) {
  if (typeof v !== "string") return v;
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function num_(v) {
  if (typeof v === "number") return v;
  var n = parseFloat(String(v == null ? "" : v).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? 0 : n;
}

/* ==================================================================
   SPECIFIERS
   ================================================================== */
function readSpecifiers_() {
  var sh = specSheet_(), c = specCols_(sh), hr = specHR_(sh);
  if (sh.getLastRow() <= hr) return { sheet: sh, cols: c, rows: [] };
  var vals = sh.getRange(hr + 1, 1, sh.getLastRow() - hr, sh.getLastColumn()).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var name = String(cell_(vals[i], c.name) || "").trim();
    if (!name) continue;
    out.push({
      row: i + hr + 1,
      name: name,
      firm:  String(cell_(vals[i], c.firm)  || "").trim(),
      phone: String(cell_(vals[i], c.phone) || "").trim(),
      city:  String(cell_(vals[i], c.city)  || "").trim(),
      email: String(cell_(vals[i], c.email) || "").trim(),
      registered: isYes_(cell_(vals[i], c.reg))
    });
  }
  return { sheet: sh, cols: c, rows: out };
}

function listSpecifiers_() {
  return readSpecifiers_().rows.map(function (s) {
    return { name: s.name, firm: s.firm, phone: s.phone, city: s.city,
             email: s.email, registered: s.registered };
  });
}

function writeSpecRow_(sh, c, rowNum, d) {
  var width = Math.max(sh.getLastColumn(), 1);
  var row = sh.getRange(rowNum, 1, 1, width).getValues()[0];
  while (row.length < width) row.push("");
  if (c.name  >= 0) row[c.name]  = safe_(d.name || "");
  if (c.firm  >= 0) row[c.firm]  = safe_(d.firm || "");
  if (c.phone >= 0) row[c.phone] = safe_(d.phone || "");
  if (c.city  >= 0) row[c.city]  = safe_(d.city || "");
  if (c.email >= 0) row[c.email] = safe_(d.email || "");
  if (c.reg   >= 0) row[c.reg]   = d.registered ? "Y" : "N";
  if (c.regdt >= 0 && d.registered && !row[c.regdt]) {
    row[c.regdt] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  sh.getRange(rowNum, 1, 1, width).setValues([row]);
}

function addSpecifier_(d) {
  if (!d.name) return jsonOut_({ ok: false, error: "Name is required." });
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (e) { return jsonOut_({ ok: false, error: "Sheet busy." }); }
  try {
    var s = readSpecifiers_();
    var key = norm_(d.name);
    for (var i = 0; i < s.rows.length; i++) {
      if (norm_(s.rows[i].name) === key) {           // already there — update it
        writeSpecRow_(s.sheet, s.cols, s.rows[i].row, d);
        return jsonOut_({ ok: true, updated: true });
      }
    }
    writeSpecRow_(s.sheet, s.cols, s.sheet.getLastRow() + 1, d);
    return jsonOut_({ ok: true, added: true });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

function updateSpecifier_(d) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (e) { return jsonOut_({ ok: false, error: "Sheet busy." }); }
  try {
    var s = readSpecifiers_();
    var key = norm_(d.originalName || d.name);
    for (var i = 0; i < s.rows.length; i++) {
      if (norm_(s.rows[i].name) === key) {
        writeSpecRow_(s.sheet, s.cols, s.rows[i].row, d);
        if (d.originalName && d.name && d.originalName !== d.name) renameInTxns_(d.originalName, d.name);
        return jsonOut_({ ok: true });
      }
    }
    return jsonOut_({ ok: false, error: "Specifier not found: " + (d.originalName || d.name) });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

function deleteSpecifier_(d) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (e) { return jsonOut_({ ok: false, error: "Sheet busy." }); }
  try {
    var s = readSpecifiers_();
    var key = norm_(d.name);
    for (var i = 0; i < s.rows.length; i++) {
      if (norm_(s.rows[i].name) === key) {
        s.sheet.deleteRow(s.rows[i].row);
        return jsonOut_({ ok: true });
      }
    }
    return jsonOut_({ ok: false, error: "Specifier not found." });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

function bulkSpecifiers_(items) {
  if (!items || !items.length) return jsonOut_({ ok: false, error: "No rows received." });
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return jsonOut_({ ok: false, error: "Sheet busy." }); }
  try {
    var s = readSpecifiers_(), sh = s.sheet, c = s.cols;
    var index = {};
    for (var i = 0; i < s.rows.length; i++) index[norm_(s.rows[i].name)] = s.rows[i].row;

    var width = Math.max(sh.getLastColumn(), SPEC_HEADERS.length);
    var added = 0, updated = 0, newRows = [];

    for (var j = 0; j < items.length; j++) {
      var it = items[j];
      if (!it || !it.name) continue;
      var at = index[norm_(it.name)];
      if (at) { writeSpecRow_(sh, c, at, it); updated++; continue; }

      var row = [];
      for (var k = 0; k < width; k++) row.push("");
      if (c.name  >= 0) row[c.name]  = safe_(it.name);
      if (c.firm  >= 0) row[c.firm]  = safe_(it.firm || "");
      if (c.phone >= 0) row[c.phone] = safe_(it.phone || "");
      if (c.city  >= 0) row[c.city]  = safe_(it.city || "");
      if (c.email >= 0) row[c.email] = safe_(it.email || "");
      if (c.reg   >= 0) row[c.reg]   = it.registered ? "Y" : "N";
      newRows.push(row);
      index[norm_(it.name)] = -1;      // guard against duplicates inside the batch
      added++;
    }
    if (newRows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, newRows.length, width).setValues(newRows);
    }
    return jsonOut_({ ok: true, added: added, updated: updated });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

/* ==================================================================
   TRANSACTIONS (bills)
   ================================================================== */
function readTransactions_() {
  var sh = txnSheet_(), c = txnCols_(sh), hr = txnHR_(sh);
  if (sh.getLastRow() <= hr) return { sheet: sh, cols: c, rows: [] };
  var vals = sh.getRange(hr + 1, 1, sh.getLastRow() - hr, sh.getLastColumn()).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var who = String(cell_(vals[i], c.spec) || "").trim();
    if (!who) continue;
    out.push({
      row: i + hr + 1, specifier: who,
      date:  fmtDate_(cell_(vals[i], c.date)),
      billNo: String(cell_(vals[i], c.bill) || ""),
      branch: String(cell_(vals[i], c.branch) || "").trim(),
      client: String(cell_(vals[i], c.client) || "").trim(),
      brand: String(cell_(vals[i], c.brand) || "").trim(),
      value: num_(cell_(vals[i], c.value)),
      notes: String(cell_(vals[i], c.notes) || "")
    });
  }
  return { sheet: sh, cols: c, rows: out };
}

function addTransaction_(d) {
  if (!d.specifier) return jsonOut_({ ok: false, error: "Specifier is required." });

  // One invoice can carry several brands. Each brand becomes its own row so
  // that brand-wise and scheme totals keep working exactly as before.
  var lines = [];
  if (d.lines && d.lines.length) {
    for (var i = 0; i < d.lines.length; i++) {
      var ln = d.lines[i];
      if (!ln) continue;
      var v = num_(ln.value);
      if (!ln.brand && !v) continue;
      lines.push({ brand: ln.brand || "", value: v });
    }
  } else if (d.brand || num_(d.value)) {
    lines.push({ brand: d.brand || "", value: num_(d.value) });
  }
  if (!lines.length) return jsonOut_({ ok: false, error: "Add at least one brand with a value." });

  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (e) { return jsonOut_({ ok: false, error: "Sheet busy." }); }
  try {
    var sh = txnSheet_(), hr = txnHR_(sh), c = txnCols_(sh);
    if (c.branch < 0) { ensureColumn_(sh, hr, "Branch"); c = txnCols_(sh); }
    if (c.client < 0) { ensureColumn_(sh, hr, "Client / Ledger"); c = txnCols_(sh); }

    var width = Math.max(sh.getLastColumn(), TXN_HEADERS.length);
    var rows = [], total = 0;
    for (var k = 0; k < lines.length; k++) {
      var row = [];
      for (var z = 0; z < width; z++) row.push("");
      if (c.date   >= 0) row[c.date]   = d.date || fmtDate_(new Date());
      if (c.bill   >= 0) row[c.bill]   = safe_(d.billNo || "");
      if (c.branch >= 0) row[c.branch] = d.branch || "";
      if (c.client >= 0) row[c.client] = safe_(d.client || "");
      if (c.spec   >= 0) row[c.spec]   = safe_(d.specifier);
      if (c.brand  >= 0) row[c.brand]  = lines[k].brand;
      if (c.value  >= 0) row[c.value]  = lines[k].value;
      if (c.notes  >= 0) row[c.notes]  = safe_(d.notes || "");
      rows.push(row);
      total += lines[k].value;
    }
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, width).setValues(rows);
    return jsonOut_({ ok: true, rows: rows.length, total: total });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

function renameInTxns_(oldName, newName) {
  var t = readTransactions_();
  if (t.cols.spec < 0) return;
  for (var i = 0; i < t.rows.length; i++) {
    if (norm_(t.rows[i].specifier) === norm_(oldName)) {
      t.sheet.getRange(t.rows[i].row, t.cols.spec + 1).setValue(newName);
    }
  }
}

/* ==================================================================
   BRANDS
   ================================================================== */
function brandSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHEETS.brands);
  if (!sh) {
    sh = createSheet_(SHEETS.brands, ["Brand"]);
    // seed with the defaults plus any brand already used on a bill
    var seed = DEFAULT_BRANDS.slice();
    var t = readTransactions_().rows;
    for (var i = 0; i < t.length; i++) {
      if (t[i].brand && seed.indexOf(t[i].brand) === -1) seed.push(t[i].brand);
    }
    if (seed.length) {
      sh.getRange(2, 1, seed.length, 1).setValues(seed.map(function (b) { return [b]; }));
    }
  }
  return sh;
}

function listBrands_() {
  var sh = brandSheet_();
  var hr = headerRowOf_(sh, ["brand"]) || 1;
  if (sh.getLastRow() <= hr) return DEFAULT_BRANDS;
  var vals = sh.getRange(hr + 1, 1, sh.getLastRow() - hr, 1).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var b = String(vals[i][0] || "").trim();
    if (b && out.indexOf(b) === -1) out.push(b);
  }
  return out.length ? out : DEFAULT_BRANDS;
}

function addBrand_(d) {
  var name = String(d.name || "").trim();
  if (!name) return jsonOut_({ ok: false, error: "Brand name is required." });
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return jsonOut_({ ok: false, error: "Sheet busy." }); }
  try {
    var sh = brandSheet_();
    var current = listBrands_();
    for (var i = 0; i < current.length; i++) {
      if (norm_(current[i]) === norm_(name)) {
        return jsonOut_({ ok: true, duplicate: true, brands: current });
      }
    }
    sh.getRange(sh.getLastRow() + 1, 1).setValue(name);
    return jsonOut_({ ok: true, brands: listBrands_() });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

function deleteBrand_(d) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return jsonOut_({ ok: false, error: "Sheet busy." }); }
  try {
    var sh = brandSheet_();
    var hr = headerRowOf_(sh, ["brand"]) || 1;
    if (sh.getLastRow() <= hr) return jsonOut_({ ok: false, error: "Brand not found." });
    var vals = sh.getRange(hr + 1, 1, sh.getLastRow() - hr, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (norm_(vals[i][0]) === norm_(d.name)) {
        sh.deleteRow(hr + 1 + i);
        return jsonOut_({ ok: true, brands: listBrands_() });
      }
    }
    return jsonOut_({ ok: false, error: "Brand not found." });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

/* ==================================================================
   REPORT  (specifiers with their transactions attached)
   ================================================================== */
function buildReport_() {
  var specs = readSpecifiers_().rows;
  var txns  = readTransactions_().rows;

  var byName = {};
  var out = specs.map(function (s) {
    var o = { name: s.name, firm: s.firm, phone: s.phone, city: s.city,
              email: s.email, registered: s.registered, transactions: [] };
    byName[norm_(s.name)] = o;
    return o;
  });

  for (var i = 0; i < txns.length; i++) {
    var t = txns[i];
    var target = byName[norm_(t.specifier)];
    if (!target) {                      // a bill for someone not in the master list
      target = { name: t.specifier, firm: "", phone: "", city: "", email: "",
                 registered: false, transactions: [] };
      byName[norm_(t.specifier)] = target;
      out.push(target);
    }
    target.transactions.push({ date: t.date, brand: t.brand, value: t.value, billNo: t.billNo, branch: t.branch, client: t.client });
  }

  return { specifiers: out, brands: listBrands_() };
}

/* ==================================================================
   VISITS
   ================================================================== */
function addVisit_(d) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return jsonOut_({ ok: false, error: "Sheet busy." }); }
  try {
    var sh = visitSheet_();

    // idempotent — a queued visit may be retried after a dropped connection
    var vhr = visitHR_(sh);
    if (sh.getLastRow() > vhr) {
      var ids = sh.getRange(vhr + 1, 1, sh.getLastRow() - vhr, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(d.id)) return jsonOut_({ ok: true, duplicate: true });
      }
    }

    var photoUrl = "";
    if (d.photo && String(d.photo).indexOf("data:image") === 0) {
      try { photoUrl = savePhoto_(d.photo, d.id, d.exec, d.party); } catch (e) { photoUrl = ""; }
    }

    // KM columns are added to an existing Visits tab automatically (nothing else is touched)
    var kmCol  = ensureColumn_(sh, vhr, "RoadKM");
    var kmeCol = ensureColumn_(sh, vhr, "KmEstimated");

    sh.appendRow([
      d.id || "", d.date || "", d.time || "", d.ts || "",
      d.exec || "", d.branch || "", d.visitType || "",
      d.party || "", d.firm || "", d.contact || "",
      d.purpose || "", d.outcome || "", num_(d.potential),
      d.followup || "", d.notes || "",
      d.lat || "", d.lng || "", d.acc || "",
      d.address || "", d.locality || "", d.mapLink || "",
      photoUrl, (d.flags || []).join(" | "), new Date()
    ].map(safe_));
    var lastRow = sh.getLastRow();
    if (d.km !== undefined && d.km !== null && d.km !== "") sh.getRange(lastRow, kmCol + 1).setValue(num_(d.km));
    if (d.kmEst) sh.getRange(lastRow, kmeCol + 1).setValue("Y");
    return jsonOut_({ ok: true, photoUrl: photoUrl });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally { lock.releaseLock(); }
}

function savePhoto_(dataUri, id, exec, party) {
  var folders = DriveApp.getFoldersByName(PHOTO_FOLDER);
  var folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(PHOTO_FOLDER);
  var parts = dataUri.split(",");
  var mime  = parts[0].match(/data:([^;]+)/)[1];
  var name  = [id, String(exec || "PR").replace(/\W+/g, "_"),
               String(party || "visit").replace(/\W+/g, "_")].join("_") + ".jpg";
  var blob  = Utilities.newBlob(Utilities.base64Decode(parts[1]), mime, name);
  var file  = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return "https://drive.google.com/uc?id=" + file.getId();
}

function getVisits_(e) {
  var sh = visitSheet_(), hr = visitHR_(sh);
  if (sh.getLastRow() <= hr) return jsonOut_({ visits: [] });

  var hm = hdrMap_(sh, hr);
  var kmIdx  = col_(hm, ["roadkm"]);
  var kmeIdx = col_(hm, ["kmestimated"]);
  var rows = sh.getRange(hr + 1, 1, sh.getLastRow() - hr, Math.max(sh.getLastColumn(), VISIT_HEADERS.length)).getValues();
  var from = (e && e.parameter && e.parameter.from) || "";
  var to   = (e && e.parameter && e.parameter.to)   || "";

  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    var date = fmtDate_(r[1]);
    if (from && date < from) continue;
    if (to && date > to) continue;
    out.push({
      id: String(r[0]), date: date, time: fmtTime_(r[2]),
      ts: r[3] ? new Date(r[3]).toISOString() : "",
      exec: r[4], branch: r[5], type: r[6],
      party: r[7], firm: r[8], contact: String(r[9] || ""),
      purpose: r[10], outcome: r[11], potential: num_(r[12]),
      followup: r[13] ? fmtDate_(r[13]) : "", notes: r[14],
      lat: num_(r[15]) || null, lng: num_(r[16]) || null, acc: num_(r[17]),
      address: r[18], locality: r[19], mapLink: r[20],
      photoUrl: r[21], photo: "",
      flags: r[22] ? String(r[22]).split(" | ").filter(String) : [],
      km: (kmIdx >= 0 && r[kmIdx] !== "" && r[kmIdx] != null) ? num_(r[kmIdx]) : null,
      kmEst: kmeIdx >= 0 ? isYes_(r[kmeIdx]) : false,
      synced: true
    });
  }
  out.sort(function (a, b) { return String(a.ts).localeCompare(String(b.ts)); });
  return jsonOut_({ visits: out });
}

/* ==================================================================
   DIAGNOSE — run this once from the editor before deploying
   ================================================================== */
function diagnose() {
  var L = [];
  L.push("SSPL PR DESK \u2014 BACKEND CHECK");
  L.push("Run at: " + new Date());
  L.push("Spreadsheet: " + ss_().getName());
  L.push("Tabs present: " + ss_().getSheets().map(function (s) { return s.getName(); }).join(", "));
  L.push("");

  var sp = specSheet_(), sc = specCols_(sp), specs = readSpecifiers_().rows;
  L.push("SPECIFIERS  -> tab '" + sp.getName() + "'   " + specs.length + " records");
  L.push("   header row: " + specHR_(sp) + "  (data starts on the row below)");
  L.push("   columns: name=" + sc.name + " firm=" + sc.firm + " phone=" + sc.phone +
         " city=" + sc.city + " email=" + sc.email + " registered=" + sc.reg + "   (-1 = not found)");
  if (specs.length) L.push("   first record: " + specs[0].name + " | " + specs[0].phone);
  else L.push("   *** NO RECORDS FOUND \u2014 this is why the dropdown is empty ***");

  var tx = txnSheet_(), tc = txnCols_(tx), txns = readTransactions_().rows;
  L.push("");
  L.push("BILLS       -> tab '" + tx.getName() + "'   " + txns.length + " records");
  L.push("   header row: " + txnHR_(tx));
  L.push("   columns: date=" + tc.date + " branch=" + tc.branch + " client=" + tc.client +
         " specifier=" + tc.spec + " brand=" + tc.brand + " value=" + tc.value +
         "   (-1 = column will be added when you save the next bill)");
  if (txns.length) L.push("   first record: " + txns[0].date + " | " + txns[0].specifier + " | " + txns[0].value);
  if (tx.getName() !== SHEETS.transactions) {
    L.push("   *** WARNING: bills are going to '" + tx.getName() + "', not '" + SHEETS.transactions + "'. Check the tab name in SHEETS. ***");
  }

  var vs = visitSheet_();
  L.push("");
  L.push("VISITS      -> tab '" + vs.getName() + "'   " + Math.max(0, vs.getLastRow() - visitHR_(vs)) + " records");
  L.push("BRANDS      -> " + listBrands_().join(", "));
  L.push("");
  L.push("If a tab name is wrong, edit SHEETS at the top of this file.");
  L.push("If a column shows -1, rename that column's header in the sheet.");

  var text = L.join("\n");
  Logger.log(text);

  // also write it into a tab so it can be read without opening the log
  try {
    var ss = ss_();
    var out = ss.getSheetByName("_CHECK") || ss.insertSheet("_CHECK");
    out.clear();
    out.getRange(1, 1, L.length, 1).setValues(L.map(function (x) { return [x]; }));
    out.setColumnWidth(1, 760);
    out.getRange(1, 1).setFontWeight("bold");
    ss.setActiveSheet(out);
  } catch (e) {}

  return text;
}