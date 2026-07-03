// CATEGORY LOGIC UPDATER V2.0
//
// Two-phase nightly pipeline:
//   2 AM — applyAndPromoteOverrides()
//          Reads the "Category Overrides" inbox, pushes corrections back to the specific
//          source job rows where each mismatch was detected, promotes all entries to the
//          permanent "Learned Rules" tab, then wipes the inbox.
//
//   3 AM — updateCategoryOverrides()
//          Scans every inventory job sheet in the GSID. Compares each row against
//          getCategoryLogic (which now includes Learned Rules). Any row whose Category/
//          Subcategory still doesn't match lands in the "Category Overrides" inbox for review.

// ==========================================
// COLUMN LAYOUT (both Category Overrides and Learned Rules share the same 7-column schema)
// ==========================================
// A: Key               — BOM ID if present; else "NOBOM_" + normalized description
// B: BOM ID            — raw BOM ID (uppercase)
// C: Normalized Desc   — output of normalizeDescription()
// D: Category          — correct category value
// E: Subcategory       — correct subcategory value
// F: Source Jobs       — comma-separated job numbers where the mismatch was found
// G: Last Updated      — ISO date the row was written or promoted

const COV_COLS = 7;

// ==========================================
// SHEET SETUP HELPERS
// ==========================================
function _getOrCreateOverridesSheet(ss) {
  let sheet = ss.getSheetByName("Category Overrides");
  if (!sheet) {
    sheet = ss.insertSheet("Category Overrides");
    _initSheetHeaders(sheet, "#fff8e1");
  } else {
    _migrateOldFormat(sheet);
  }
  return sheet;
}

function _getOrCreateLearnedRulesSheet(ss) {
  let sheet = ss.getSheetByName("Learned Rules");
  if (!sheet) {
    sheet = ss.insertSheet("Learned Rules");
    _initSheetHeaders(sheet, "#e8f5e9");
  }
  return sheet;
}

function _initSheetHeaders(sheet, bgColor) {
  const headers = [["Key", "BOM ID", "Normalized Description", "Category", "Subcategory", "Source Jobs", "Last Updated"]];
  sheet.getRange(1, 1, 1, COV_COLS).setValues(headers).setFontWeight("bold").setBackground(bgColor);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(3, 260);
  sheet.setColumnWidth(6, 200);
  _applyHeaderNotes(sheet);
}

function _applyHeaderNotes(sheet) {
  sheet.getRange("A1").setNote(
    "Match key used by getCategoryLogic.\n" +
    "• BOM ID (if the item has one)\n" +
    "• NOBOM_<normalized description> (for items with no BOM ID)\n" +
    "This key is how corrections are linked back to job sheet rows."
  );
  sheet.getRange("D1").setNote(
    "Corrected Category value.\n" +
    "getCategoryLogic returns this instead of its algorithm output for any row whose key matches."
  );
  sheet.getRange("E1").setNote("Corrected Subcategory value paired with the Category correction.");
  sheet.getRange("F1").setNote(
    "Job numbers where this mismatch was originally detected (comma-separated).\n" +
    "During the 2 AM apply run, only rows in these specific jobs are updated — " +
    "not every job in the GSID. This prevents correcting a different client's inventory."
  );
  sheet.getRange("G1").setNote("Timestamp this row was last written or promoted.");
}

function _migrateOldFormat(sheet) {
  // If the old 8-column layout with 'Manual Override' in column F still exists, remove that column
  if (sheet.getLastColumn() >= 6) {
    const fHeader = sheet.getRange(1, 6).getValue().toString().trim();
    if (fHeader === "Manual Override") {
      sheet.deleteColumn(6);
    }
  }
  // Ensure notes are current
  _applyHeaderNotes(sheet);
}

// ==========================================
// 2 AM: APPLY OVERRIDES → JOB SHEETS + PROMOTE → LEARNED RULES
// ==========================================
function applyAndPromoteOverrides(silent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ovSheet = _getOrCreateOverridesSheet(ss);

  if (ovSheet.getLastRow() < 2) {
    if (!silent) showAlert("Category Overrides is empty — nothing to apply.");
    return;
  }

  const ovData = ovSheet.getRange(2, 1, ovSheet.getLastRow() - 1, COV_COLS).getValues()
    .filter(r => r[0].toString().trim() !== "");  // skip blank rows

  if (ovData.length === 0) {
    if (!silent) showAlert("Category Overrides is empty — nothing to apply.");
    return;
  }

  // Build a job → inventory SS ID map from GSID (read once)
  const gsidSheet = ss.getSheetByName("GSID Database");
  const invIdByJob = new Map();
  if (gsidSheet) {
    const gsidData = gsidSheet.getDataRange().getValues();
    for (let r = 1; r < gsidData.length; r++) {
      const job   = gsidData[r][0] ? gsidData[r][0].toString().trim().toUpperCase() : "";
      const invId = gsidData[r][3] ? gsidData[r][3].toString().trim() : "";
      if (job && invId && invId !== "No sheet found") invIdByJob.set(job, invId);
    }
  }

  // Step 1: Push corrections to source job rows only
  let rowsUpdated = 0, jobsUpdated = new Set();

  for (const row of ovData) {
    const key        = row[0].toString().trim();
    const bomId      = row[1].toString().trim().toUpperCase();
    const newCat     = row[3].toString().trim();
    const newSub     = row[4].toString().trim();
    const sourceJobs = row[5].toString().split(",").map(j => j.trim().toUpperCase()).filter(Boolean);

    if (!key || !newCat || sourceJobs.length === 0) continue;

    for (const jobNum of sourceJobs) {
      const invId = invIdByJob.get(jobNum);
      if (!invId) continue;

      let jobSheet;
      try { jobSheet = SpreadsheetApp.openById(invId).getSheetByName(jobNum); }
      catch(e) { continue; }
      if (!jobSheet || jobSheet.getLastRow() < 2) continue;

      const lastCol = jobSheet.getLastColumn();
      const headers = sanitizeHeaders(jobSheet.getRange(1, 1, 1, lastCol).getValues()[0]);
      const cBom  = findCol(headers, ["BOMID"]);
      const cDesc = findCol(headers, ["Item Description", "Item"]);
      const cCat  = findCol(headers, ["Category"]);
      const cSub  = findCol(headers, ["Subcategory"]);
      if (cCat === -1) continue;

      const rows = jobSheet.getRange(2, 1, jobSheet.getLastRow() - 1, lastCol).getValues();

      // Collect row indices that need updating (A1-notation row = data index + 2)
      const catRanges = [], subRanges = [];
      for (let i = 0; i < rows.length; i++) {
        const rowBom  = cBom  > -1 ? rows[i][cBom].toString().trim().toUpperCase() : "";
        const rowDesc = cDesc > -1 ? rows[i][cDesc].toString().trim() : "";
        const rowKey  = rowBom || ("NOBOM_" + normalizeDescription(rowDesc).replace(/[^A-Z0-9]/g, "_"));
        if (rowKey !== key) continue;

        const curCat = rows[i][cCat] ? rows[i][cCat].toString().trim() : "";
        const curSub = cSub > -1 && rows[i][cSub] ? rows[i][cSub].toString().trim() : "";
        if (curCat === newCat && curSub === newSub) continue;  // already correct

        const sheetRow = i + 2;
        catRanges.push(jobSheet.getRange(sheetRow, cCat + 1).getA1Notation());
        if (cSub > -1) subRanges.push(jobSheet.getRange(sheetRow, cSub + 1).getA1Notation());
        rowsUpdated++;
      }

      if (catRanges.length > 0) {
        jobSheet.getRangeList(catRanges).setValue(newCat);
        if (subRanges.length > 0) jobSheet.getRangeList(subRanges).setValue(newSub);
        jobsUpdated.add(jobNum);
      }
    }
  }

  // Step 2: Promote all override entries to Learned Rules (permanent layer)
  _promoteToLearnedRules(ss, ovData);

  // Step 3: Wipe the override inbox
  ovSheet.getRange(2, 1, Math.max(ovSheet.getLastRow() - 1, 1), COV_COLS).clearContent();

  // Invalidate in-memory caches so next getCategoryLogic picks up the new Learned Rules
  _learnedRules     = null;
  _categoryOverrides = null;

  if (!silent) {
    showAlert(
      `Overrides Applied!\n\n` +
      `Job sheets updated: ${jobsUpdated.size}\n` +
      `Rows corrected: ${rowsUpdated}\n` +
      `Rules promoted to Learned Rules: ${ovData.length}`
    );
  }
}

function _promoteToLearnedRules(ss, ovData) {
  const lrSheet = _getOrCreateLearnedRulesSheet(ss);
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

  // Build existing key → sheet row map so we can update in-place rather than duplicate
  const existing = new Map();
  if (lrSheet.getLastRow() >= 2) {
    const existingData = lrSheet.getRange(2, 1, lrSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < existingData.length; i++) {
      const k = existingData[i][0] ? existingData[i][0].toString().trim() : "";
      if (k) existing.set(k, i + 2);
    }
  }

  const newRows = [];
  for (const row of ovData) {
    const key = row[0] ? row[0].toString().trim() : "";
    if (!key) continue;
    const promoted = [row[0], row[1], row[2], row[3], row[4], row[5], dateStr];
    if (existing.has(key)) {
      lrSheet.getRange(existing.get(key), 1, 1, COV_COLS).setValues([promoted]);
    } else {
      newRows.push(promoted);
    }
  }
  if (newRows.length > 0) {
    const startRow = lrSheet.getLastRow() + 1;
    lrSheet.getRange(startRow, 1, newRows.length, COV_COLS).setValues(newRows);
  }
}

// ==========================================
// 3 AM: SCAN JOB SHEETS → POPULATE OVERRIDE INBOX
// ==========================================
function updateCategoryOverrides(silent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ovSheet = _getOrCreateOverridesSheet(ss);

  // Disable only the temporary override inbox during the scan — Learned Rules stay active
  // because they represent the permanent "correct" answer. If a job sheet matches a learned
  // rule, that's not a discrepancy. Only items where neither learned rules nor the algorithm
  // produce the right answer should land in the inbox.
  const savedOV = _categoryOverrides;
  _categoryOverrides = new Map();

  const gsidSheet = ss.getSheetByName("GSID Database");
  if (!gsidSheet) {
    _learnedRules      = savedLR;
    _categoryOverrides = savedOV;
    if (!silent) showAlert("Error: GSID Database sheet not found.");
    return;
  }
  const gsidData = gsidSheet.getDataRange().getValues();

  const detectedMap = new Map();
  let jobsScanned = 0, jobsFailed = 0;

  try {
    for (let r = 1; r < gsidData.length; r++) {
      const jobNum = gsidData[r][0] ? gsidData[r][0].toString().trim().toUpperCase() : "";
      const invId  = gsidData[r][3] ? gsidData[r][3].toString().trim() : "";
      if (!jobNum || !invId || invId === "No sheet found" || invId === "") continue;

      let jobSheet;
      try { jobSheet = SpreadsheetApp.openById(invId).getSheetByName(jobNum); }
      catch(e) { jobsFailed++; continue; }
      if (!jobSheet || jobSheet.getLastRow() < 2) continue;

      const lastCol = jobSheet.getLastColumn();
      if (lastCol < 1) continue;
      const headers = sanitizeHeaders(jobSheet.getRange(1, 1, 1, lastCol).getValues()[0]);
      const cBom  = findCol(headers, ["BOMID"]);
      const cDesc = findCol(headers, ["Item Description", "Item"]);
      const cCat  = findCol(headers, ["Category"]);
      const cSub  = findCol(headers, ["Subcategory"]);
      if (cCat === -1) continue;

      const rows = jobSheet.getRange(2, 1, jobSheet.getLastRow() - 1, lastCol).getValues();
      jobsScanned++;

      for (const row of rows) {
        const bomId    = cBom  > -1 ? row[cBom].toString().trim().toUpperCase() : "";
        const desc     = cDesc > -1 ? row[cDesc].toString().trim() : "";
        const sheetCat = cCat  > -1 ? row[cCat].toString().trim()  : "";
        const sheetSub = cSub  > -1 ? row[cSub].toString().trim()  : "";
        if (!sheetCat || (!bomId && !desc)) continue;

        // Compare against raw algorithm (both rule tables are empty above)
        const algo = getCategoryLogic(bomId, desc);
        if (sheetCat === algo.category && (!sheetSub || sheetSub === algo.subcat)) continue;

        const key      = bomId || ("NOBOM_" + normalizeDescription(desc).replace(/[^A-Z0-9]/g, "_"));
        const normDesc = normalizeDescription(desc);

        if (detectedMap.has(key)) {
          const ex = detectedMap.get(key);
          if (!ex.sourceJobs.includes(jobNum)) ex.sourceJobs.push(jobNum);
          ex.category = sheetCat;
          ex.subcat   = sheetSub || algo.subcat;
        } else {
          detectedMap.set(key, { bomId, normDesc, category: sheetCat, subcat: sheetSub || algo.subcat, sourceJobs: [jobNum] });
        }
      }
    }
  } finally {
    _categoryOverrides = savedOV;
  }

  // Rebuild the override inbox
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  const finalRows = [];
  for (const [key, ov] of detectedMap) {
    finalRows.push([key, ov.bomId, ov.normDesc, ov.category, ov.subcat, ov.sourceJobs.join(", "), dateStr]);
  }

  const clearRows = Math.max(ovSheet.getLastRow() - 1, 1);
  ovSheet.getRange(2, 1, clearRows, COV_COLS).clearContent();
  if (finalRows.length > 0) {
    ovSheet.getRange(2, 1, finalRows.length, COV_COLS).setValues(finalRows);
  }

  _categoryOverrides = null;

  if (!silent) {
    showAlert(
      `Category Logic Scan Complete!\n\n` +
      `Job sheets scanned: ${jobsScanned}\n` +
      `New discrepancies found: ${detectedMap.size}` +
      (jobsFailed > 0 ? `\nSheets unavailable: ${jobsFailed}` : "")
    );
  }
}

// ==========================================
// DAILY TRIGGERS
// ==========================================
function ensureDailyApplyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "applyAndPromoteOverrides") return;
  }
  ScriptApp.newTrigger("applyAndPromoteOverrides").timeBased().everyDays(1).atHour(2).create();
}

function ensureDailyCategoryTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "updateCategoryOverrides") return;
  }
  ScriptApp.newTrigger("updateCategoryOverrides").timeBased().everyDays(1).atHour(3).create();
}
