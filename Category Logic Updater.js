// CATEGORY LOGIC UPDATER V1.0

// ==========================================
// CATEGORY OVERRIDE UPDATER
// ==========================================
// Scans every inventory job sheet listed in the GSID Database. Any row whose recorded
// Category/Subcategory differs from the algorithm's raw output is saved as a correction
// in the "Category Overrides" tab, where getCategoryLogic checks first on future calls.
//
// Manual Override column (F): check this box on any row to pin it permanently.
// Pinned rows are never modified or removed by this scanner.
function updateCategoryOverrides(silent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Step 1: Get or create "Category Overrides" sheet
  let ovSheet = ss.getSheetByName("Category Overrides");
  if (!ovSheet) {
    ovSheet = ss.insertSheet("Category Overrides");
    ovSheet.getRange(1, 1, 1, 8).setValues([[
      "Key", "BOM ID", "Normalized Description", "Category", "Subcategory",
      "Manual Override", "Source Jobs", "Last Updated"
    ]]);
    ovSheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#e8eaf6");
    ovSheet.setFrozenRows(1);
    ovSheet.setColumnWidth(1, 220);
    ovSheet.setColumnWidth(3, 260);
    ovSheet.setColumnWidth(7, 200);
  }

  // Step 2: Preserve any manually-pinned rows (Manual Override = TRUE)
  const manualOverrides = new Map();
  const existingLastRow = ovSheet.getLastRow();
  if (existingLastRow >= 2) {
    const existing = ovSheet.getRange(2, 1, existingLastRow - 1, 8).getValues();
    for (const row of existing) {
      const key = row[0] ? row[0].toString().trim() : "";
      if (key && row[5] === true) manualOverrides.set(key, [...row]);
    }
  }

  // Step 3: Load GSID for all inventory spreadsheet IDs
  const gsidSheet = ss.getSheetByName("GSID Database");
  if (!gsidSheet) {
    if (!silent) showAlert("Error: GSID Database sheet not found.");
    return;
  }
  const gsidData = gsidSheet.getDataRange().getValues();

  // Step 4: Temporarily disable in-memory overrides so getCategoryLogic returns raw algorithm output
  const savedOverrides = _categoryOverrides;
  _categoryOverrides = new Map();  // empty map = raw algorithm active for this scan

  const detectedMap = new Map();
  let jobsScanned = 0, jobsFailed = 0;

  try {
    for (let r = 1; r < gsidData.length; r++) {
      const jobNum = gsidData[r][0] ? gsidData[r][0].toString().trim().toUpperCase() : "";
      const invId  = gsidData[r][3] ? gsidData[r][3].toString().trim() : "";
      if (!jobNum || !invId || invId === "No sheet found" || invId === "") continue;

      let jobSheet;
      try {
        jobSheet = SpreadsheetApp.openById(invId).getSheetByName(jobNum);
      } catch(e) {
        jobsFailed++;
        continue;
      }
      if (!jobSheet || jobSheet.getLastRow() < 2) continue;

      const lastCol = jobSheet.getLastColumn();
      if (lastCol < 1) continue;
      const headers = sanitizeHeaders(jobSheet.getRange(1, 1, 1, lastCol).getValues()[0]);

      const cBom  = findCol(headers, ["BOMID"]);
      const cDesc = findCol(headers, ["Item Description", "Item"]);
      const cCat  = findCol(headers, ["Category"]);
      const cSub  = findCol(headers, ["Subcategory"]);
      if (cCat === -1) continue;  // can't learn from a sheet with no Category column

      const rows = jobSheet.getRange(2, 1, jobSheet.getLastRow() - 1, lastCol).getValues();
      jobsScanned++;

      for (const row of rows) {
        const bomId    = cBom  > -1 ? row[cBom].toString().trim().toUpperCase() : "";
        const desc     = cDesc > -1 ? row[cDesc].toString().trim() : "";
        const sheetCat = cCat  > -1 ? row[cCat].toString().trim()  : "";
        const sheetSub = cSub  > -1 ? row[cSub].toString().trim()  : "";
        if (!sheetCat || (!bomId && !desc)) continue;

        // Compare against raw algorithm output (overrides disabled above)
        const algo = getCategoryLogic(bomId, desc);
        if (sheetCat === algo.category && (!sheetSub || sheetSub === algo.subcat)) continue;

        const key      = bomId || ("NOBOM_" + normalizeDescription(desc).replace(/[^A-Z0-9]/g, "_"));
        const normDesc = normalizeDescription(desc);

        if (detectedMap.has(key)) {
          const ex = detectedMap.get(key);
          if (!ex.sourceJobs.includes(jobNum)) ex.sourceJobs.push(jobNum);
          // If multiple jobs have the same correction, it's more trustworthy. Last job wins on conflict.
          ex.category = sheetCat;
          ex.subcat   = sheetSub || algo.subcat;
        } else {
          detectedMap.set(key, {
            bomId, normDesc,
            category:   sheetCat,
            subcat:     sheetSub || algo.subcat,
            sourceJobs: [jobNum]
          });
        }
      }
    }
  } finally {
    // Always restore the previous override state, even if scan throws
    _categoryOverrides = savedOverrides;
  }

  // Step 5: Rebuild the overrides sheet (auto-detected + preserved manual rows)
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  const finalRows = [];

  for (const [key, ov] of detectedMap) {
    if (manualOverrides.has(key)) continue;  // manual pins take priority
    finalRows.push([
      key, ov.bomId, ov.normDesc, ov.category, ov.subcat,
      false, ov.sourceJobs.join(", "), dateStr
    ]);
  }
  for (const [, row] of manualOverrides) {
    finalRows.push(row);  // append preserved manual rows as-is
  }

  // Clear old content then write new rows
  if (existingLastRow >= 2) {
    ovSheet.getRange(2, 1, existingLastRow - 1, 8).clearContent().clearDataValidations();
  }
  if (finalRows.length > 0) {
    ovSheet.getRange(2, 1, finalRows.length, 8).setValues(finalRows);
    // Apply checkbox data validation to the Manual Override column
    ovSheet.getRange(2, 6, finalRows.length, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  }

  // Invalidate in-memory cache so the next getCategoryLogic call reloads from the updated sheet
  _categoryOverrides = null;

  if (!silent) {
    showAlert(
      `Category Logic Updated!\n\n` +
      `Job sheets scanned: ${jobsScanned}\n` +
      `Overrides detected: ${detectedMap.size}\n` +
      `Manual overrides preserved: ${manualOverrides.size}` +
      (jobsFailed > 0 ? `\nSheets unavailable: ${jobsFailed}` : "")
    );
  }
}

// ==========================================
// DAILY CATEGORY LOGIC TRIGGER
// ==========================================
// Registers a 3 AM nightly time-based trigger for updateCategoryOverrides, but only if one
// doesn't already exist. Called from onOpen() so the trigger self-heals if ever deleted.
function ensureDailyCategoryTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "updateCategoryOverrides") return;
  }
  ScriptApp.newTrigger("updateCategoryOverrides")
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}
