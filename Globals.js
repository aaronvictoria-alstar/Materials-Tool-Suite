// GLOBALS V1.02

// ==========================================
// MASTER UI MENU TRIGGER
// ==========================================
// Builds the dropdown menu in the Google Sheet UI for manual admin triggers like GSID updates or history healing.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ Admin Tools')
    .addItem('Update GSID Database', 'updateGSIDDatabase')
    .addItem('Update Category Logic', 'updateCategoryOverrides')
    .addItem('Bidirectional Heal (MMT ⇆ KT History)', 'showHealModal')
    .addToUi();

  // Programmatically guarantee the nightly background triggers are active
  try { ensureDailyGSIDTrigger(); } catch (e) {}
  try { ensureDailyCategoryTrigger(); } catch (e) {}
}

// ==========================================
// MASTER ON EDIT TRIGGER (ROUTES RT & PT)
// ==========================================
// Acts as a traffic controller. Listens to UI clicks and safely routes checkbox toggles to the correct visibility filter or sorter based on the active sheet.
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const editedRange = e.range;

  // ----------------------------------------
  // ROUTE 1: PIVOT TOOL (WMS DASHBOARD)
  // ----------------------------------------
  if (sheetName === "Pivot Tool") {
    let checkboxRange;
    try { checkboxRange = sheet.getRange("PT_input_balanced"); } catch(err) {}

    if (checkboxRange && editedRange.getRow() === checkboxRange.getRow() && editedRange.getColumn() === checkboxRange.getColumn()) {
      const isShowBalanced = (String(e.value).toUpperCase() === "TRUE");
      applyBalancedFilter(sheet, isShowBalanced);
    }
    return; // Exit after handling PT
  }

  // ----------------------------------------
  // ROUTE 2: RECEIVING TOOL (RT)
  // ----------------------------------------
  if (sheetName === RT_SHEET_NAME) {
    let pipeInputRange, completeInputRange, sortInputRange;
    try { pipeInputRange     = sheet.getRange("RT_input_pipe"); } catch (err) {}
    try { completeInputRange = sheet.getRange("RT_input_complete"); } catch (err) {}
    try { sortInputRange     = sheet.getRange("RT_input_orderBy"); } catch (err) {}

    const isPipeToggle = pipeInputRange &&
      editedRange.getRow()    === pipeInputRange.getRow() &&
      editedRange.getColumn() === pipeInputRange.getColumn();
    const isCompleteToggle = completeInputRange &&
      editedRange.getRow()    === completeInputRange.getRow() &&
      editedRange.getColumn() === completeInputRange.getColumn();
    const isSortToggle = sortInputRange &&
      editedRange.getRow()    === sortInputRange.getRow() &&
      editedRange.getColumn() === sortInputRange.getColumn();
    if (isPipeToggle || isCompleteToggle || isSortToggle) {
      const lock = LockService.getScriptLock();
      try {
        if (!lock.tryLock(3000)) return;
        if (isPipeToggle || isCompleteToggle) {
          applyMasterFilters(sheet);
        } else if (isSortToggle) {
          fastSortRtTable(sheet);
          applyMasterFilters(sheet);
        }

      } finally {
        lock.releaseLock();
      }
    }
    return;
  }
}

// ==========================================
// GLOBAL HELPER: QCPR STATS FETCHER
// ==========================================
// Connects to the QCPR file to aggregate total inches, priority distributions, and main NPS sizes for Kitting Batches.
function fetchQcprStats(jobNum, globalDrawings) {
  let cSpool = 0, cPri = 6, cInch = 9, cSize = 12;
  const coords = getJobCoordinatesFromGSID(jobNum);
  if (coords) {
    if (coords.qcprSpool) cSpool = coords.qcprSpool.colIdx;
    if (coords.qcprInch) cInch = coords.qcprInch.colIdx;
    if (coords.qcprPriority) cPri = coords.qcprPriority.colIdx;
    if (coords.qcprSize) cSize = coords.qcprSize.colIdx;
  }

  let totalInches = 0, matchesFound = 0;
  const priorityCounts = {}, sizeCounts = {};

  try {
    const qcprSS = getQcprSpreadsheet(jobNum);
    if (!qcprSS) {
      SpreadsheetApp.getActiveSpreadsheet().toast("Could not locate a QCPR file in Drive for Job: " + jobNum, "QCPR Warning", 5);
      return { totalInches, priorityCounts, sizeCounts };
    }

    const fabSheet = qcprSS.getSheetByName("Fab Data");
    if (!fabSheet) {
      SpreadsheetApp.getActiveSpreadsheet().toast("QCPR File found, but missing the 'Fab Data' tab.", "QCPR Warning", 5);
      return { totalInches, priorityCounts, sizeCounts };
    }

    const fLastRow = fabSheet.getLastRow();
    const fLastCol = fabSheet.getLastColumn();
    if (fLastRow >= 3) {
      const fabData = fabSheet.getRange(1, 1, fLastRow, fLastCol).getValues();
      for (let i = 2; i < fabData.length; i++) {
        const rawDraw   = fabData[i][cSpool] ? fabData[i][cSpool].toString().toUpperCase().trim() : "";
        const cleanDraw = cleanDrawingNumber(rawDraw);

        if (globalDrawings.has(rawDraw) || globalDrawings.has(cleanDraw)) {
          matchesFound++;
          const pri = fabData[i][cPri] ? fabData[i][cPri].toString().trim() : "";
          if (pri) priorityCounts[pri] = (priorityCounts[pri] || 0) + 1;
          totalInches += parseFloat(fabData[i][cInch]) || 0;

          const size = fabData[i][cSize] ? fabData[i][cSize].toString().replace(/"/g, '').trim() : "";
          if (size) sizeCounts[size] = (sizeCounts[size] || 0) + 1;
        }
      }

      if (matchesFound === 0) {
        SpreadsheetApp.getActiveSpreadsheet().toast("QCPR checked, but no drawing numbers matched.", "QCPR Warning", 5);
      }
    }
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Error reading QCPR: " + e.toString(), "QCPR Error", 5);
  }

  return { totalInches, priorityCounts, sizeCounts };
}

// ==========================================
// GLOBAL HELPER: SHARED RAW DATA PULLER
// ==========================================
// Automatically fetches historical inventory data from the Client's Job Sheet and maps it to the local RT/PT Data tabs.
function updateRawDataTab(jobSheetName, clientName, rawSheetName, silent = true) {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const rawDataSheet = ss.getSheetByName(rawSheetName);

  if (!rawDataSheet) {
    if (!silent) showAlert(`Error: Tab '${rawSheetName}' missing.`);
    return [];
  }

  const targetSS = getInventorySpreadsheet(clientName, jobSheetName);
  if (!targetSS) return [];

  const targetSheet = targetSS.getSheetByName(jobSheetName);
  const ptLastCol   = rawDataSheet.getLastColumn() || 14;
  const rawLastRow  = Math.max(rawDataSheet.getLastRow(), 2);

  rawDataSheet.getRange(2, 1, rawLastRow, ptLastCol).clearContent();
  if (!targetSheet || targetSheet.getLastRow() < 2) {
    if (!silent) showAlert(`Notice: No historical inventory found for Job '${jobSheetName}'.`);
    return [];
  }

  const tLastRow = targetSheet.getLastRow();
  const tLastCol = targetSheet.getLastColumn();
  const tHeaders = sanitizeHeaders(targetSheet.getRange(1, 1, 1, tLastCol).getValues()[0]);
  let ptHeaders;
  try { ptHeaders = sanitizeHeaders(rawDataSheet.getRange(1, 1, 1, ptLastCol).getValues()[0]); }
  catch (e) { ptHeaders = tHeaders; }

  const inventoryData = targetSheet.getRange(2, 1, tLastRow - 1, tLastCol).getValues();
  try {
    const mappedData = inventoryData.map(row => {
      const newRow = new Array(ptLastCol).fill("");
      ptHeaders.forEach((ptHead, i) => {
        if (!ptHead) return;
        const aliases = HEADER_ALIASES[ptHead] || [ptHead];
        const sourceIdx = findCol(tHeaders, aliases);
        if (sourceIdx > -1) newRow[i] = row[sourceIdx];
      });
      return newRow;
    });
    rawDataSheet.getRange(2, 1, mappedData.length, ptLastCol).setValues(mappedData);
    if (!silent) showAlert(`Success! Loaded ${mappedData.length} lines of historical inventory.`);
    return [ptHeaders, ...mappedData];
  } catch (e) {
    showAlert(`Error updating ${rawSheetName}: ${e.message}`);
    return [];
  }
}

// ==========================================
// GLOBAL HELPER: DRIVE & CACHE
// ==========================================
// Handles Google Drive searches for MMTs, QCPRs, and Inventory sheets, utilizing CacheService to dramatically speed up subsequent searches.
function _driveFetch(cacheKey, query, findLatest = false) {
  const cache  = CacheService.getScriptCache();
  const fileId = cache.get(cacheKey);
  if (fileId) {
    try { return SpreadsheetApp.openById(fileId); }
    catch (e) { cache.remove(cacheKey); }
  }

  const files = DriveApp.searchFiles(query);
  let latestFile = null, latestTime = 0;
  while (files.hasNext()) {
    const file = files.next();
    if (findLatest) {
      if (!query.includes("Master Material Tracker") && file.getName().includes("Master Material Tracker")) continue;
      const fileTime = file.getLastUpdated().getTime();
      if (fileTime > latestTime) { latestTime = fileTime; latestFile = file; }
    } else {
      latestFile = file; break;
    }
  }

  if (latestFile) {
    cache.put(cacheKey, latestFile.getId(), 21600);
    return SpreadsheetApp.openById(latestFile.getId());
  }
  return null;
}

function getSpreadsheetFromGSID(jobNum, colIdx) {
  const cleanJob = jobNum.toString().trim().toUpperCase();
  const cacheKey = "GSID_COL" + colIdx + "_" + cleanJob;
  const cache    = CacheService.getScriptCache();

  const cachedId = cache.get(cacheKey);
  if (cachedId) {
    try { return SpreadsheetApp.openById(cachedId); }
    catch (e) { cache.remove(cacheKey); }
  }

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const gsidSheet = ss.getSheetByName("GSID Database");
  if (!gsidSheet) return null;

  const data = gsidSheet.getDataRange().getDisplayValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() === cleanJob) {
      const fileId = data[i][colIdx] ? data[i][colIdx].toString().trim() : "";
      if (fileId === "No sheet found" || fileId === "") return null;
      try {
        const foundSS = SpreadsheetApp.openById(fileId);
        cache.put(cacheKey, fileId, 21600);
        return foundSS;
      } catch (e) { return null; }
    }
  }
  return null;
}

function autoDiscoverAndAddJob(cleanJob) {
  const mmtQuery = `title contains '${cleanJob}' and title contains 'Master Material Tracker' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`;
  const mmtFiles = DriveApp.searchFiles(mmtQuery);
  let mmtId = "", mmtTime = 0;
  while (mmtFiles.hasNext()) {
    const file  = mmtFiles.next();
    const fTime = file.getLastUpdated().getTime();
    if (fTime > mmtTime) { mmtTime = fTime; mmtId = file.getId(); }
  }

  const qcprQuery = `title contains '${cleanJob}' and title contains 'QCPR' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`;
  const qcprFiles = DriveApp.searchFiles(qcprQuery);
  let qcprId = "", qcprTime = 0;
  while (qcprFiles.hasNext()) {
    const file  = qcprFiles.next();
    const fTime = file.getLastUpdated().getTime();
    if (fTime > qcprTime) { qcprTime = fTime; qcprId = file.getId(); }
  }

  if (mmtId || qcprId) {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const gsidSheet = ss.getSheetByName("GSID Database");

    if (gsidSheet) {
      const existingData  = gsidSheet.getDataRange().getValues();
      const alreadyExists = existingData.some(row => row[0].toString().trim().toUpperCase() === cleanJob);

      if (!alreadyExists) {
        gsidSheet.appendRow([cleanJob, mmtId || "No sheet found", qcprId || "No sheet found", ""]);
      }
    }
    return { mmt: mmtId, qcpr: qcprId };
  }

  return null;
}

function getTrackerSpreadsheet(jobNum) {
  const cleanJob = jobNum.toString().trim().toUpperCase();
  const gsidSS   = getSpreadsheetFromGSID(cleanJob, 1);
  if (gsidSS) return gsidSS;
  const newIds = autoDiscoverAndAddJob(cleanJob);
  if (newIds && newIds.mmt) return SpreadsheetApp.openById(newIds.mmt);
  showAlert(`Error: Could not find Tracker for Job ${cleanJob}. Check for typos.`);
  return null;
}

function getQcprSpreadsheet(jobNum) {
  if (!jobNum) return null;
  const cleanJob = jobNum.toString().trim().toUpperCase();
  const gsidSS   = getSpreadsheetFromGSID(cleanJob, 2);
  if (gsidSS) return gsidSS;

  const newIds = autoDiscoverAndAddJob(cleanJob);
  if (newIds && newIds.qcpr) return SpreadsheetApp.openById(newIds.qcpr);

  return null;
}

function getInventorySpreadsheet(rawClientName, jobSheetName = "") {
  const cleanJob = jobSheetName.toString().trim().toUpperCase();
  if (cleanJob) {
    const gsidSS = getSpreadsheetFromGSID(cleanJob, 3);
    if (gsidSS) return gsidSS;
  }

  let client      = rawClientName.toString().trim();
  const clientUpper = client.toUpperCase();
  if      (clientUpper.includes("AOC"))     client = "AOC";
  else if (clientUpper.includes("CENOVUS")) client = "Cenovus";
  else if (clientUpper.includes("CNRL"))    client = "CNRL";

  let searchTerm = client + "INV";
  if (client === "CNRL") {
    searchTerm = (cleanJob === "241143C") ? "CNRLInv2024-2" : "CNRLINV2025";
  }

  const query = `title contains '${searchTerm}' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`;
  const ss    = _driveFetch("INV_" + searchTerm.replace(/\s+/g, ''), query, true);
  if (!ss) showAlert(`Error: Could not find inventory file for '${searchTerm}'.`);
  return ss;
}
