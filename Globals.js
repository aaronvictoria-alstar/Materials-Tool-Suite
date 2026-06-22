// GLOBALS V1.02

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
// GLOBAL HELPER: KITTING BATCH ROLLUP ENGINE
// ==========================================
// Consolidates identical BOM IDs, calculates the most common drawing requirements, and dynamically formats the spool list for the KT UI.
function renderBatchedTableRollup(ktSheet, batchIdStr, jobNum, rawItemsArray) {
  const rollupMap = new Map();
  const globalDrawings = new Set();
  const unpostedDrawings = new Set();
  const drawingPostedMap = new Map();
  for (const item of rawItemsArray) {
    const cleanDraw = item.draw ? item.draw.toString().toUpperCase().trim() : "UNKNOWN";
    globalDrawings.add(cleanDraw);
    const isPosted = (item.datePosted && item.datePosted.toString().trim() !== "");
    
    if (isPosted) {
      drawingPostedMap.set(cleanDraw, true);
    } else {
      unpostedDrawings.add(cleanDraw);
      const rKey = item.desc + "|||" + item.bom;
      if (rollupMap.has(rKey)) {
        rollupMap.get(rKey).qty += item.qty;
        const existing = rollupMap.get(rKey).drawings.get(cleanDraw) || 0;
        rollupMap.get(rKey).drawings.set(cleanDraw, existing + item.qty);
      } else {
        rollupMap.set(rKey, {
          desc: item.desc, bom: item.bom, qty: item.qty,
          heat1: item.heat1, heat2: item.heat2, loc: item.loc,
          drawings: new Map([[cleanDraw, item.qty]])
        });
      }
    }
  }

  const rollupItems = [];
  const unpostedDrawingsArray = Array.from(unpostedDrawings);
  for (const [, v] of rollupMap) {
    const catInfo  = getCategoryLogic(v.bom, v.desc);
    v.cat          = catInfo.category;
    v.subcat       = catInfo.subcat;
    v.size         = catInfo.size;
    v.catWeight    = CAT_SORT_ORDER[v.cat] || 99;

    const qtyMap  = new Map();
    let maxCount  = 0, modeQty = -1, tie = false;
    for (const d of unpostedDrawingsArray) {
      const q = v.drawings.get(d) || 0;
      if (!qtyMap.has(q)) qtyMap.set(q, []);
      qtyMap.get(q).push(d);

      const currentCount = qtyMap.get(q).length;
      if (currentCount > maxCount) {
        maxCount = currentCount;
        modeQty = q; tie = false;
      }
      else if (currentCount === maxCount) { tie = true; }
    }

    if (qtyMap.size === 1) {
      v.drawingsStr = modeQty + " each";
    } else if (tie || maxCount === 1) {
      v.drawingsStr = Array.from(v.drawings.entries())
        .sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true }))
        .map(entry => `${entry[0]} (${entry[1]})`)
        .join("\n");
    } else {
      const outliers = [];
      for (const [q, arr] of qtyMap.entries()) {
        if (q !== modeQty) arr.forEach(d => outliers.push(`${d} (${q})`));
      }
      v.drawingsStr = outliers.sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).join("\n");
    }

    rollupItems.push(v);
  }

  rollupItems.sort((a, b) => {
    const locA   = a.loc ? a.loc.toString().trim() : "ZZZZ";
    const locB   = b.loc ? b.loc.toString().trim() : "ZZZZ";
    const locCmp = locA.localeCompare(locB, undefined, { numeric: true });
    if (locCmp !== 0) return locCmp;
    if (a.catWeight !== b.catWeight) return a.catWeight - b.catWeight;

    if (["Pipe", "Grayloc"].includes(a.cat)) {
      if (a.size !== b.size) return a.size - b.size;
      return a.subcat.localeCompare(b.subcat);
    }
 
    else if (["Flange", "Fittings", "Valve", "Support", "Bolt-Up & Gaskets"].includes(a.cat)) {
      const subCmp = a.subcat.localeCompare(b.subcat);
      if (subCmp !== 0) return subCmp;
      return a.size - b.size;
    } else {
      return a.subcat.localeCompare(b.subcat);
    }
  });
  const batchedOutput = [];
  for (const item of rollupItems) {
    batchedOutput.push([item.desc, "", item.qty, item.heat1, item.heat2, item.loc, item.drawingsStr]);
  }

  try { ktSheet.getRange("KT_input_batch").setValue(batchIdStr); } catch (e) {}
  try { ktSheet.getRange("KT_batch_job").setValue(jobNum); } catch (e) {}

  try {
    const spoolListRange  = ktSheet.getRange("KT_spool_list");
    const sortedDrawings  = Array.from(globalDrawings).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    const numSpools       = sortedDrawings.length;
    if (numSpools < 50) {
      const rowsToClear     = 50 - numSpools;
      const clearDrawRange  = spoolListRange.offset(numSpools + 1, 0, rowsToClear, 1);
      const clearCheckRange = spoolListRange.offset(numSpools + 1, 1, rowsToClear, 1);
      const defaultFmtRange = spoolListRange.offset(numSpools + 1, 2, rowsToClear, 1);
      
      clearDrawRange.clearContent();
      clearCheckRange.clearContent();
      clearCheckRange.clearDataValidations();

      defaultFmtRange.copyTo(clearDrawRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      defaultFmtRange.copyTo(clearCheckRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }

    if (numSpools > 0) {
      const outDrawings = sortedDrawings.map(d => [d]);
      const outChecks = sortedDrawings.map(d => [drawingPostedMap.has(d) ? true : false]);

      const drawTargetRange = spoolListRange.offset(1, 0, numSpools, 1);
      const checkTargetRange = spoolListRange.offset(1, 1, numSpools, 1);
      
      drawTargetRange.setValues(outDrawings);
      
      spoolListRange.offset(1, 0).copyTo(drawTargetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      spoolListRange.offset(1, 0).copyTo(checkTargetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

      checkTargetRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
      checkTargetRange.setValues(outChecks);
    } else {
      spoolListRange.offset(1, 0).clearContent();
      spoolListRange.offset(1, 1).clearContent();
      spoolListRange.offset(1, 1).clearDataValidations();

      const singleDefault = spoolListRange.offset(1, 2);
      singleDefault.copyTo(spoolListRange.offset(1, 0), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      singleDefault.copyTo(spoolListRange.offset(1, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
  } catch (e) {}

  const { totalInches, priorityCounts, sizeCounts } = fetchQcprStats(jobNum, unpostedDrawings);
  let finalPri = "";
  const uniquePris = Object.keys(priorityCounts);

  if (uniquePris.length === 1) {
    finalPri = uniquePris[0];
  } else if (uniquePris.length > 1) {
    const pSet = new Set();
    uniquePris.forEach(priStr => {
      const match = priStr.match(/P(\d+)/i);
      if (match) pSet.add(parseInt(match[1], 10));
    });
    const pArr = Array.from(pSet).sort((a, b) => a - b);
    
    if (pArr.length === 0) {
      finalPri = uniquePris[0];
    } else if (pArr.length === 1) {
      finalPri = "P" + pArr[0];
    } else if (pArr.length === 2) {
      finalPri = "P" + pArr[0] + " & " + pArr[1];
    } else if (pArr.length <= 5) {
      const last = pArr.pop();
      finalPri = "P" + pArr.join(", ") + ", & " + last;
    } else {
      const firstFive = pArr.slice(0, 5);
      finalPri = "P" + firstFive.join(", ") + ", etc.";
    }
  }

  let finalSizeStr = "";
  if (Object.keys(sizeCounts).length > 0) {
    let maxCount = 0;
    for (const s in sizeCounts) {
      if (sizeCounts[s] > maxCount) maxCount = sizeCounts[s];
    }
    const majoritySizes = Object.keys(sizeCounts).filter(s => sizeCounts[s] === maxCount);
    const parseSizeMath = (val) => {
      let total = 0;
      val.split(' ').forEach(p => {
        if (p.includes('/')) {
          const f = p.split('/');
          total += (parseFloat(f[0]) / parseFloat(f[1])) || 0;
        } else {
          total += parseFloat(p) || 0;
        }
      });
      return total;
    };

    majoritySizes.sort((a, b) => parseSizeMath(a) - parseSizeMath(b));

    if (majoritySizes.length === 1)      finalSizeStr = majoritySizes[0] + '" Pipe';
    else if (majoritySizes.length === 2) finalSizeStr = majoritySizes[0] + '" & ' + majoritySizes[1] + '" Pipe';
    else                                 finalSizeStr = majoritySizes[majoritySizes.length - 1] + '" Pipe';
  }

  try { ktSheet.getRange("KT_batch_priority").setValue(finalPri); } catch (e) {}
  try { ktSheet.getRange("KT_batch_size").setValue(finalSizeStr); } catch (e) {}
  try { ktSheet.getRange("KT_batch_inches").setValue(totalInches > 0 ? "Total Inches - " + totalInches.toFixed(2) : ""); } catch (e) {}

  const ktLastRow    = Math.max(ktSheet.getLastRow(), KT_START_ROW);
  const rowsToClear  = Math.max(ktLastRow - KT_START_ROW + 1, 50);
  const batchedRange = ktSheet.getRange(KT_START_ROW, 10, rowsToClear, 7);
  
  batchedRange.clearContent();
  batchedRange.clearDataValidations();
  if (batchedOutput.length > 0) {
    ktSheet.getRange(KT_START_ROW, 10, batchedOutput.length, 7).setValues(batchedOutput);
  }

  return batchedOutput.length;
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

// ==========================================
// GLOBAL HELPER: SAFE RT TABLE CLEARER
// ==========================================
// Wipes the Receiving Tool data grid while safely preserving column headers, formulas, and specific cell background colors.
function clearRtTable(rtSheet) {
  const lastCol = rtSheet.getLastColumn();
  const lastRow = Math.max(rtSheet.getLastRow(), RT_START_ROW);
  const headers = sanitizeHeaders(rtSheet.getRange(RT_HEADER_ROW, 1, 1, lastCol).getValues()[0]);
  const colsToClear = [
    "PO/PL Item #", "Material Description", "BOMID", "Remaining Units",
    "Received Units", "Logged Heat #", "New Heat #", "Dimensions",
    "Current Location", "New Location", "Notes"
  ];
  const rangesToClear = [];
  const rowsToClear = Math.max(lastRow - RT_START_ROW + 1, 50);
 
  let remColRange = null;
  colsToClear.forEach(colName => {
    const colIdx = headers.indexOf(colName);
    if (colIdx > -1) {
      const clearWidth = (colName === "Logged Heat #") ? 2 : 1;
      const targetRange = rtSheet.getRange(RT_START_ROW, colIdx + 1, rowsToClear, clearWidth);
      
      rangesToClear.push(targetRange.getA1Notation());
      
      if (colName === "Remaining Units") {
        remColRange = targetRange;
      }
    }
  });
  
  if (rangesToClear.length > 0) {
    const rangeList = rtSheet.getRangeList(rangesToClear);
    rangeList.clearContent();
    rangeList.clearDataValidations();
  }
 
  if (remColRange) {
    remColRange.setBackground("#f3f3f3");
  }

  // SURGICAL CLEAR: Uncheck the "Update Location" column without destroying the UI checkboxes!
  const upLocIdx = headers.indexOf("Update Location");
  if (upLocIdx > -1) {
     const upLocRange = rtSheet.getRange(RT_START_ROW, upLocIdx + 1, rowsToClear, 1);
     upLocRange.uncheck();
  }
}

// ==========================================
// GLOBAL HELPER: INSTANT TABLE SORTER
// ==========================================
// Sorts the RT grid locally in memory based on PO lines and category weights, then instantly pastes the sorted data back.
function fastSortRtTable(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
 
  const headers = sanitizeHeaders(sheet.getRange(RT_HEADER_ROW, 1, 1, lastCol).getValues()[0]);
  const poColIdx   = headers.indexOf("PO/PL Item #") + 1;
  const descColIdx = headers.indexOf("Material Description") + 1;
  const bomColIdx  = headers.indexOf("BOMID") + 1;
 
  if (poColIdx === 0 || descColIdx === 0 || bomColIdx === 0) return;
  const maxRows = sheet.getMaxRows();
  const rawData = sheet.getRange(RT_START_ROW, 1, maxRows - RT_START_ROW + 1, lastCol).getValues();
  let realLastRow = RT_START_ROW - 1;
  for (let i = 0; i < rawData.length; i++) {
    if (rawData[i][descColIdx - 1] || rawData[i][bomColIdx - 1]) {
      realLastRow = RT_START_ROW + i;
    }
  }
 
  if (realLastRow < RT_START_ROW) return; 
 
  const numRows = realLastRow - RT_START_ROW + 1;
  const dataToSort = sheet.getRange(RT_START_ROW, 1, numRows, lastCol).getValues();
  const orderByPoLine = String(sheet.getRange("RT_input_orderBy").getValue()).toUpperCase().trim() === "TRUE";
  const mapped = dataToSort.map(row => {
    const rawPo = row[poColIdx - 1] ? row[poColIdx - 1].toString() : "";
    const desc  = row[descColIdx - 1] ? row[descColIdx - 1].toString() : "";
    const bom   = row[bomColIdx - 1] ? row[bomColIdx - 1].toString() : "";
    
    const poLines = rawPo.split(/[\n,;]+/).map(l => parseFloat(l.trim())).filter(n => !isNaN(n));
    const primaryPoLine = poLines.length > 0 ? Math.min(...poLines) : 999999;
    
    const catInfo = getCategoryLogic(bom, desc);
    const catWeight = CAT_SORT_ORDER[catInfo.category] || 99;
    
    return { row, primaryPoLine, catWeight, cat: catInfo.category, subcat: catInfo.subcat, size: catInfo.size };
  });
  mapped.sort((a, b) => {
    if (orderByPoLine) {
      if (a.primaryPoLine !== b.primaryPoLine) return a.primaryPoLine - b.primaryPoLine;
    }
    if (a.catWeight !== b.catWeight) return a.catWeight - b.catWeight;
    if (["Pipe", "Grayloc"].includes(a.cat)) {
      if (a.size !== b.size) return a.size - b.size;
      return a.subcat.localeCompare(b.subcat);
    } else if (["Flange", "Fittings", "Valve", "Support", "Bolt-Up & Gaskets"].includes(a.cat)) {
      const subCmp = a.subcat.localeCompare(b.subcat);
      if (subCmp !== 0) return subCmp;
  
      return a.size - b.size;
    } else {
      return a.subcat.localeCompare(b.subcat);
    }
  });
  const sortedData = mapped.map(obj => obj.row);
  sheet.getRange(RT_START_ROW, 1, numRows, lastCol).setValues(sortedData);
}

// ==========================================
// GLOBAL HELPER: PAD PO TO 4 DIGITS
// ==========================================
// Ensures purchase order strings are stripped of accidental leading zeros and cleanly padded to exactly 4 digits.
function padPoNumber(poNum) {
  if (!poNum && poNum !== 0) return "";
  const strippedPo = poNum.toString().trim().replace(/^0+/, '');
  return strippedPo.padStart(4, '0').toUpperCase();
}