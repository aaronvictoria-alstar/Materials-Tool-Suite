// KITTING TOOL V1.02

// ==========================================
// 6. KITTING TOOL: SEARCH & PULL DRAWING
// ==========================================
// Pulls specific drawing requirements from the MMT FAB sheet, maps them against historical batches, and queues them into the Kitting Tool grid.
function searchAndPullKittingDrawing() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ktSheet = ss.getSheetByName(KT_SHEET_NAME);

  if (!ktSheet) {
    showAlert(`Error: '${KT_SHEET_NAME}' sheet not found.`);
    return;
  }

  const jobNum     = getInputValue(ktSheet, "KT_input_job");
  const rawDrawNum = getInputValue(ktSheet, "KT_input_drawing");
  const clientName = getInputValue(ktSheet, "KT_input_client");

  if (!jobNum || !rawDrawNum || !clientName) {
    showAlert("Please ensure Job Number, Client Name, and Drawing Number are filled out.");
    return;
  }

  // Clean the job number for our history matching
  const cleanJobNum = jobNum.toString().toUpperCase().trim();
  const searchDrawings = rawDrawNum.split(',').map(d => d.trim().toUpperCase()).filter(Boolean);
  if (searchDrawings.length === 0) return;

  const searchTargets = searchDrawings.map(d => {
    const esc = d.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    return {
      raw: d,
      regex: new RegExp(`(^|[\\/\\s\\-,\\&])${esc}([\\/\\s\\-,\\&]|$)`, 'i')
    };
  });

  updateRawDataTab(jobNum, clientName, "KT Data", true);

  const trackerSS = getTrackerSpreadsheet(jobNum);
  if (!trackerSS) return;

  const fabSheet = trackerSS.getSheetByName("Item Report-FAB");
  if (!fabSheet) {
    showAlert("Error: Could not find the 'Item Report-FAB' sheet in the Tracker.");
    return;
  }

  const coords = getJobCoordinatesFromGSID(jobNum);
  if (!coords || !coords.fabDraw || !coords.fabDesc || !coords.fabBom || !coords.fabItem) {
    showAlert("Error: Missing FAB column coordinates in the GSID Database. Run the Nightly Updater.");
    return;
  }

  const lastRow  = fabSheet.getLastRow();
  const lastCol  = fabSheet.getLastColumn();
  if (lastRow < coords.fabDraw.dataRowStart) {
    showAlert("The 'Item Report-FAB' sheet is empty.");
    return;
  }

  const ktLastCol = ktSheet.getLastColumn();
  const ktHeaders = sanitizeHeaders(ktSheet.getRange(KT_HEADER_ROW, 1, 1, ktLastCol).getValues()[0]);
  const colMap = {
    item:  findCol(ktHeaders, ["Item #", "Item Number", "Item"]) + 1,
    qty:   findCol(ktHeaders, ["Quantity", "Qty", "Req Qty"]) + 1,
    desc:  findCol(ktHeaders, ["Item Description", "Description", "Material Description"]) + 1,
    bom:   findCol(ktHeaders, ["BOMID", "BOM ID"]) + 1,
    draw:  findCol(ktHeaders, ["Drawing #", "Drawing", "Drawing Number", "ISO"]) + 1,
    batch: findCol(ktHeaders, ["Batch ID", "Batch", "Batch #"]) + 1
  };

  const existingSet = new Set();
  const ktLastRow   = ktSheet.getLastRow();
  if (ktLastRow >= KT_START_ROW) {
    const currentKtData = ktSheet.getRange(KT_START_ROW, 1, ktLastRow - KT_START_ROW + 1, ktLastCol).getValues();
    for (const row of currentKtData) {
      const rowItem = colMap.item > 0 ? row[colMap.item - 1].toString().trim() : "";
      const rowDraw = colMap.draw > 0 ? row[colMap.draw - 1].toString().toUpperCase().trim() : "";
      if (rowItem && rowDraw) existingSet.add(`${rowDraw}|||${rowItem}`);
    }
  }

  const histSheet       = ss.getSheetByName("KT Batch History");
  const batchHistoryMap = new Map();

  if (histSheet) {
    const histLastRow = histSheet.getLastRow();
    if (histLastRow >= 2) {
      const histData = histSheet.getRange(2, 1, histLastRow - 1, 12).getValues();
      for (const r of histData) {
        // THE FIX: Pull the Job Number from the history row
        const hJob   = r[0] ? r[0].toString().toUpperCase().trim() : "";
        const hDraw  = r[1] ? r[1].toString().toUpperCase().trim() : "";
        const hItem  = r[2] ? r[2].toString().trim() : "";
        const hBatch = r[10] ? r[10].toString().padStart(4, '0') : "";
        
        // Ensure it ONLY maps rows if they belong to the Job Number you are searching!
        if (hJob === cleanJobNum && hDraw && hItem && hBatch) {
          batchHistoryMap.set(`${hDraw}|||${hItem}`, hBatch);
        }
      }
    }
  }

  const fabData = fabSheet.getRange(
    coords.fabDraw.dataRowStart,
    1,
    lastRow - coords.fabDraw.headerRow,
    lastCol
  ).getValues();
  
  const pulledItems = [];

  for (const row of fabData) {
    const rawDraw  = row[coords.fabDraw.colIdx] ? row[coords.fabDraw.colIdx].toString().toUpperCase().trim() : "";
    const cleanDraw = cleanDrawingNumber(rawDraw);
    let matchedDraw = "";
    for (const target of searchTargets) {
      if (target.regex.test(rawDraw) || target.regex.test(cleanDraw)) {
        matchedDraw = target.raw;
        break;
      }
    }

    if (matchedDraw) {
      const rowDesc = row[coords.fabDesc.colIdx] ? row[coords.fabDesc.colIdx].toString() : "";
      const rowBom  = row[coords.fabBom.colIdx]  ? row[coords.fabBom.colIdx].toString()  : "";
      const rowItem = row[coords.fabItem.colIdx]  ? row[coords.fabItem.colIdx].toString().trim() : "";
      const qty     = coords.fabQty ? row[coords.fabQty.colIdx] : 0;
      const catInfo = getCategoryLogic(rowBom, rowDesc);
      
      if (catInfo.category !== "Pipe") {
        const compositeKey = `${cleanDraw}|||${rowItem}`;
        if (!existingSet.has(compositeKey)) {
          pulledItems.push({
            item:  rowItem,
            qty,
            desc:  rowDesc,
            bom:   rowBom,
            draw:  cleanDraw,
            batch: batchHistoryMap.get(compositeKey) || ""
          });
          existingSet.add(compositeKey);
        }
      }
    }
  }

  if (pulledItems.length === 0) {
    showAlert("No new items (excluding pipe) found for the provided drawings.\n(Note: Duplicates currently in the table are ignored).");
    return;
  }

  const outItem = [], outQty = [], outDesc = [], outBom = [], outDraw = [], outBatch = [];
  for (const p of pulledItems) {
    outItem.push([p.item]);
    outQty.push([p.qty]);
    outDesc.push([p.desc]);
    outBom.push([p.bom]);
    outDraw.push([p.draw]);
    outBatch.push([p.batch]);
  }

  const checkColIdx = colMap.draw > 0 ? colMap.draw : (colMap.item > 0 ? colMap.item : 1);
  const nextRow     = getSafeNextRow(ktSheet, checkColIdx, KT_START_ROW);
  
  if (colMap.item  > 0) ktSheet.getRange(nextRow, colMap.item,  outItem.length,  1).setValues(outItem);
  if (colMap.qty   > 0) ktSheet.getRange(nextRow, colMap.qty,   outQty.length,   1).setValues(outQty);
  if (colMap.desc  > 0) ktSheet.getRange(nextRow, colMap.desc,  outDesc.length,  1).setValues(outDesc);
  if (colMap.bom   > 0) ktSheet.getRange(nextRow, colMap.bom,   outBom.length,   1).setValues(outBom);
  if (colMap.draw  > 0) ktSheet.getRange(nextRow, colMap.draw,  outDraw.length,  1).setValues(outDraw);
  if (colMap.batch > 0) ktSheet.getRange(nextRow, colMap.batch, outBatch.length, 1).setValues(outBatch);
  
  clearNamedRanges(ktSheet, ["KT_input_drawing"]);

  showAlert(`Success! Added ${pulledItems.length} new item(s) across ${searchDrawings.length} drawing(s).`);
}

// ==========================================
// 7. CLEAR KITTING TOOL
// ==========================================
// Clears the input fields and drawing table to reset the Kitting Tool UI.
function clearKittingTool() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ktSheet = ss.getSheetByName(KT_SHEET_NAME);
  if (!ktSheet) {
    showAlert(`Error: '${KT_SHEET_NAME}' sheet not found.`);
    return;
  }

  clearNamedRanges(ktSheet, ["KT_input_job", "KT_input_drawing"]);
  
  const lastRow     = Math.max(ktSheet.getLastRow(), KT_START_ROW);
  const rowsToClear = Math.max(lastRow - KT_START_ROW + 1, 50);
  
  const drawingTable = ktSheet.getRange(KT_START_ROW, 2, rowsToClear, 7);
  drawingTable.clearContent();
  drawingTable.clearDataValidations();

  ss.toast("Drawing Table cleared.", "Success", 3);
}

// ==========================================
// 9. KITTING TOOL: CREATE BATCH
// ==========================================
// Consolidates items in the UI into a unique Batch ID, logs them to history, and generates a rollup pick-list for the shop.
function createKittingBatch() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const ktSheet   = ss.getSheetByName(KT_SHEET_NAME);
  const histSheet = ss.getSheetByName("KT Batch History");
  const dataSheet = ss.getSheetByName("KT Data");
  
  if (!histSheet || !dataSheet) {
    showAlert("Error: 'KT Batch History' or 'KT Data' sheet missing.");
    return;
  }

  const rawJobNum = getInputValue(ktSheet, "KT_input_job");
  if (!rawJobNum) {
    showAlert("Error: Please enter a Job Number before creating a batch.");
    return;
  }

  const jobNum    = rawJobNum.toString().toUpperCase().trim();
  const lastRow   = Math.max(ktSheet.getLastRow(), KT_START_ROW);
  const rawKtData = ktSheet.getRange(KT_START_ROW, 2, lastRow - KT_START_ROW + 1, 7).getValues();

  const histLastRow = Math.max(histSheet.getLastRow(), 1);
  const histData    = histLastRow > 1 ? histSheet.getRange(2, 1, histLastRow - 1, 12).getValues() : [];
  
  const drawingBatchMap  = new Map();
  const historyItemsSet  = new Set();
  let maxBatchNum = 0;
  
  for (const r of histData) {
    const hJob      = r[0] ? r[0].toString().toUpperCase().trim() : "";
    const hDraw     = r[1] ? r[1].toString().toUpperCase().trim() : "";
    const hItem     = r[2] ? r[2].toString().trim() : "";
    const hBatchStr = r[10] ? r[10].toString().padStart(4, '0') : "";

    if (hBatchStr) {
      // 1. GLOBAL CHECK: Find the highest Batch ID across ALL jobs so the next one is always sequentially fresh
      const bNum = parseInt(hBatchStr, 10);
      if (!isNaN(bNum) && bNum > maxBatchNum) maxBatchNum = bNum;
     
      // 2. STRICT FILTER: ONLY pull drawings and items into our mapping memory if they belong to the current job!
      if (hJob === jobNum && hDraw) {
        if (!drawingBatchMap.has(hDraw)) drawingBatchMap.set(hDraw, hBatchStr);
        if (hItem) historyItemsSet.add(`${hDraw}|||${hItem}`);
      }
    }
  }

  const newBatchIdStr = (maxBatchNum + 1).toString().padStart(4, '0');
  
  const historyMap = {};
  const dLastRow   = Math.max(dataSheet.getLastRow(), 1);
  
  if (dLastRow > 1) {
    const dData    = dataSheet.getRange(1, 1, dLastRow, dataSheet.getLastColumn()).getValues();
    const dHeaders = sanitizeHeaders(dData[0]);
    const dBomCol  = findCol(dHeaders, ["BOMID", "BOM ID"]);
    const dHeatCol = findCol(dHeaders, ["Heat #", "Heat Number"]);
    const dLocCol  = findCol(dHeaders, ["Location", "Full Location"]);

    if (dBomCol > -1) {
      for (let i = 1; i < dData.length; i++) {
        const hBom = dData[i][dBomCol].toString().trim().toUpperCase();
        if (!hBom) continue;
        if (!historyMap[hBom]) historyMap[hBom] = { heats: new Set(), locs: new Set() };
        
        const hHeat = dHeatCol > -1 ? dData[i][dHeatCol].toString().trim() : "";
        const hLoc  = dLocCol  > -1 ? dData[i][dLocCol].toString().trim()  : "";

        if (hHeat) hHeat.split(/[\n,;]+/).forEach(h => { if (h.trim()) historyMap[hBom].heats.add(h.trim()); });
        if (hLoc)  hLoc.split(/[\n,;]+/).forEach(l => {
          const cl = l.trim();
          if (cl && !["K+R", "QUARANTINE", "LEGACY"].includes(cl.toUpperCase())) {
            historyMap[hBom].locs.add(cl);
          }
        });
      }
    }
  }

  const histRowsToPush  = [];
  const batchIdUpdates  = [];
  const batchIdCounts   = new Map();
  let validItemsCount   = 0;
  
  for (const row of rawKtData) {
    const draw = row[0] ? row[0].toString().toUpperCase().trim() : "";
    const item = row[1] ? row[1].toString().trim() : "";
    const desc = row[2], bom = row[4], qty = row[5];
    
    if (draw && item && desc) {
      validItemsCount++;

      let assignedBatch = drawingBatchMap.get(draw);
      if (!assignedBatch) {
        assignedBatch = newBatchIdStr;
        drawingBatchMap.set(draw, newBatchIdStr);
      }

      batchIdUpdates.push([assignedBatch]);
      batchIdCounts.set(assignedBatch, (batchIdCounts.get(assignedBatch) || 0) + 1);

      const compositeKey = `${draw}|||${item}`;
      if (!historyItemsSet.has(compositeKey)) {
        const bomKey = bom ? bom.toString().toUpperCase().trim() : "";
        const hist   = historyMap[bomKey];
        let heatStr1 = "", heatStr2 = "", locStr = "";
        
        if (hist) {
          const heatArr = Array.from(hist.heats).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
          const locArr  = Array.from(hist.locs).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
          const leftHeats = [], rightHeats = [];
          
          heatArr.forEach((h, i) => { if (i % 2 === 0) leftHeats.push(h); else rightHeats.push(h); });
          
          if (leftHeats.length  > 0) heatStr1 = leftHeats.join("\n");
          if (rightHeats.length > 0) heatStr2 = rightHeats.join("\n");
          if (locArr.length     > 0) locStr   = locArr.join("\n");
        }

        histRowsToPush.push([jobNum, draw, item, desc, "", bom, parseFloat(qty) || 0, heatStr1, heatStr2, locStr, assignedBatch, ""]);
        historyItemsSet.add(compositeKey);
      }
    } else {
      batchIdUpdates.push([""]);
    }
  }

  if (validItemsCount === 0) {
    showAlert("There are no valid items in the Drawing Table to process.");
    return;
  }

  if (histRowsToPush.length > 0) {
    const targetRange = histSheet.getRange(histLastRow + 1, 1, histRowsToPush.length, 12);
    targetRange.setValues(histRowsToPush);

    if (histLastRow >= 2) {
      const formatSource = histSheet.getRange(histLastRow, 1, 1, 12);
      formatSource.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      formatSource.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
  }

  ktSheet.getRange(KT_START_ROW, 8, batchIdUpdates.length, 1).setValues(batchIdUpdates);

  let majorityBatchId = "";
  let maxCount = 0;

  for (const [bId, count] of batchIdCounts.entries()) {
    if (count > maxCount) {
      maxCount = count;
      majorityBatchId = bId;
    }
  }

  const rollupData = [];
  
  for (const r of histData) {
    // SECURITY FILTER: Only rollup old items if they belong to this job!
    if (r[0] && r[0].toString().toUpperCase().trim() === jobNum && r[10] && r[10].toString().padStart(4, '0') === majorityBatchId) {
      rollupData.push({ draw: r[1], desc: r[3], bom: r[5].toString().toUpperCase().trim(), qty: parseFloat(r[6]) || 0, heat1: r[7], heat2: r[8], loc: r[9], datePosted: r[11] });
    }
  }
  for (const r of histRowsToPush) {
    if (r[10] === majorityBatchId) {
      rollupData.push({ draw: r[1], desc: r[3], bom: r[5].toString().toUpperCase().trim(), qty: parseFloat(r[6]) || 0, heat1: r[7], heat2: r[8], loc: r[9], datePosted: r[11] });
    }
  }

  const finalListCount = renderBatchedTableRollup(ktSheet, majorityBatchId, jobNum, rollupData);

  let msg = "Successfully processed the Drawing Table!\n\n";
  msg += histRowsToPush.length > 0
    ? `Added ${histRowsToPush.length} missing line(s) to the Batch History.\n`
    : "No missing lines found (History is already up to date).\n";
  msg += `Displaying Batch '${majorityBatchId}' with ${finalListCount} pick-list items.`;

  showAlert(msg);
}

// ==========================================
// 10. KITTING TOOL: SEARCH / RECALL BATCH
// ==========================================
// Recalls a previously saved Batch ID from the history log and re-renders the pick-list rollup in the UI.
function searchKittingBatch() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const ktSheet   = ss.getSheetByName(KT_SHEET_NAME);
  const histSheet = ss.getSheetByName("KT Batch History");
  
  if (!histSheet) {
    showAlert("Error: 'KT Batch History' sheet missing.");
    return;
  }

  const searchBatchInput = getInputValue(ktSheet, "KT_input_batch");
  if (!searchBatchInput) {
    showAlert("Please enter a Batch ID (or multiple separated by commas) in the Batched Table to recall.");
    return;
  }

  const targetBatches = searchBatchInput.toString().split(',').map(b => b.trim().padStart(4, '0')).filter(b => b !== "");
  const histLastRow = histSheet.getLastRow();
  
  if (histLastRow < 2) {
    showAlert("The Batch History sheet is empty.");
    return;
  }

  const histData       = histSheet.getRange(2, 1, histLastRow - 1, 12).getValues();
  const rawRollupItems = [];
  const jobsFound      = new Set();
  
  for (const row of histData) {
    const rowBatchId = row[10] ? row[10].toString().padStart(4, '0') : "";
    
    if (targetBatches.includes(rowBatchId)) {
      const rowJob = row[0] ? row[0].toString().toUpperCase().trim() : "";
      if (rowJob) jobsFound.add(rowJob);
      
      rawRollupItems.push({
        draw: row[1],
        desc: row[3],
        bom:  row[5].toString().toUpperCase().trim(),
        qty:  parseFloat(row[6]) || 0,
        heat1: row[7],
        heat2: row[8],
        loc:   row[9],
        datePosted: row[11]
      });
    }
  }

  if (rawRollupItems.length === 0) {
    showAlert(`No items found for Batch ID(s) '${targetBatches.join(", ")}' in the history.`);
    return;
  }

  // OPTION A: Safety Lock Check
  if (jobsFound.size > 1) {
    showAlert("Error: You can only combine batches belonging to the same Job Number.\nFound jobs: " + Array.from(jobsFound).join(", "));
    return;
  }

  const recalledJobNum = Array.from(jobsFound)[0];
  const combinedBatchStr = targetBatches.join(", "); 
  
  const finalListCount = renderBatchedTableRollup(ktSheet, combinedBatchStr, recalledJobNum, rawRollupItems);
  showAlert(`Success! Recalled Batch(es) '${combinedBatchStr}' (Job: ${recalledJobNum}) with ${finalListCount} pick-list items.`);
}

// ==========================================
// 11. PRINT KITTING BATCH
// ==========================================
// Highlights the active pick-list table so the user can easily print it for the shop floor.
function printKittingBatch() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const ktSheet = ss.getSheetByName(KT_SHEET_NAME);
  if (!ktSheet) return;
  
  const jValues = ktSheet.getRange(1, 10, ktSheet.getMaxRows(), 1).getValues();
  let lastRow   = KT_START_ROW;
  
  for (let i = jValues.length - 1; i >= 0; i--) {
    if (jValues[i][0] && jValues[i][0].toString().trim() !== "") {
      lastRow = i + 1;
      break;
    }
  }

  if (lastRow < KT_START_ROW) lastRow = KT_START_ROW;

  const printStartRow = 2;
  ktSheet.setActiveRange(ktSheet.getRange(printStartRow, 9, lastRow - printStartRow + 1, 10));
}

// ==========================================
// 12. KITTING TOOL: POST BATCH
// ==========================================
// Finalizes a batch by stamping completion dates to the Job Sheet History, QCPR (Issued), and MMT (Kitted). Uses surgical RangeLists to prevent ghost edits.
function postKittingBatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ktSheet = ss.getSheetByName(KT_SHEET_NAME);
  const histSheet = ss.getSheetByName("KT Batch History");
  
  if (!ktSheet || !histSheet) {
    showAlert("Error: One or more required sheets are missing.");
    return;
  }

  const batchInput = getInputValue(ktSheet, "KT_input_batch").toString().trim();
  const jobNum = getInputValue(ktSheet, "KT_batch_job").toString().toUpperCase().trim();
  
  if (!batchInput || !jobNum) {
    showAlert("Please ensure a Batch is loaded and the Job Number is present before posting.");
    return;
  }

  const batchIds = batchInput.split(',').map(b => b.trim().padStart(4, '0')).filter(b => b !== "");

  const spoolListRange = ktSheet.getRange("KT_spool_list");
  const spoolData = ktSheet.getRange(spoolListRange.getRow() + 1, spoolListRange.getColumn(), 50, 2).getValues();
 
  const checkedDrawings = [];
  for (let r = 0; r < spoolData.length; r++) {
    if (spoolData[r][0] && spoolData[r][1] === true) {
      checkedDrawings.push(spoolData[r][0].toString().toUpperCase().trim());
    }
  }

  if (checkedDrawings.length === 0) {
    showAlert("No drawings are checked for posting.");
    return;
  }

  const histData = histSheet.getDataRange().getValues();
  const drawingsToPostMap = new Map();
  const alreadyPostedDrawings = new Set();
  const historyRowUpdates = [];
  
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  for (let r = 1; r < histData.length; r++) {
    const rowBatch = histData[r][10] ? histData[r][10].toString().padStart(4, '0') : "";
    const rowDraw = histData[r][1] ? histData[r][1].toString().toUpperCase().trim() : "";
    const rowItem = histData[r][2] ? histData[r][2].toString().trim() : "";
    const rowDate = histData[r][11];

    if (batchIds.includes(rowBatch) && checkedDrawings.includes(rowDraw)) {
      if (rowDate && rowDate.toString().trim() !== "") {
        alreadyPostedDrawings.add(rowDraw);
      } else {
        if (!drawingsToPostMap.has(rowDraw)) {
          drawingsToPostMap.set(rowDraw, new Set());
        }
        drawingsToPostMap.get(rowDraw).add(rowItem);
        historyRowUpdates.push({ rowIdx: r + 1, colIdx: 12, val: today });
      }
    }
  }

  for (const d of alreadyPostedDrawings) {
    drawingsToPostMap.delete(d);
  }

  if (drawingsToPostMap.size === 0) {
    showAlert("No new drawings selected to post.\n(Note: Drawings that were already posted have been safely ignored.)");
    return;
  }

  for (const update of historyRowUpdates) {
    histSheet.getRange(update.rowIdx, update.colIdx).setValue(update.val);
  }

  const coords = getJobCoordinatesFromGSID(jobNum);
  
  if (!coords || !coords.qcprSpool || !coords.qcprIssued || !coords.fabKitted || !coords.fabItem || !coords.fabDraw || !coords.fabBom || !coords.fabDesc) {
     showAlert("Missing required coordinates in GSID. Ensure 'Kitted', 'Issued to Shop', 'Spool', 'BOM ID', and 'Description' columns are mapped.");
     return;
  }

  let updatedQcpr = false;
  let updatedMmt = false;
  
  // --- SURGICAL UPDATE: QCPR ---
  const qcprSS = getQcprSpreadsheet(jobNum);
  if (qcprSS) {
    const fabSheet = qcprSS.getSheetByName("Fab Data");
    if (fabSheet) {
      const fLastRow = fabSheet.getLastRow();
      const startRow = Math.min(coords.qcprSpool.dataRowStart, coords.qcprIssued.dataRowStart);
      
      if (fLastRow >= startRow) {
          const spoolColData = fabSheet.getRange(startRow, coords.qcprSpool.colIdx + 1, fLastRow - startRow + 1, 1).getValues();
          const qcprIssuedLetter = getColLetter(coords.qcprIssued.colIdx + 1);
          const qcprUpdateCells = [];
          
          for (let i = 0; i < spoolColData.length; i++) {
            const rawDraw = spoolColData[i][0] ? spoolColData[i][0].toString().toUpperCase().trim() : "";
            const cleanDraw = cleanDrawingNumber(rawDraw);

            if (drawingsToPostMap.has(rawDraw) || drawingsToPostMap.has(cleanDraw)) {
              qcprUpdateCells.push(qcprIssuedLetter + (startRow + i));
            }
          }

          if (qcprUpdateCells.length > 0) {
            fabSheet.getRangeList(qcprUpdateCells).setValue(today);
            updatedQcpr = true;
          }
      }
    }
  }

  // --- SURGICAL UPDATE: MMT ---
  const mmtSS = getTrackerSpreadsheet(jobNum);
  if (mmtSS) {
    const mmtFabSheet = mmtSS.getSheetByName("Item Report-FAB");
    if (mmtFabSheet) {
      const fLastRow = mmtFabSheet.getLastRow();
      const startRow = Math.min(coords.fabItem.dataRowStart, coords.fabDraw.dataRowStart, coords.fabKitted.dataRowStart);
      
      if (fLastRow >= startRow) {
          const drawColData = mmtFabSheet.getRange(startRow, coords.fabDraw.colIdx + 1, fLastRow - startRow + 1, 1).getValues();
          const itemColData = mmtFabSheet.getRange(startRow, coords.fabItem.colIdx + 1, fLastRow - startRow + 1, 1).getValues();
          const bomColData = mmtFabSheet.getRange(startRow, coords.fabBom.colIdx + 1, fLastRow - startRow + 1, 1).getValues();
          const descColData = mmtFabSheet.getRange(startRow, coords.fabDesc.colIdx + 1, fLastRow - startRow + 1, 1).getValues();
         
          const mmtKittedLetter = getColLetter(coords.fabKitted.colIdx + 1);
          const mmtUpdateCells = [];
          
          for (let i = 0; i < drawColData.length; i++) {
            const rawDraw = drawColData[i][0] ? drawColData[i][0].toString().toUpperCase().trim() : "";
            const cleanDraw = cleanDrawingNumber(rawDraw);
            const itemNo = itemColData[i][0] ? itemColData[i][0].toString().trim() : "";
            const bom = bomColData[i][0] ? bomColData[i][0].toString().trim() : "";
            const desc = descColData[i][0] ? descColData[i][0].toString().trim() : "";

            let matchedDraw = "";
            if (drawingsToPostMap.has(rawDraw)) matchedDraw = rawDraw;
            else if (drawingsToPostMap.has(cleanDraw)) matchedDraw = cleanDraw;
            
            if (matchedDraw) {
              const requiredItems = drawingsToPostMap.get(matchedDraw);
              const isPipe = getCategoryLogic(bom, desc).category === "Pipe";
             
              if (requiredItems.has(itemNo) || isPipe) {
                mmtUpdateCells.push(mmtKittedLetter + (startRow + i));
              }
            }
          }

          if (mmtUpdateCells.length > 0) {
            mmtFabSheet.getRangeList(mmtUpdateCells).setValue(today);
            updatedMmt = true;
          }
      }
    }
  }

  let finalMsg = `Successfully processed ${drawingsToPostMap.size} new drawing(s) for Posting!\n\n`;
  finalMsg += "Updates:\n";
  finalMsg += `✅ KT Batch History\n`;
  finalMsg += updatedQcpr ? `✅ QCPR (Issued to Shop)\n` : `❌ QCPR (Skipped or Failed)\n`;
  finalMsg += updatedMmt ? `✅ MMT (Kitted & Issued)\n` : `❌ MMT (Skipped or Failed)\n`;

  showAlert(finalMsg);
}

// ==========================================
// 15. ADMIN TOOLS: UI TRIGGER FOR HEAL MODAL
// ==========================================
// Serves the HTML modal allowing the user to select whether to scan all history or target a specific job.
function showHealModal() {
  const template = HtmlService.createTemplateFromFile("HealModal");
  const html = template.evaluate().setWidth(400).setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, "🔧 Bidirectional Heal");
}

// ==========================================
// 16. ADMIN TOOLS: BIDIRECTIONAL MMT / HISTORY SYNC
// ==========================================
// Scans the Kitting History to retroactively heal missing dates in the MMT, and vice-versa. 
// Accepts an optional targetJob string to drastically speed up execution by ignoring unrelated history.
function syncKittingHistoryAndMMTs(targetJob = "") {
  // Normalize the input just in case it's triggered directly from a menu click event object
  if (typeof targetJob === "object") targetJob = ""; 
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const histSheet = ss.getSheetByName("KT Batch History");
  if (!histSheet) {
    showAlert("Error: 'KT Batch History' sheet missing.");
    return;
  }

  const isTargeted = typeof targetJob === "string" && targetJob.trim() !== "";
  const jobFilter = isTargeted ? targetJob.toUpperCase().trim() : null;

  ss.toast(isTargeted ? `Scanning Batch History for Job: ${jobFilter}...` : "Scanning entire Batch History... This may take a minute.", "Sync Tool Running", -1);
  
  // 1. Load History Data into Memory
  const histData = histSheet.getDataRange().getValues();
  const historyByJob = new Map();
  
  for (let r = 1; r < histData.length; r++) {
    const job = histData[r][0] ? histData[r][0].toString().toUpperCase().trim() : "";
    const draw = histData[r][1] ? histData[r][1].toString().toUpperCase().trim() : "";
    const item = histData[r][2] ? histData[r][2].toString().trim() : "";
    const postDate = histData[r][11]; // Date Posted (Col L)

    if (!job || !draw) continue;
    
    // SURGICAL FILTER: If the user provided a job number, completely ignore rows belonging to other jobs
    if (isTargeted && job !== jobFilter) continue;

    if (!historyByJob.has(job)) historyByJob.set(job, { drawings: new Map() });

    const jobData = historyByJob.get(job);
    if (!jobData.drawings.has(draw)) {
      jobData.drawings.set(draw, { postDate: postDate || null, items: new Set(), rowIndices: [] });
    }

    const drawData = jobData.drawings.get(draw);
    if (item) drawData.items.add(item);
    drawData.rowIndices.push(r);
    
    if (postDate && !drawData.postDate) drawData.postDate = postDate;
  }

  let mmtUpdatesCount = 0;
  let historyRowUpdatesCount = 0;
  let pipeFixedCount = 0;
  
  // 2. Process Each Job Individually
  for (const [jobNum, jobData] of historyByJob.entries()) {
    ss.toast(`Syncing Job: ${jobNum}...`, "Sync Tool Running", 5);
    const mmtSS = getTrackerSpreadsheet(jobNum);
    const coords = getJobCoordinatesFromGSID(jobNum);
   
    if (!mmtSS || !coords || !coords.fabDraw || !coords.fabKitted || !coords.fabItem || !coords.fabBom || !coords.fabDesc) {
      continue; // Skip if MMT or GSID mapping is invalid
    }

    const mmtFabSheet = mmtSS.getSheetByName("Item Report-FAB");
    if (!mmtFabSheet) continue;

    const startRow = Math.min(coords.fabItem.dataRowStart, coords.fabDraw.dataRowStart, coords.fabKitted.dataRowStart);
    const fLastRow = mmtFabSheet.getLastRow();
    if (fLastRow < startRow) continue;
    
    const drawColData = mmtFabSheet.getRange(startRow, coords.fabDraw.colIdx + 1, fLastRow - startRow + 1, 1).getValues();
    const itemColData = mmtFabSheet.getRange(startRow, coords.fabItem.colIdx + 1, fLastRow - startRow + 1, 1).getValues();
    const kittedColData = mmtFabSheet.getRange(startRow, coords.fabKitted.colIdx + 1, fLastRow - startRow + 1, 1).getValues();
    const bomColData = mmtFabSheet.getRange(startRow, coords.fabBom.colIdx + 1, fLastRow - startRow + 1, 1).getValues();
    const descColData = mmtFabSheet.getRange(startRow, coords.fabDesc.colIdx + 1, fLastRow - startRow + 1, 1).getValues();

    // PASS 1: (MMT -> History) Find dates in MMT for drawings missing dates in History
    const mmtDatesByDrawing = new Map();
    for (let i = 0; i < drawColData.length; i++) {
      const rawDraw = drawColData[i][0] ? drawColData[i][0].toString().toUpperCase().trim() : "";
      const cleanDraw = rawDraw.replace(/^[A-Z]+\s+/i, "");
      const kittedDate = kittedColData[i][0];
      
      if (kittedDate && kittedDate.toString().trim() !== "") {
        if (rawDraw && !mmtDatesByDrawing.has(rawDraw)) mmtDatesByDrawing.set(rawDraw, kittedDate);
        if (cleanDraw && !mmtDatesByDrawing.has(cleanDraw)) mmtDatesByDrawing.set(cleanDraw, kittedDate);
      }
    }

    for (const [draw, drawData] of jobData.drawings.entries()) {
      if (!drawData.postDate && mmtDatesByDrawing.has(draw)) {
        const foundDate = mmtDatesByDrawing.get(draw);
        drawData.postDate = foundDate; // Heal the memory

        // Heal the History Sheet array directly
        for (const rIdx of drawData.rowIndices) {
          histData[rIdx][11] = foundDate;
          historyRowUpdatesCount++;
        }
      }
    }

    // PASS 2: (History -> MMT) SURGICAL SYNC
    const mmtKittedLetter = getColLetter(coords.fabKitted.colIdx + 1);
    const mmtUpdatesByDate = {}; // Maps a specific date to an array of A1 cell coordinates
    let mmtMadeChanges = false;

    for (let i = 0; i < drawColData.length; i++) {
      const rawDraw = drawColData[i][0] ? drawColData[i][0].toString().toUpperCase().trim() : "";
      const cleanDraw = rawDraw.replace(/^[A-Z]+\s+/i, "");
      const itemNo = itemColData[i][0] ? itemColData[i][0].toString().trim() : "";
      const bom = bomColData[i][0] ? bomColData[i][0].toString().trim() : "";
      const desc = descColData[i][0] ? descColData[i][0].toString().trim() : "";
      const currentKitted = kittedColData[i][0];
      
      let matchedDraw = "";
      if (jobData.drawings.has(rawDraw)) matchedDraw = rawDraw;
      else if (jobData.drawings.has(cleanDraw)) matchedDraw = cleanDraw;
      
      if (matchedDraw) {
        const drawData = jobData.drawings.get(matchedDraw);
        if (drawData.postDate) {
          const isPipe = getCategoryLogic(bom, desc).category === "Pipe";
          // If it's Pipe OR it's an explicitly batched item...
          if (drawData.items.has(itemNo) || isPipe) {
            // ...and it doesn't already have a date stamped...
            if (!currentKitted || currentKitted.toString().trim() === "") {
              
              // Queue it for the surgical strike!
              const targetDateStr = drawData.postDate.toString(); 
              if (!mmtUpdatesByDate[targetDateStr]) mmtUpdatesByDate[targetDateStr] = { dateObj: drawData.postDate, cells: [] };
              
              mmtUpdatesByDate[targetDateStr].cells.push(mmtKittedLetter + (startRow + i));
              
              mmtMadeChanges = true;
              if (isPipe) pipeFixedCount++;
            }
          }
        }
      }
    }

    // Push surgical changes back to MMT grouped by date
    if (mmtMadeChanges) {
      for (const targetDate in mmtUpdatesByDate) {
        const updateGroup = mmtUpdatesByDate[targetDate];
        if (updateGroup.cells.length > 0) {
          mmtFabSheet.getRangeList(updateGroup.cells).setValue(updateGroup.dateObj);
        }
      }
      mmtUpdatesCount++;
    }
  }

  // 3. Push changes back to History tab if we found missing dates
  if (historyRowUpdatesCount > 0) {
    histSheet.getRange(1, 1, histData.length, histData[0].length).setValues(histData);
  }

  ss.toast("Sync complete!", "Success", 3);
 
  let finalMsg = `Bidirectional Sync Complete!\n\n`;
  finalMsg += `🔄 MMTs Updated: ${mmtUpdatesCount} jobs repaired.\n`;
  finalMsg += `🔧 Pipe Rows Fixed: ${pipeFixedCount} missing pipe stamps added to MMTs.\n`;
  finalMsg += `📅 History Rows Healed: ${historyRowUpdatesCount} missing dates pulled from MMTs into KT Batch History.\n`;
 
  showAlert(finalMsg);
}

// ==========================================
// 17. KITTING TOOL: POST ALL BATCH
// ==========================================
// Automatically checks all valid drawings in the spool list and triggers the post sequence.
function postAllKittingBatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ktSheet = ss.getSheetByName(KT_SHEET_NAME);
  if (!ktSheet) {
    showAlert("Error: One or more required sheets are missing.");
    return;
  }

  const spoolListRange = ktSheet.getRange("KT_spool_list");
  const spoolDataRange = ktSheet.getRange(spoolListRange.getRow() + 1, spoolListRange.getColumn(), 50, 1);
  const spoolData = spoolDataRange.getValues();
 
  let validCount = 0;

  for (let r = 0; r < spoolData.length; r++) {
    if (spoolData[r][0] && spoolData[r][0].toString().trim() !== "") {
      validCount++;
    }
  }

  if (validCount === 0) {
    showAlert("No drawings found in the list to post.");
    return;
  }

  // Write the true checks back to the sheet dynamically matching the exact row count
  const checksToPush = new Array(validCount).fill([true]);
  ktSheet.getRange(spoolListRange.getRow() + 1, spoolListRange.getColumn() + 1, validCount, 1).setValues(checksToPush);
  
  // Force Google Sheets to complete the visual update of the checkboxes before moving on
  SpreadsheetApp.flush();

  // Trigger the existing post logic
  postKittingBatch();
}