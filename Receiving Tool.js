// RECEIVING TOOL V1.02

// ==========================================
// 1. PUSH TO JOB SHEET SCRIPT
// ==========================================
// Scrapes the Receiving Tool UI, formats the line items, pushes them directly to the Client's Job Sheet Inventory, and triggers the Master Log update.
function pushReceivingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rtSheet = ss.getSheetByName(RT_SHEET_NAME);

  const jobSheetName = getInputValue(rtSheet, "RT_input_job");
  const poNumber     = getInputValue(rtSheet, "RT_input_po");
  const packingList  = getInputValue(rtSheet, "RT_input_packing");
  const entryType    = getInputValue(rtSheet, "RT_input_type");
  const clientName   = getInputValue(rtSheet, "RT_input_client");
  const receiverName = getInputValue(rtSheet, "RT_input_receiver");
  const dateLogged   = getInputValue(rtSheet, "RT_input_date");
  
  if (!jobSheetName || !entryType) {
    showAlert("Error: Please ensure Job Number and Entry Type are filled out.");
    return;
  }
  if (!clientName || clientName === "#N/A") {
    showAlert("Error: The Client formula did not return a valid client.");
    return;
  }

  // Safely convert the input date into a true Date Object for formulas
  let finalDate = dateLogged;
  if (dateLogged) {
    const parsed = new Date(dateLogged);
    if (!isNaN(parsed.getTime())) finalDate = parsed;
  }

  // Pad the PO number to exactly 4 digits
  const paddedPo = padPoNumber(poNumber);
  const targetSS = getInventorySpreadsheet(clientName, jobSheetName);
  if (!targetSS) return;
  if (targetSS.getName().includes("Master Material Tracker")) {
    showAlert("FATAL ERROR: Script attempted to push data to the Master Material Tracker. Push aborted to protect the file.");
    return;
  }

  const targetSheet = targetSS.getSheetByName(jobSheetName);
  if (!targetSheet) {
    showAlert(`Error: Tab '${jobSheetName}' not found in ${targetSS.getName()}.`);
    return;
  }

  const sourceLastCol = rtSheet.getLastColumn();
  const sourceLastRow = rtSheet.getLastRow();
  if (sourceLastRow < RT_START_ROW) {
    showAlert("No item data found to push.");
    return;
  }

  const sHeaders   = sanitizeHeaders(rtSheet.getRange(RT_HEADER_ROW, 1, 1, sourceLastCol).getValues()[0]);
  const iDesc      = sHeaders.indexOf("Material Description");
  const iBom       = sHeaders.indexOf("BOMID");
  const iDim       = sHeaders.indexOf("Dimensions");
  const iRec       = sHeaders.indexOf("Received Units");
  const iNewHeat   = sHeaders.indexOf("New Heat #");
  const iNewLoc    = sHeaders.indexOf("New Location");
  const iNotes     = sHeaders.indexOf("Notes"); 
  const iUpdateLoc = sHeaders.indexOf("Update Location");

  const tLastCol = targetSheet.getLastColumn();
  const tHeaders = sanitizeHeaders(targetSheet.getRange(1, 1, 1, tLastCol).getValues()[0]);
  const tCols = {
    type:  findCol(tHeaders, ["Type"]),
    form:  findCol(tHeaders, ["Form"]),
    date:  findCol(tHeaders, ["Date Logged", "Date Entered", "Date Form"]),
    po:    findCol(tHeaders, ["PO", "PO #", "PO Number"]),
    cat:   findCol(tHeaders, ["Category"]),
    sub:   findCol(tHeaders, ["Subcategory"]),
    desc:  findCol(tHeaders, ["Item Description", "Item"]),
    bom:   findCol(tHeaders, ["BOMID"]),
    dim:   findCol(tHeaders, ["Dimensions", "Dimension", "Dim"]),
    qty:   findCol(tHeaders, ["Qty", "Qnty"]),
    tag:   findCol(tHeaders, ["Tag #"]),
    heat:  findCol(tHeaders, ["Heat #", "Heat Number"]),
    loc:   findCol(tHeaders, ["Location", "Full Location"]),
    notes: findCol(tHeaders, ["Notes", "Note"]), 
    recv:  findCol(tHeaders, ["Receiver/Kitter", "Receiver"])
  };
  
  const itemData = rtSheet.getRange(RT_START_ROW, 1, sourceLastRow - RT_HEADER_ROW, sourceLastCol).getValues();
  const rowsToPush = [];
  const locationUpdates = []; // Collects historical location overwrites
  
  for (const row of itemData) {
    const recUnits = iRec > -1 ? row[iRec].toString().trim() : "";
    if (recUnits === "" || recUnits === "0") continue;

    const matDesc = iDesc    > -1 ? row[iDesc].toString().trim()  : "";
    const bomId   = iBom     > -1 ? row[iBom].toString().trim()   : "";
    const dimVal  = iDim     > -1 ? row[iDim].toString().trim()   : "";
    const rawNewH = (iNewHeat > -1 && row[iNewHeat] !== "") ? row[iNewHeat].toString().trim() : "";
    const rawNewL = (iNewLoc  > -1 && row[iNewLoc]  !== "") ? row[iNewLoc].toString().trim()  : "";
    const noteVal = (iNotes   > -1 && row[iNotes] !== undefined) ? row[iNotes].toString().trim() : "";
    const heatNum  = formatMultiLine(rawNewH);
    const location = formatMultiLine(rawNewL);
    const categorization = getCategoryLogic(bomId, matDesc);

    // Track Location Override Checkbox
    const updateLocFlag = (iUpdateLoc > -1 && row[iUpdateLoc] === true);
    if (updateLocFlag && bomId) {
      locationUpdates.push({ bom: bomId, desc: matDesc, loc: location });
    }

    let newRow = new Array(tLastCol).fill("");
    if (tCols.type > -1) newRow[tCols.type] = entryType;
    if (tCols.form > -1) newRow[tCols.form] = packingList;
    if (tCols.date > -1) newRow[tCols.date] = finalDate; // Pushes the true Date object
    if (tCols.po   > -1) newRow[tCols.po]   = paddedPo;
    if (tCols.cat  > -1) newRow[tCols.cat]  = categorization.category;
    if (tCols.sub  > -1) newRow[tCols.sub]  = categorization.subcat;
    if (tCols.desc > -1) newRow[tCols.desc] = matDesc;
    if (tCols.bom  > -1) newRow[tCols.bom]  = bomId;
    if (tCols.dim  > -1) newRow[tCols.dim]  = dimVal;
    if (tCols.qty  > -1) newRow[tCols.qty]  = recUnits;
    if (tCols.heat > -1) newRow[tCols.heat] = heatNum;
    if (tCols.loc  > -1) newRow[tCols.loc]  = location;
    if (tCols.notes > -1) newRow[tCols.notes] = noteVal; 
    if (tCols.recv > -1) newRow[tCols.recv] = receiverName;

    rowsToPush.push(newRow);
  }

  if (rowsToPush.length > 0) {
    const checkColIdx = tCols.desc > -1 ? tCols.desc + 1 : (tCols.bom > -1 ? tCols.bom + 1 : 7);
    const nextRow = getSafeNextRow(targetSheet, checkColIdx, 2);
    const targetRange = targetSheet.getRange(nextRow, 1, rowsToPush.length, tLastCol);
    
    // SURGICAL TEXT FORMATTING: Only format the specific PO column as plain text to protect the leading zeros
    if (tCols.po > -1) {
      targetSheet.getRange(nextRow, tCols.po + 1, rowsToPush.length, 1).setNumberFormat("@");
    }

    targetRange.setValues(rowsToPush);

    if (nextRow > 2) {
      const formatSource = targetSheet.getRange(nextRow - 1, 1, 1, tLastCol);
      formatSource.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      formatSource.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
    
    // BATCH UPDATE: Overwrite Historical Locations in Job Sheet
    if (locationUpdates.length > 0 && tCols.bom > -1 && tCols.desc > -1 && tCols.loc > -1) {
      const histLastRow = targetSheet.getLastRow();
      if (histLastRow > 1) { // 1 is header
        const histData = targetSheet.getRange(1, 1, histLastRow, tLastCol).getValues();
        const locColData = targetSheet.getRange(1, tCols.loc + 1, histLastRow, 1).getValues();
        let madeHistChanges = false;
        
        for (let i = 1; i < histData.length; i++) {
          const rBom  = histData[i][tCols.bom] ? histData[i][tCols.bom].toString().trim().toUpperCase() : "";
          const rDesc = histData[i][tCols.desc] ? histData[i][tCols.desc].toString().trim().toUpperCase() : "";
          
          // Find if this BOM and Description combination is in our updates list
          const match = locationUpdates.find(u => u.bom.toUpperCase() === rBom && u.desc.toUpperCase() === rDesc);
          if (match) {
            locColData[i][0] = match.loc;
            madeHistChanges = true;
          }
        }
        
        if (madeHistChanges) {
          targetSheet.getRange(1, tCols.loc + 1, locColData.length, 1).setValues(locColData);
        }
      }
    }

    handleMasterLogPush(jobSheetName, paddedPo, packingList, receiverName, rowsToPush.length, targetSS.getName());
  } else {
    showAlert("No valid items found to push. Ensure 'Received Units' are filled out.");
  }
}

// ==========================================
// 2. CLEAR TOOL SCRIPT
// ==========================================
// Clears out the user input fields and runs the safe table clearer to reset the UI for the next shipment.
function clearReceivingTool() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rtSheet = ss.getSheetByName(RT_SHEET_NAME);
  
  clearNamedRanges(rtSheet, ["RT_input_job", "RT_input_po", "RT_input_type", "RT_input_packing", "RT_input_receiver", "RT_input_date"]);
  // Uses the new DRY helper to safely wipe the table
  clearRtTable(rtSheet);

  ss.toast("Receiving Tool cleared.", "Success", 3);
}

// ==========================================
// 3. SEARCH AND PULL SCRIPT (VISTA BOM & HISTORY)
// ==========================================
// Merges the expected VISTA BOM lists with the actual Job Sheet history, calculates remaining balances, and injects the output into the RT grid.
function searchAndPullFromVista() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rtSheet = ss.getSheetByName(RT_SHEET_NAME);
  const jobNum     = getInputValue(rtSheet, "RT_input_job");
  const poNum      = getInputValue(rtSheet, "RT_input_po");
  const clientName = getInputValue(rtSheet, "RT_input_client");
  const orderByPoLine = String(getInputValue(rtSheet, "RT_input_orderBy")).toUpperCase().trim() === "TRUE";
  if (!jobNum || !poNum) {
    showAlert("Please enter both a Job Number and a PO Number.");
    return;
  }

  const cleanJob = jobNum.toString().toUpperCase().trim();
  const cleanPo  = padPoNumber(poNum); 
  const searchKey = cleanJob + cleanPo;
  const trackerSS = getTrackerSpreadsheet(jobNum);
  if (!trackerSS) return;

  const vistaSheet = trackerSS.getSheetByName("VISTA BOM");
  if (!vistaSheet) {
    showAlert("Error: Could not find the 'VISTA BOM' sheet in the Tracker.");
    return;
  }

  const coords = getJobCoordinatesFromGSID(jobNum);
  if (!coords || !coords.vistaPo || !coords.vistaDesc || !coords.vistaBom) {
    showAlert("Error: Missing VISTA BOM column coordinates in the GSID Database.");
    return;
  }

  const lastRowVista = vistaSheet.getLastRow();
  const lastColVista = vistaSheet.getLastColumn();
  if (lastRowVista < coords.vistaPo.dataRowStart) {
    showAlert("The VISTA BOM sheet is empty.");
    return;
  }

  const vistaData = vistaSheet.getRange(coords.vistaPo.dataRowStart, 1, lastRowVista - coords.vistaPo.headerRow, lastColVista).getValues();
  const pulledMap = new Map();
  // --- 1. PULL FROM VISTA ---
  for (const row of vistaData) {
    if (row[coords.vistaPo.colIdx] && row[coords.vistaPo.colIdx].toString().trim().toUpperCase() === searchKey) {
      
      const rawOrdered = coords.vistaQtyOrdered ? parseFloat(row[coords.vistaQtyOrdered.colIdx].toString().trim()) || 0 : (coords.vistaQtyDue ? parseFloat(row[coords.vistaQtyDue.colIdx].toString().trim()) || 0 : 0);
      const rawVistaRecv = coords.vistaQtyRecv ? parseFloat(row[coords.vistaQtyRecv.colIdx].toString().trim()) || 0 : 0;
      
      const rawBom    = row[coords.vistaBom.colIdx].toString().trim().toUpperCase();
      const rawDesc   = row[coords.vistaDesc.colIdx].toString().trim();
      const rawPoLine = (coords.vistaPoLine && row[coords.vistaPoLine.colIdx]) ? row[coords.vistaPoLine.colIdx].toString().trim() : "";
      
      const mapKey    = getUnifiedItemKey(rawBom, rawDesc);

      if (pulledMap.has(mapKey)) {
        pulledMap.get(mapKey).ordered += rawOrdered;
        pulledMap.get(mapKey).vistaRecv += rawVistaRecv;
        if (rawPoLine) pulledMap.get(mapKey).poLines.add(rawPoLine);
        
        // Consolidate onto the cleaner, shorter master description if needed
        if (rawDesc && rawDesc.length < pulledMap.get(mapKey).desc.length) {
          pulledMap.get(mapKey).desc = rawDesc;
        }
      } else {
        pulledMap.set(mapKey, { 
          desc: rawDesc, bom: rawBom, 
          ordered: rawOrdered, vistaRecv: rawVistaRecv, 
          poLines: new Set(rawPoLine ? [rawPoLine] : []) 
        });
      }
    }
  }

  if (pulledMap.size === 0) {
    showAlert(`No items found for PO '${searchKey}'.`);
    return;
  }

  const pulledItems = [];
  for (const [, item] of pulledMap) {
    if (item.bom.includes("PIPE")) {
      item.ordered = convertMmToFt(item.ordered);
      item.vistaRecv = convertMmToFt(item.vistaRecv);
    }
    const catInfo = getCategoryLogic(item.bom, item.desc);
    item.cat = catInfo.category; item.subcat = catInfo.subcat;
    item.size = catInfo.size;
    item.catWeight = CAT_SORT_ORDER[catInfo.category] || 99;

    const sortedLines = Array.from(item.poLines).sort((a, b) => (parseFloat(a) || 0) - (parseFloat(b) || 0));
    item.poLineStr = sortedLines.join("\n");
    item.primaryPoLine = sortedLines.length > 0 ? (parseFloat(sortedLines[0]) || 999999) : 999999;
    pulledItems.push(item);
  }

  pulledItems.sort((a, b) => {
    if (orderByPoLine) { if (a.primaryPoLine !== b.primaryPoLine) return a.primaryPoLine - b.primaryPoLine; }
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
  
  // --- 2. PULL FROM LOCAL HISTORY ---
  const historyData = updateRawDataTab(jobNum, clientName, "RT Data", true);
  const historyMap  = {};
  if (historyData && historyData.length > 0) {
    const hHeaders = sanitizeHeaders(historyData[0]);
    const hBomCol  = findCol(hHeaders, ["BOMID"]);
    const hDescCol = findCol(hHeaders, ["Material Description", "Description", "Item Description"]);
    const hHeatCol = findCol(hHeaders, ["Heat #", "Heat Number"]);
    const hLocCol  = findCol(hHeaders, ["Location", "Full Location"]);
    const hQtyCol  = findCol(hHeaders, ["Qty", "Quantity"]); 
    const hPoCol   = findCol(hHeaders, ["PO", "PO #", "PO Number"]);
    if (hBomCol > -1) {
      for (let i = 1; i < historyData.length; i++) {
        const row  = historyData[i];
        let isMatchPO = false;
        if (hPoCol > -1) {
          const rawHistPo = row[hPoCol] ? row[hPoCol].toString().trim() : "";
          const histPoPadded = padPoNumber(rawHistPo);
          if (histPoPadded === cleanPo || rawHistPo.toUpperCase() === searchKey) {
            isMatchPO = true;
          }
        }

        const hBom  = row[hBomCol].toString().trim().toUpperCase();
        const hDesc = hDescCol > -1 ? row[hDescCol].toString().trim() : ""; 
        const hHeat = hHeatCol > -1 ? row[hHeatCol].toString().trim() : "";
        const hLoc  = hLocCol  > -1 ? row[hLocCol].toString().trim()  : "";
        const hQty  = hQtyCol  > -1 ? parseFloat(row[hQtyCol]) || 0 : 0; 

        const hMapKey = getUnifiedItemKey(hBom, hDesc);

        if (!historyMap[hMapKey]) historyMap[hMapKey] = { heats: new Set(), locs: new Set(), localRecvThisPo: 0 };
        if (isMatchPO) {
          historyMap[hMapKey].localRecvThisPo += hQty;
        }

        if (hHeat) hHeat.split(/[\n,;]+/).forEach(h => { if (h.trim()) historyMap[hMapKey].heats.add(h.trim()); });
        if (hLoc) hLoc.split(/[\n,;]+/).forEach(l => {
          const cleanL = l.trim();
          if (cleanL && !["K+R", "QUARANTINE", "LEGACY"].includes(cleanL.toUpperCase())) historyMap[hMapKey].locs.add(cleanL);
        });
      }
    }
  }

  // --- 3. RECONCILIATION MATH, COLORS, & NOTES ---
  const pushArr = { poLine: [], desc: [], bom: [], qty: [], qtyNotes: [], heat1: [], heat2: [], loc: [], updateLoc: [], notes: [], bgColors: [], sysStatus: [] };
  
  for (const item of pulledItems) {
    const searchMapKey = getUnifiedItemKey(item.bom, item.desc);
    const hist = historyMap[searchMapKey];
    const localRecvThisPo = hist ? hist.localRecvThisPo : 0;
    
    let trueRemaining = item.ordered - localRecvThisPo;
    if (item.bom.includes("PIPE")) trueRemaining = trueRemaining.toFixed(2);
    else trueRemaining = Math.max(0, trueRemaining).toString(); 

    pushArr.poLine.push([item.poLineStr]); 
    pushArr.desc.push([item.desc]);
    pushArr.bom.push([item.bom]);
    pushArr.qty.push([trueRemaining]);
    pushArr.updateLoc.push([false]); // Ensure checkbox is populated unchecked
    
    const hasMismatch = Math.abs(item.vistaRecv - localRecvThisPo) > 0.05;
    
    // Calculate Shortage & Tolerance Flag
    let isShort = false;
    let statusFlag = ""; 
    if (item.bom.includes("PIPE")) {
      const hasTolerance = isWithinPipeTolerance(localRecvThisPo, item.ordered);
      isShort = !hasTolerance && (localRecvThisPo < item.ordered);
      if (hasTolerance && localRecvThisPo < item.ordered) {
        statusFlag = "COMPLETE_TOLERANCE";
      }
    } else {
      isShort = localRecvThisPo < item.ordered;
    }
    pushArr.sysStatus.push([statusFlag]); 

    // UX Feature: Determine Colors and Hover Notes
    let qNoteArr = [];
    let bgColor = null; 
    let noteStr = ""; 

    // Step 1: Base Status (Tolerance vs Fully Received)
    if (statusFlag === "COMPLETE_TOLERANCE") {
      bgColor = "#ead1dc"; // Purple/Pink
      qNoteArr.push("✅ Complete (Within 10% Tolerance)");
    } else if (localRecvThisPo >= item.ordered) {
      bgColor = "#d9ead3"; // Green
    }

    // Step 2: Mismatch Override (SCM Error Catching with VISTA Ordered included)
    if (hasMismatch) {
      bgColor = "#FFF2CC"; // Yellow OVERRIDES green/purple
      
      // Cleanly format numbers for the notes based on material type
      const fOrd = item.bom.includes("PIPE") ? item.ordered.toFixed(2) : Math.round(item.ordered).toString();
      const fVis = item.bom.includes("PIPE") ? item.vistaRecv.toFixed(2) : Math.round(item.vistaRecv).toString();
      const fYrd = item.bom.includes("PIPE") ? localRecvThisPo.toFixed(2) : Math.round(localRecvThisPo).toString();
      const mismatchMsg = `⚠️ Discrepancy: VISTA Ordered: ${fOrd} | VISTA Recv: ${fVis} | Yard Recv: ${fYrd}`;
      qNoteArr.push(mismatchMsg);
      // Only physically write the note into the 'Notes' column if we actually need the material
      if (isShort) {
        noteStr = mismatchMsg;
      }
    }

    pushArr.qtyNotes.push([qNoteArr.join("\n\n")]);
    pushArr.notes.push([noteStr]);
    pushArr.bgColors.push([bgColor]);
    
    // Format Heats and Locations
    let heatStr1 = "", heatStr2 = "", locStr = "";
    if (hist) {
      const heatArr = Array.from(hist.heats).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
      const locArr  = Array.from(hist.locs).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
      const leftHeats = [], rightHeats = [];
      heatArr.forEach((h, i) => { if (i % 2 === 0) leftHeats.push(h); else rightHeats.push(h); });
      if (leftHeats.length > 0) heatStr1 = leftHeats.join("\n");
      if (rightHeats.length > 0) heatStr2 = rightHeats.join("\n");
      if (locArr.length > 0) locStr = locArr.join("\n");
    }

    pushArr.heat1.push([heatStr1 || "--"]);
    pushArr.heat2.push([heatStr2 || ""]);
    pushArr.loc.push([locStr || "--"]);
  }

  // --- 4. OUTPUT TO SHEET ---
  const rtHeaders = sanitizeHeaders(rtSheet.getRange(RT_HEADER_ROW, 1, 1, rtSheet.getLastColumn()).getValues()[0]);
  const colMap = {
    poLine:    rtHeaders.indexOf("PO/PL Item #") + 1, 
    desc:      rtHeaders.indexOf("Material Description") + 1,
    bom:       rtHeaders.indexOf("BOMID") + 1,
    rem:       rtHeaders.indexOf("Remaining Units") + 1,
    heat:      rtHeaders.indexOf("Logged Heat #") + 1,
    loc:       rtHeaders.indexOf("Current Location") + 1,
    updateLoc: rtHeaders.indexOf("Update Location") + 1,
    notes:     rtHeaders.indexOf("Notes") + 1,
    sysStatus: rtHeaders.indexOf("System_Status") + 1 
  };
  
  clearRtTable(rtSheet);
  clearNamedRanges(rtSheet, ["RT_input_packing", "RT_input_receiver", "RT_input_date"]);
  
  if (colMap.poLine > 0) rtSheet.getRange(RT_START_ROW, colMap.poLine, pushArr.poLine.length, 1).setValues(pushArr.poLine);
  if (colMap.desc > 0) rtSheet.getRange(RT_START_ROW, colMap.desc, pushArr.desc.length, 1).setValues(pushArr.desc);
  if (colMap.bom  > 0) rtSheet.getRange(RT_START_ROW, colMap.bom,  pushArr.bom.length, 1).setValues(pushArr.bom);
  
  if (colMap.rem  > 0) {
    const remRange = rtSheet.getRange(RT_START_ROW, colMap.rem, pushArr.qty.length, 1);
    remRange.setValues(pushArr.qty);
    remRange.setBackgrounds(pushArr.bgColors);
    remRange.setNotes(pushArr.qtyNotes); // Pushes the hover UI tooltips
  }
  
  if (colMap.heat > 0) {
    rtSheet.getRange(RT_START_ROW, colMap.heat,     pushArr.heat1.length, 1).setValues(pushArr.heat1);
    rtSheet.getRange(RT_START_ROW, colMap.heat + 1, pushArr.heat2.length, 1).setValues(pushArr.heat2);
  }
  
  if (colMap.loc  > 0) rtSheet.getRange(RT_START_ROW, colMap.loc,  pushArr.loc.length, 1).setValues(pushArr.loc);
  
  if (colMap.updateLoc > 0) {
    const upLocRange = rtSheet.getRange(RT_START_ROW, colMap.updateLoc, pushArr.updateLoc.length, 1);
    upLocRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    upLocRange.setValues(pushArr.updateLoc);
  }
  
  if (colMap.notes > 0) rtSheet.getRange(RT_START_ROW, colMap.notes, pushArr.notes.length, 1).setValues(pushArr.notes);
  
  if (colMap.sysStatus > 0) {
    rtSheet.getRange(RT_START_ROW, colMap.sysStatus, pushArr.sysStatus.length, 1).setValues(pushArr.sysStatus);
  }

  const finalLastRow = RT_START_ROW + pulledItems.length - 1;
  const maxRows = rtSheet.getMaxRows();
  const frameSize = 2;
  if (maxRows >= RT_START_ROW) rtSheet.showRows(RT_START_ROW, maxRows - RT_START_ROW + 1);

  const startHideRow = finalLastRow + 1;
  const endHideRow = maxRows - frameSize;
  const hideCount = endHideRow - startHideRow + 1;
  if (hideCount > 0) rtSheet.hideRows(startHideRow, hideCount);
  
  applyMasterFilters(rtSheet);

  const orderMsg = orderByPoLine ? "(Sorted numerically by PO Line #)" : "";
  showAlert(`Success! Pulled & grouped ${pulledItems.length} unique item(s) from VISTA BOM.\n${orderMsg}`);
}

// ==========================================
// 8. MASTER ON EDIT TRIGGER (ROUTES RT & PT)
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

function applyMasterFilters(sheet) {
  const maxRows = sheet.getMaxRows();
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  
  const headers = sanitizeHeaders(sheet.getRange(RT_HEADER_ROW, 1, 1, lastCol).getValues()[0]);
  const iDesc   = headers.indexOf("Material Description") + 1;
  const iBom    = headers.indexOf("BOMID") + 1;
  const remColIdx = findCol(headers, ["Remaining Units", "Remaining", "Rem Qty", "Remaining Quantity", "Qty To Receive", "Balance"]) + 1;
  const statusColIdx = findCol(headers, ["System_Status"]) + 1; // NEW: Find the hidden status column
  
  if (iDesc === 0 || iBom === 0) return;
  // Grab all data including the hidden column
  const data = sheet.getRange(RT_START_ROW, 1, maxRows - RT_START_ROW + 1, lastCol).getValues();
  // 1. Find the REAL last row of the items
  let realLastRow = RT_START_ROW - 1;
  for (let i = 0; i < data.length; i++) {
    const desc = data[i][iDesc - 1] ? data[i][iDesc - 1].toString().trim() : "";
    const bom  = data[i][iBom  - 1] ? data[i][iBom  - 1].toString().trim() : "";
    if (desc !== "" || bom !== "") {
      realLastRow = RT_START_ROW + i;
    }
  }

  if (realLastRow < RT_START_ROW) return;

  let hidePipe = false, hideComplete = false;
  try { hidePipe     = String(sheet.getRange("RT_input_pipe").getValue()).toUpperCase().trim()     === "FALSE"; } catch (e) {}
  try { hideComplete = String(sheet.getRange("RT_input_complete").getValue()).toUpperCase().trim() === "FALSE"; } catch (e) {}

  // 2. Calculate target state for EVERY item row in memory
  const visibilityMap = [];
  for (let i = 0; i < (realLastRow - RT_START_ROW + 1); i++) {
    const desc = data[i][iDesc - 1] ? data[i][iDesc - 1].toString() : "";
    const bom  = data[i][iBom  - 1] ? data[i][iBom  - 1].toString() : "";
    let shouldHide = false;
    if (desc !== "" || bom !== "") {
      if (hidePipe && getCategoryLogic(bom, desc).category === "Pipe") shouldHide = true;
      if (hideComplete && !shouldHide) {
        let isComplete = false;
        // Check 1: Is it mathematically zero or less?
        if (remColIdx > 0) {
          const remVal = parseFloat(data[i][remColIdx - 1]);
          if (!isNaN(remVal) && remVal <= 0) isComplete = true;
        }
        
        // Check 2 (NEW): Did the backend flag it as complete via tolerance?
        if (statusColIdx > 0) {
          const sysStatus = data[i][statusColIdx - 1] ? data[i][statusColIdx - 1].toString() : "";
          if (sysStatus === "COMPLETE_TOLERANCE") {
            isComplete = true;
          }
        }
        
        if (isComplete) shouldHide = true;
      }
    }
    visibilityMap.push({ row: RT_START_ROW + i, hide: shouldHide });
  }

  // 3. Batch apply the visibility state
  if (visibilityMap.length > 0) {
    let currentMode = visibilityMap[0].hide;
    let startRow = visibilityMap[0].row;
    let count = 1;

    for (let i = 1; i < visibilityMap.length; i++) {
      if (visibilityMap[i].hide === currentMode) {
        count++;
      } else {
        if (currentMode) sheet.hideRows(startRow, count);
        else sheet.showRows(startRow, count);
        
        currentMode = visibilityMap[i].hide;
        startRow = visibilityMap[i].row;
        count = 1;
      }
    }
    if (currentMode) sheet.hideRows(startRow, count);
    else sheet.showRows(startRow, count);
  }
}

// ==========================================
// MASTER LOG RECLOG2 PUSH SEQUENCE
// ==========================================
// Scans RecLog2 for matching open shipments to update with end receiving timestamps. If multiple exist, it triggers the HTML modal to let the user select.
function handleMasterLogPush(job, po, packingList, receiver, pushCount, targetSSName) {
  const logSS = SpreadsheetApp.openById(MASTER_LOG_ID);
  const rl2 = logSS.getSheetByName("RecLog2");
  if (!rl2) {
    showAlert(`Success! Pushed ${pushCount} items to ${targetSSName}.\n(Could not update Master Log: 'RecLog2' tab not found)`);
    return;
  }

  const lastRow = rl2.getLastRow();
  const lastCol = rl2.getLastColumn();
  if (lastRow < 2) return;
  // NEW: Absolute coordinate fetching. Guarantees array index 100% matches sheet row!
  const data = rl2.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = sanitizeHeaders(data[0]);

  const cols = {
    job:       findCol(headers, ["Job", "Job #"]),
    po:        findCol(headers, ["PO", "PO #"]),
    pl:        findCol(headers, ["Packing List"]),
    endRec:    findCol(headers, ["End Receiving"]),
    receiver:  findCol(headers, ["Receiver"]),
    subMat:    findCol(headers, ["Submitted by Materials"]),
    freeIssue: findCol(headers, ["Free Issue"]),
    mrr:       findCol(headers, ["MRR Completed"]),
  
    load:      findCol(headers, ["Load Type"]),
    arrival:   findCol(headers, ["Actual Arrival"]),
    recId:     findCol(headers, ["Rec ID"])
  };
  if (cols.job < 0 || cols.po < 0 || cols.pl < 0 || cols.endRec < 0) {
    showAlert(`Success! Pushed ${pushCount} items to ${targetSSName}.\n(Could not update Master Log: Critical headers missing)`);
    return;
  }

  const matches = [];

  // Scan RecLog2 for matching shipments
  for (let i = 1; i < data.length; i++) {
    const rowJob    = data[i][cols.job] ? data[i][cols.job].toString().trim().toUpperCase() : "";
    const rowPo     = data[i][cols.po] ? padPoNumber(data[i][cols.po]) : "";
    const rowPl     = data[i][cols.pl] ? data[i][cols.pl].toString().trim().toUpperCase() : "";
    const endRecVal = data[i][cols.endRec] ? data[i][cols.endRec].toString().trim() : "";
    // Target: Exact Match AND End Receiving is Blank
    if (rowJob === job.toUpperCase() && rowPo === po && rowPl === packingList.toUpperCase() && endRecVal === "") {
      
      let arrVal = data[i][cols.arrival];
      if (arrVal instanceof Date) {
        arrVal = Utilities.formatDate(arrVal, Session.getScriptTimeZone(), "MMM/dd/yyyy HH:mm");
      }
      
      matches.push({
        rowIndex: i + 1, // Absolutely maps to the correct Google Sheet row
        row: data[i],
        loadType: data[i][cols.load] || "--",
        arrival: arrVal || "--",
        recId: cols.recId > -1 ? data[i][cols.recId] : "Unknown"
      });
    }
  }

  const pushMsg = `Success! Pushed ${pushCount} items to ${targetSSName}.`;
  if (matches.length === 0) {
    showAlert(`${pushMsg}\n\n(Note: No open shipment was found in RecLog2 for this Job/PO/Packing List combination.)`);
  } else if (matches.length === 1) {
    // Exactly 1 match! Update it silently in the background
    updateMasterLogRow(rl2, matches[0].rowIndex, cols, matches[0].row, receiver);
    showAlert(`${pushMsg}\nMaster Log (RecLog2) was also successfully updated!`);
  } else {
    // Multiple matches! Trigger the Custom HTML Popup
    const template = HtmlService.createTemplateFromFile("MasterLogModal");
    template.matches = matches;
    template.receiver = receiver;
    const html = template.evaluate().setWidth(650).setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, "⚠️ Multiple Open Shipments Found");
  }
}

// Helper to write the Date objects and formatting to the Master Log
function updateMasterLogRow(sheet, rowIndex, cols, rowData, receiver) {
  const now = new Date();
  const format = "mmm/dd/yyyy hh:mm"; // Yields: May/22/2026 15:43
  
  if (cols.receiver > -1) {
    sheet.getRange(rowIndex, cols.receiver + 1).setValue(receiver);
  }
  
  if (cols.endRec > -1) {
    sheet.getRange(rowIndex, cols.endRec + 1).setValue(now).setNumberFormat(format);
  }
  
  if (cols.subMat > -1) {
    sheet.getRange(rowIndex, cols.subMat + 1).setValue(now).setNumberFormat(format);
  }
  
  if (cols.freeIssue > -1 && cols.mrr > -1) {
    const freeIssueVal = rowData[cols.freeIssue] ? rowData[cols.freeIssue].toString().trim().toUpperCase() : "";
    if (freeIssueVal === "Y" || freeIssueVal === "YES") {
      sheet.getRange(rowIndex, cols.mrr + 1).setValue(now).setNumberFormat(format);
    }
  }
}

// Function called by the HTML Popup once the user makes a selection
function processMasterLogModal(rowIndex, receiver) {
  const logSS = SpreadsheetApp.openById(MASTER_LOG_ID);
  const rl2 = logSS.getSheetByName("RecLog2");
  const lastCol = rl2.getLastColumn();
  
  // Re-verify columns locally to prevent JSON data dropping
  const headers = sanitizeHeaders(rl2.getRange(1, 1, 1, lastCol).getValues()[0]);
  const cols = {
    receiver:  findCol(headers, ["Receiver"]),
    endRec:    findCol(headers, ["End Receiving"]),
    subMat:    findCol(headers, ["Submitted by Materials"]),
    freeIssue: findCol(headers, ["Free Issue"]),
    mrr:       findCol(headers, ["MRR Completed"])
  };
  const rowData = rl2.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  updateMasterLogRow(rl2, rowIndex, cols, rowData, receiver);
}