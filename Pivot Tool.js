// PIVOT TOOL V1.02

// ==========================================
// PIVOT TOOL: INSTANT ROW FILTERING (VISIBILITY MAP)
// ==========================================
// Utilizes a Visibility Map to instantly batch-hide perfectly balanced rows (0-delta) without needing to reload the data.
function applyBalancedFilter(sheet, isShowBalanced) {
  const START_ROW = 7;
  const lastRow = sheet.getLastRow();
 
  if (lastRow < START_ROW) return;




  // Step 1: Grab ONLY the Delta (Shop vs MMT) values from Column K (Index 11)
  const deltas = sheet.getRange(START_ROW, 11, lastRow - START_ROW + 1, 1).getValues();




  // Step 2: Calculate target state for EVERY row in memory
  const visibilityMap = [];
 
  for (let i = 0; i < deltas.length; i++) {
    const val = parseFloat(deltas[i][0]);
    let shouldHide = false;
   
    // If we are NOT showing balanced rows, hide the ones with ~0 delta
    if (!isShowBalanced && !isNaN(val) && Math.abs(val) <= 0.05) {
      shouldHide = true;
    }
   
    visibilityMap.push({ row: START_ROW + i, hide: shouldHide });
  }




  // Step 3: Batch apply the visibility state
  if (visibilityMap.length > 0) {
    let currentMode = visibilityMap[0].hide;
    let startRow = visibilityMap[0].row;
    let count = 1;




    for (let i = 1; i < visibilityMap.length; i++) {
      if (visibilityMap[i].hide === currentMode) {
        count++;
      } else {
        // State changed, execute the batch!
        if (currentMode) sheet.hideRows(startRow, count);
        else sheet.showRows(startRow, count);
       
        currentMode = visibilityMap[i].hide;
        startRow = visibilityMap[i].row;
        count = 1;
      }
    }
    // Catch the final batch at the end of the loop
    if (currentMode) sheet.hideRows(startRow, count);
    else sheet.showRows(startRow, count);
  }
}

// ==========================================
// 4. WMS DASHBOARD ENGINE (PIVOT TOOL / WEB APP API)
// ==========================================
// Executes the core two-way join between MMT data and Job Sheet PT Data. Highlights discrepancies and applies the 10% Pipe Tolerance rule.
function loadPivotData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pivotSheet = ss.getSheetByName("Pivot Tool");




  const jobNum       = getInputValue(pivotSheet, "PT_input_job");
  const clientName   = getInputValue(pivotSheet, "PT_input_client");
  const rawShowBal   = getInputValue(pivotSheet, "PT_input_balanced");
 
  const isShowBalanced = (String(rawShowBal).toUpperCase() === "TRUE" || rawShowBal === true);
 
  if (!jobNum || !clientName) {
    showAlert("Error: Please enter a Job Number and Client.");
    return;
  }




  const cleanJobNum = String(jobNum).trim().toUpperCase();
  const jobBase     = cleanJobNum.replace(/[^0-9]/g, "");
  const jobLetter   = cleanJobNum.replace(/[0-9]/g, "");




  // 1. PULL MMT REQUIREMENTS (The Benchmark)
  const trackerSS = getTrackerSpreadsheet(jobNum);
  const coords = getJobCoordinatesFromGSID(cleanJobNum);




  if (!trackerSS || !coords) {
    showAlert(`Could not locate the MMT/Tracker or GSID coordinates for Job ${jobNum}.`);
    return;
  }




  // --- STEP 1A: GET THE SUMMARY TOTALS (Master Material Data) ---
  const mmtSheet = trackerSS.getSheetByName("Master Material Data");
  if (!mmtSheet) {
    showAlert("Could not find the 'Master Material Data' tab in the MMT.");
    return;
  }




  const mmtData = mmtSheet.getDataRange().getValues();
  let headerRowIdx = -1, cMmtBom = -1, cMmtReq = -1, cMmtRecv = -1, cMmtDesc = -1;




  for (let i = 0; i < Math.min(20, mmtData.length); i++) {
    const tempHeaders = sanitizeHeaders(mmtData[i]);
    const bomIdx = tempHeaders.indexOf("BOM ID");
    const reqIdx = findCol(tempHeaders, ["Total Required on all Drawings", "Total Required"]);
    const recvIdx = findCol(tempHeaders, ["Vista Received QTY", "Vista Received"]);
    const descIdx = findCol(tempHeaders, ["Description", "Item Description", "Material Description"]);




    if (bomIdx > -1 && reqIdx > -1) {
      headerRowIdx = i; cMmtBom = bomIdx; cMmtReq = reqIdx; cMmtRecv = recvIdx; cMmtDesc = descIdx;
      break;
    }
  }




  if (headerRowIdx === -1) {
    showAlert("Error: Could not find 'BOM ID' and 'Total Required' in the top 20 rows of the MMT.");
    return;
  }




  const mmtInfo = {};
 
  for (let i = headerRowIdx + 1; i < mmtData.length; i++) {
    const bomId = String(mmtData[i][cMmtBom]).trim().toUpperCase();
    const reqQty = parseFloat(mmtData[i][cMmtReq]) || 0;
    const recvQty = cMmtRecv > -1 ? (parseFloat(mmtData[i][cMmtRecv]) || 0) : 0;
    const mmtDesc = cMmtDesc > -1 ? String(mmtData[i][cMmtDesc]).trim() : "";
   
    if (bomId) {
      if (!mmtInfo[bomId]) {
        mmtInfo[bomId] = { req: reqQty, recv: recvQty, pos: {}, desc: mmtDesc };
      } else {
        mmtInfo[bomId].req = Math.max(mmtInfo[bomId].req, reqQty);
        mmtInfo[bomId].recv = Math.max(mmtInfo[bomId].recv, recvQty);
        if (mmtDesc && !mmtInfo[bomId].desc) {
          mmtInfo[bomId].desc = mmtDesc;
        }
      }
    }
  }




  // --- STEP 1B: GET EXACT PO BREAKDOWN (VISTA BOM TAB) STRICTLY VIA GSID ---
  const vistaSheet = trackerSS.getSheetByName("VISTA BOM") || trackerSS.getSheetByName("Vista BOM");
 
  if (vistaSheet) {
    if (coords.vistaBom && coords.vistaPo && coords.vistaQtyRecv) {
      const vData = vistaSheet.getDataRange().getValues();
     
      const startRow = Math.min(coords.vistaBom.dataRowStart, coords.vistaPo.dataRowStart, coords.vistaQtyRecv.dataRowStart) - 1;
      const cvBom = coords.vistaBom.colIdx;
      const cvPo = coords.vistaPo.colIdx;
      const cvRecv = coords.vistaQtyRecv.colIdx;
      const cvDesc = coords.vistaDesc ? coords.vistaDesc.colIdx : -1;




      for (let i = startRow; i < vData.length; i++) {
        const bomId = String(vData[i][cvBom]).trim().toUpperCase();
        const recvQty = parseFloat(vData[i][cvRecv]) || 0;
        const rawPo = String(vData[i][cvPo]).trim().toUpperCase();
        const rawDesc = cvDesc > -1 ? String(vData[i][cvDesc]).trim() : "";




        if (bomId && mmtInfo[bomId]) {
          if (rawDesc && !mmtInfo[bomId].desc) {
            mmtInfo[bomId].desc = rawDesc;
          }
          if (recvQty !== 0) {
            let mmtPo = rawPo.replace(/[^0-9A-Z]/g, "");
            if (jobBase && mmtPo.startsWith(jobBase)) {
              const remaining = mmtPo.substring(jobBase.length);
              if (jobLetter && remaining.startsWith(jobLetter)) mmtPo = remaining.substring(jobLetter.length);
              else mmtPo = remaining;
            }
            if (/^\d+$/.test(mmtPo)) mmtPo = mmtPo.padStart(4, '0');




            if (!mmtInfo[bomId].pos[mmtPo]) mmtInfo[bomId].pos[mmtPo] = 0;
            mmtInfo[bomId].pos[mmtPo] += recvQty;
          }
        }
      }
    } else {
      showAlert("⚠️ WARNING: GSID Database is missing coordinates for the VISTA BOM tab on this job. Please run the GSID Updater.");
    }
  } else {
    showAlert("⚠️ WARNING: Could not find the 'VISTA BOM' tab in the MMT. PO Deltas will be missing.");
  }




  // 2. PULL & AGGREGATE JOB SHEET INVENTORY
  const historyData = updateRawDataTab(jobNum, clientName, "PT Data", true);
  if (!historyData || historyData.length <= 1) {
    showAlert("No inventory found on the Job Sheet for this project.");
    return;
  }




  const hHeaders = sanitizeHeaders(historyData[0]);
  const hTypeCol = findCol(hHeaders, ["Type"]);
  const hCatCol  = findCol(hHeaders, ["Category"]);
  const hSubCol  = findCol(hHeaders, ["Subcategory"]);
  const hDescCol = findCol(hHeaders, ["Item Description", "Material Description", "Description"]);
  const hBomCol  = findCol(hHeaders, ["BOMID"]);
  const hQtyCol  = findCol(hHeaders, ["Qty", "Quantity"]);
  const hHeatCol = findCol(hHeaders, ["Heat #", "Heat Number"]);
  const hLocCol  = findCol(hHeaders, ["Location", "Full Location"]);
  const hTagCol  = findCol(hHeaders, ["Tag #", "Tag Number", "Tag"]);
  const hPoCol   = findCol(hHeaders, ["PO", "PO #", "PO Number"]);
  const hFormCol = findCol(hHeaders, ["Form", "Packing List", "PL", "PL #"]);




  if (hTypeCol === -1 || hBomCol === -1 || hQtyCol === -1) {
    showAlert("Error: PT Data is missing required Job Sheet headers (Type, BOMID, Qty).");
    return;
  }




  const inventoryMap = new Map();




  for (let i = 1; i < historyData.length; i++) {
    const row = historyData[i];
    const type = String(row[hTypeCol]).trim().toUpperCase();
    const bomId = hBomCol > -1 ? String(row[hBomCol]).trim().toUpperCase() : "";
    const desc = hDescCol > -1 ? String(row[hDescCol]).trim() : "";
    const qty = parseFloat(row[hQtyCol]) || 0;
    const heat = hHeatCol > -1 ? String(row[hHeatCol]).trim() : "";
    const loc = hLocCol > -1 ? String(row[hLocCol]).trim() : "";
    const tag = hTagCol > -1 ? String(row[hTagCol]).trim() : "";
    const formVal = hFormCol > -1 ? String(row[hFormCol]).trim() : "";
   
    const rawPo = hPoCol > -1 ? String(row[hPoCol]).trim().toUpperCase() : "";
    let invPo = rawPo;
    if (/^\d+$/.test(invPo)) invPo = invPo.padStart(4, '0');
   
    if (qty === 0) continue;
    if (!bomId && !desc) continue; // Skip only if both are blank

    const mapKey = getUnifiedItemKey(bomId, desc);




    if (!inventoryMap.has(mapKey)) {
      inventoryMap.set(mapKey, {
        cat: hCatCol > -1 ? String(row[hCatCol]).trim() : "",
        sub: hSubCol > -1 ? String(row[hSubCol]).trim() : "",
        desc: desc,
        bom: bomId,
        netQty: 0,
        heats: new Set(),
        locs: new Set(),
        tags: new Set(),
        poDetails: {},
        loggedDescriptions: new Set([desc]), // Keep track of logged descriptions to detect drafting mismatches
        descCounts: desc ? { [desc]: 1 } : {}
      });
    }




    const item = inventoryMap.get(mapKey);
    item.loggedDescriptions.add(desc);
    if (desc) item.descCounts[desc] = (item.descCounts[desc] || 0) + 1;




    if (type === "REC" || type === "RETURN") {
      item.netQty += qty;
     
      if (invPo) {
        if (!item.poDetails[invPo]) item.poDetails[invPo] = { qty: 0, pls: new Set() };
        item.poDetails[invPo].qty += qty;
        if (formVal) item.poDetails[invPo].pls.add(formVal);
      }
     
    } else if (type === "KITTED" || type === "KIT" || type === "ISSUED" || type === "ISSUE") {
      item.netQty -= qty;
    }




    if (item.netQty > 0) {
      if (heat) heat.split(/[\n,;]+/).forEach(h => { if (h.trim()) item.heats.add(h.trim()); });
      if (tag) tag.split(/[\n,;]+/).forEach(t => { if (t.trim()) item.tags.add(t.trim()); });
      if (loc) loc.split(/[\n,;]+/).forEach(l => {
        const cleanL = l.trim();
        if (cleanL && !["K+R", "QUARANTINE", "LEGACY"].includes(cleanL.toUpperCase())) item.locs.add(cleanL);
      });
    }
  }




  // 3. MERGE & CALCULATE DELTAS
  for (const [, item] of inventoryMap) {
    if (item.descCounts && Object.keys(item.descCounts).length > 0) {
      let bestDesc = item.desc, maxCount = 0;
      for (const d in item.descCounts) {
        if (item.descCounts[d] > maxCount) { maxCount = item.descCounts[d]; bestDesc = d; }
      }
      item.desc = bestDesc;
    }
  }

  const outputArray = [];
 
  for (const [, item] of inventoryMap) {
    const mInfo = mmtInfo[item.bom] || { req: 0, recv: 0, pos: {}, desc: "" };
   
    let shopQty = item.netQty;
    let mmtRecv = mInfo.recv;
    let mmtReq  = mInfo.req;

    // Prefer the yard description; only use MMT desc when it normalizes identically (better formatting, not a different item)
    let displayDesc = item.desc;
    if (mInfo.desc && normalizeDescription(mInfo.desc) === normalizeDescription(item.desc)) {
      displayDesc = mInfo.desc;
    }

    // Run category discrepancy checks under the same BOM ID to find drafting mistakes
    let isDraftingConflict = false;
    let conflictWarning = "";
    if (item.bom && mInfo.desc) {
      const mmtCategory = getCategoryLogic(item.bom, mInfo.desc);
      for (const loggedD of item.loggedDescriptions) {
        const yardCategory = getCategoryLogic(item.bom, loggedD);
        if (yardCategory.category !== mmtCategory.category || Math.abs(yardCategory.size - mmtCategory.size) > 0.05) {
          isDraftingConflict = true;
          conflictWarning = `⚠️ drafting error: BOM ID defined as ${mmtCategory.category} (${mmtCategory.size}") in MMT, but logged as ${yardCategory.category} (${yardCategory.size}") in Yard!`;
          break;
        }
      }
    }




    const isPipe = item.bom ? item.bom.includes("PIPE") : false;
   
    if (isPipe) {
      mmtRecv = convertMmToFt(mmtRecv);
      mmtReq = convertMmToFt(mmtReq);
     
      // 10% PIPE TOLERANCE LOGIC (TOTALS)
      if (isWithinPipeTolerance(shopQty, mmtRecv)) {
        shopQty = mmtRecv;
      }
    }




    const deltaShopMmt = shopQty - mmtRecv;
    const deltaMmtReq  = mmtRecv - mmtReq;
   
    const fmt = (num) => isPipe ? num.toFixed(2) : Math.round(num).toString();
    const fmtDelta = (num) => (num > 0 ? "+" : "") + fmt(num);




    const locStr = Array.from(item.locs).sort().join("\n") || "NO LOCATION";
    const tagStr = Array.from(item.tags).sort().join("\n") || "--";




    // --- HEAT LOGIC & STRING BUILDER ---
    const allHeats = Array.from(item.heats).sort();
    const heatCol1 = [], heatCol2 = [];
    allHeats.forEach((h, idx) => {
      if (idx % 2 === 0) heatCol1.push(h);
      else heatCol2.push(h);
    });




    // --- PO LOGIC & STRING BUILDER (WITH NOTES) ---
    const allUniquePos = new Set([...Object.keys(mInfo.pos || {}), ...Object.keys(item.poDetails || {})]);
    const poStrings = [];




    for (const po of allUniquePos) {
      if (po === "NO_PO") continue;




      let mQ = (mInfo.pos[po] || 0);
      if (isPipe) mQ = convertMmToFt(mQ);
     
      let iQ = (item.poDetails && item.poDetails[po] ? item.poDetails[po].qty : 0);
     
      if (isPipe && isWithinPipeTolerance(iQ, mQ)) {
        iQ = mQ;
      }
     
      const delta = iQ - mQ;
     
      if (Math.abs(delta) > 0.05) {
        const plArray = item.poDetails && item.poDetails[po] ? Array.from(item.poDetails[po].pls) : [];
        const plStr = plArray.length > 0 ? plArray.join(", ") : "None Logged";
        let hoverNote = `--- PO: ${po} ---\nBOM Required (Total): ${fmt(mmtReq)}\nMMT Logged (This PO): ${fmt(mQ)}\nShop Received (This PO): ${fmt(iQ)}\nPacking List(s): ${plStr}`;
        if (isDraftingConflict && conflictWarning) {
          hoverNote = `🚨 DRAFTING DISCREPANCY DETECTED 🚨\n${conflictWarning}\n\n` + hoverNote;
        }
        poStrings.push({ text: `${po} (${fmt(mQ)}) ⚠️ (${fmtDelta(delta)})`, note: hoverNote });
      }
    }




    poStrings.sort((a, b) => a.text.localeCompare(b.text));
    const poCol = [], noteCol = [];
   
    poStrings.forEach((obj) => {
      poCol.push(obj.text);
      if (obj.note) noteCol.push(obj.note);
    });




    // If there is a drafting conflict, prepend a warning icon to the displayed description
    const finalDisplayDesc = isDraftingConflict ? "⚠️ " + displayDesc : displayDesc;

    // Always push to array (Filter is handled via row hiding!)
    outputArray.push([
      item.cat || "Misc",        // 0
      item.sub || "Misc",        // 1
      item.bom || "NO_BOM",      // 2
      finalDisplayDesc,          // 3
      tagStr,                    // 4
      heatCol1.join("\n") || "--", // 5
      heatCol2.join("\n") || "",   // 6
      locStr,                    // 7
      fmt(shopQty),              // 8
      fmt(mmtRecv),              // 9
      fmtDelta(deltaShopMmt),    // 10
      fmt(mmtReq),               // 11
      fmtDelta(deltaMmtReq),     // 12
      poCol.join("\n") || "",    // 13
      noteCol.join("\n\n")       // 14
    ]);
  }




  outputArray.sort((a, b) => {
    const weightA = CAT_SORT_ORDER[a[0]] || 99;
    const weightB = CAT_SORT_ORDER[b[0]] || 99;
    if (weightA !== weightB) return weightA - weightB;
    const subCmp = a[1].localeCompare(b[1]);
    if (subCmp !== 0) return subCmp;
    return a[2].localeCompare(b[2]);
  });




  const finalValues = [];
  const finalNotes = [];
 
  for (let i = 0; i < outputArray.length; i++) {
    const row = outputArray[i];
    finalValues.push(row.slice(0, 14));
    const noteRow = new Array(14).fill("");
    noteRow[13] = row[14] || "";
    finalNotes.push(noteRow);
  }




  // 4. PRINT TO PIVOT TOOL UI
  const OUTPUT_START_ROW = 6;
  const NUM_COLS = 14;
  const CLEAR_COLS = 15;
 
  const maxRows = pivotSheet.getMaxRows();
  if (maxRows >= OUTPUT_START_ROW) {
    pivotSheet.getRange(OUTPUT_START_ROW, 1, maxRows - OUTPUT_START_ROW + 1, CLEAR_COLS)
              .breakApart()
              .clearContent()
              .clearNote()
              .setBackground(null)
              .setFontWeight("normal");
  }




  if (finalValues.length > 0) {
    const headers = [["Category", "Subcategory", "BOMID", "Description", "Tag #", "Heats", "", "Locations", "Shop Qty Recv", "MMT Qty Recv", "Delta\n(Shop vs MMT)", "MMT Qty Req", "Delta\n(MMT vs Req)", "POs w/ Delta\n(Shop vs MMT)"]];
    const headerRange = pivotSheet.getRange(OUTPUT_START_ROW, 1, 1, NUM_COLS);
   
    headerRange.setValues(headers).setFontWeight("bold").setBackground("#f3f3f3");
    pivotSheet.getRange(OUTPUT_START_ROW, 6, 1, 2).mergeAcross();
   
    const dataRange = pivotSheet.getRange(OUTPUT_START_ROW + 1, 1, finalValues.length, NUM_COLS);
    dataRange.setValues(finalValues);
    dataRange.setNotes(finalNotes);
   
    const bgColors = [];
    for (let i = 0; i < finalValues.length; i++) {
      const dShopMmt = parseFloat(finalValues[i][10]);
      const dMmtReq  = parseFloat(finalValues[i][12]);
      const rowColors = new Array(NUM_COLS).fill(null);
     
      if (Math.abs(dShopMmt) > 0.05) rowColors[10] = "#FFF2CC";
      else rowColors[10] = "#d9ead3";




      if (dMmtReq < -0.05) rowColors[12] = "#f4cccc";
      else rowColors[12] = "#d9ead3";




      bgColors.push(rowColors);
    }
    dataRange.setBackgrounds(bgColors);
    pivotSheet.autoResizeColumns(1, NUM_COLS);
   
    // --- IMMEDIATELY APPLY THE FILTER BASED ON THE CHECKBOX ---
    applyBalancedFilter(pivotSheet, isShowBalanced);
   
    showAlert(`Successfully loaded Dashboard. Balanced rows are ${isShowBalanced ? 'SHOWN' : 'HIDDEN'}.`);
  } else {
    showAlert("No inventory to display.");
  }
}