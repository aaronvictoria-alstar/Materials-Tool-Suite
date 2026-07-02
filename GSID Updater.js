// GSID UPDATER V1.02

// ==========================================
// BACKGROUND NIGHTLY JOB: GSID DATABASE UPDATER
// ==========================================
// Scans Drive for all MMT and QCPR files listed in the database. Dynamically hunts down exact header coordinates to prevent broken tools.
function updateGSIDDatabase() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const gsidSheet = ss.getSheetByName("GSID Database");
  if (!gsidSheet) return;

  const lastRow = gsidSheet.getLastRow();
  if (lastRow < 2) return;

  const findHeaderCoord = (dataGrid, searchKeys, strict = false) => {
    for (let r = dataGrid.length - 1; r >= 0; r--) {
      for (let c = 0; c < dataGrid[r].length; c++) {
        const cellVal = (dataGrid[r][c] || "").toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
        if (searchKeys.some(key => strict ? cellVal === key : cellVal.includes(key))) {
          return `R${r + 1}, Col ${getColLetter(c + 1)} (${c + 1})`;
        }
      }
    }
    return "❌";
  };
 
  const data    = gsidSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const newData = [];
  for (let i = 0; i < data.length; i++) {
    const jobNum        = data[i][0] ? data[i][0].toString().trim() : "";
    const existingInvId = data[i][3] ? data[i][3].toString().trim() : "";
    const clientMan     = data[i][4] ? data[i][4].toString().trim() : "";
    const clientAuto    = data[i][5] ? data[i][5].toString().trim() : "";
    if (!jobNum) {
      newData.push(new Array(39).fill(""));
      continue;
    }

    let mmtValid = "❌", qcprValid = "❌";
    let mmtMatDataValid = "❌", mmtTotalReqValid = "❌", mmtQtyKittedValid = "❌";
    let vistaBomValid = "❌", vistaDescValid = "❌", vistaBomIdValid = "❌", vistaQtyOrderedValid = "❌", vistaQtyRecvValid = "❌", vistaQtyDueValid = "❌", vistaPoValid = "❌", vistaPoLineValid = "❌";
    let itemFabValid = "❌", fabItemNoValid = "❌", fabQtyValid = "❌", fabCombMatValid = "❌", fabBomIdValid = "❌", fabKittedValid = "❌", fabDrawingValid = "❌";
    let itemAssmValid = "❌", assmKittedValid = "❌", assmDrawingValid = "❌";
    let assmQtyValid = "❌", assmCombMatValid = "❌", assmBomIdValid = "❌";
    let fabDataValid = "❌", spoolValid = "❌", inchValid = "❌", fabIssuedValid = "❌", qcprPriorityValid = "❌", qcprSizeValid = "❌";
    let qcprRevNotesValid = "❌", qcprFabNotesValid = "❌";
   
    // --- MMT Search & Verification ---
    const mmtQuery = `title contains '${jobNum}' and title contains 'Master Material Tracker' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`;
    const mmtFiles = DriveApp.searchFiles(mmtQuery);
    let mmtId = "", mmtTime = 0, mmtName = "❌", mmtUrl = "";
    while (mmtFiles.hasNext()) {
      const file  = mmtFiles.next();
      const fTime = file.getLastUpdated().getTime();
      if (fTime > mmtTime) { mmtTime = fTime; mmtId = file.getId(); mmtName = file.getName(); mmtUrl = file.getUrl(); }
    }

    if (!mmtId) {
      mmtId = "No sheet found";
    } else {
      mmtValid = `=HYPERLINK("${mmtUrl}", "${mmtName.replace(/"/g, '""')}")`;

      try {
        const mmtSS = SpreadsheetApp.openById(mmtId);

        // VISTA BOM CHECK
        const vSheet = mmtSS.getSheetByName("VISTA BOM");
        if (vSheet) {
          vistaBomValid = "✅";
          const vRow = vSheet.getLastRow(), vCol = vSheet.getLastColumn();
          if (vRow > 0 && vCol > 0) {
            const vTopRows     = vSheet.getRange(1, 1, Math.min(12, vRow), vCol).getValues();
            vistaDescValid     = findHeaderCoord(vTopRows, ["ITEMDESCRIPTION"]);
            vistaBomIdValid    = findHeaderCoord(vTopRows, ["BOMID"]);
            vistaQtyOrderedValid = findHeaderCoord(vTopRows, ["QTYORDERED", "ORDERQTY", "ORDEREDQTY", "QTYORDER"]); // NEW MAPPING
            vistaQtyRecvValid  = findHeaderCoord(vTopRows, ["QTYRECEIVED"]);
            vistaQtyDueValid   = findHeaderCoord(vTopRows, ["QTYDUE"]);
            vistaPoValid       = findHeaderCoord(vTopRows, ["PO"], true);
            vistaPoLineValid   = findHeaderCoord(vTopRows, ["POLINE"]);
           
            // If any VISTA column is missing, downgrade to Caution
            if ([vistaDescValid, vistaBomIdValid, vistaQtyOrderedValid, vistaQtyRecvValid, vistaQtyDueValid, vistaPoValid, vistaPoLineValid].includes("❌")) {
              vistaBomValid = "⚠️";
            }
          } else {
            vistaBomValid = "⚠️"; // Sheet exists but is totally empty
          }
        }

        // FAB CHECK
        const fabSheet = mmtSS.getSheetByName("Item Report-FAB");
        if (fabSheet) {
          itemFabValid = "✅";
          const fRow = fabSheet.getLastRow(), fCol = fabSheet.getLastColumn();
          if (fRow > 0 && fCol > 0) {
            const topRows   = fabSheet.getRange(1, 1, Math.min(12, fRow), fCol).getValues();
            fabItemNoValid  = findHeaderCoord(topRows, ["ITEMNO"]);
            fabQtyValid     = findHeaderCoord(topRows, ["QTYMM"]);
            fabCombMatValid = findHeaderCoord(topRows, ["COMBINEDMATERIAL"]);
            fabBomIdValid   = findHeaderCoord(topRows, ["BOMID"]);
            fabKittedValid  = findHeaderCoord(topRows, ["KITTEDISSUED", "KITTEDANDISSUED"]);
            fabDrawingValid = findHeaderCoord(topRows, ["DRAWING"]);

            if ([fabItemNoValid, fabQtyValid, fabCombMatValid, fabBomIdValid, fabKittedValid, fabDrawingValid].includes("❌")) {
              itemFabValid = "⚠️";
            }
          } else {
            itemFabValid = "⚠️";
          }
        }

        // MASTER MATERIAL DATA CHECK
        const matDataSheet = mmtSS.getSheetByName("Master Material Data");
        if (matDataSheet) {
          mmtMatDataValid = "✅";
          const mRow = matDataSheet.getLastRow(), mCol = matDataSheet.getLastColumn();
          if (mRow > 0 && mCol > 0) {
            const mTopRows = matDataSheet.getRange(1, 1, Math.min(20, mRow), mCol).getValues();
            mmtTotalReqValid  = findHeaderCoord(mTopRows, ["TOTALREQUIREDONALLDRAWINGS", "TOTALREQUIRED", "TOTALREQ"]);
            mmtQtyKittedValid = findHeaderCoord(mTopRows, ["KITTEDISSUEDQTY", "KITTEDISSUED", "KITTEDQTY", "QTYKITTED"]);
            if ([mmtTotalReqValid, mmtQtyKittedValid].includes("❌")) {
              mmtMatDataValid = "⚠️";
            }
          } else {
            mmtMatDataValid = "⚠️";
          }
        }

        // ASSM CHECK
        const assmSheet = mmtSS.getSheetByName("Item Report-ASSM");
        if (assmSheet) {
          itemAssmValid = "✅";
          const aRow = assmSheet.getLastRow(), aCol = assmSheet.getLastColumn();
          if (aRow > 0 && aCol > 0) {
            const aTopRows    = assmSheet.getRange(1, 1, Math.min(12, aRow), aCol).getValues();
            assmKittedValid   = findHeaderCoord(aTopRows, ["KITTEDISSUED", "KITTEDANDISSUED"]);
            assmDrawingValid  = findHeaderCoord(aTopRows, ["DRAWING"]);
            assmQtyValid      = findHeaderCoord(aTopRows, ["QTYMM"]);
            assmCombMatValid  = findHeaderCoord(aTopRows, ["COMBINEDMATERIAL"]);
            assmBomIdValid    = findHeaderCoord(aTopRows, ["BOMID"]);

            if ([assmKittedValid, assmDrawingValid, assmQtyValid, assmCombMatValid, assmBomIdValid].includes("❌")) {
              itemAssmValid = "⚠️";
            }
          } else {
            itemAssmValid = "⚠️";
          }
        }
      } catch (e) {}
    }

    // --- QCPR Search & Verification ---
    const qcprQuery = `title contains '${jobNum}' and title contains 'QCPR' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`;
    const qcprFiles = DriveApp.searchFiles(qcprQuery);
    let qcprId = "", qcprTime = 0, qcprName = "❌", qcprUrl = "";
    while (qcprFiles.hasNext()) {
      const file  = qcprFiles.next();
      const fTime = file.getLastUpdated().getTime();
      if (fTime > qcprTime) { qcprTime = fTime; qcprId = file.getId(); qcprName = file.getName(); qcprUrl = file.getUrl(); }
    }

    if (!qcprId) {
      qcprId = "No sheet found";
    } else {
      qcprValid = `=HYPERLINK("${qcprUrl}", "${qcprName.replace(/"/g, '""')}")`;

      try {
        const qcprSS    = SpreadsheetApp.openById(qcprId);
       
        // QCPR FAB DATA CHECK
        const qFabSheet = qcprSS.getSheetByName("Fab Data");
        if (qFabSheet) {
          fabDataValid = "✅";
          const qRow = qFabSheet.getLastRow(), qCol = qFabSheet.getLastColumn();
          if (qRow > 0 && qCol > 0) {
            const topRows       = qFabSheet.getRange(1, 1, Math.min(12, qRow), qCol).getValues();
            spoolValid          = findHeaderCoord(topRows, ["SPOOL"]);
            inchValid           = findHeaderCoord(topRows, ["INCH", "DIAMETER"]);
            fabIssuedValid      = findHeaderCoord(topRows, ["ISSUEDTOSHOP"]);
            qcprPriorityValid   = findHeaderCoord(topRows, ["PRIORITYNAME", "PRIORITY"]);
            qcprSizeValid       = findHeaderCoord(topRows, ["MAINNPSONSPOOL", "MAINNPS"]);
            qcprRevNotesValid   = findHeaderCoord(topRows, ["REVISIONNOTES", "REVNOTES"]);
            qcprFabNotesValid   = findHeaderCoord(topRows, ["FABRICATIONNOTES", "FABNOTES"]);

            if ([spoolValid, inchValid, fabIssuedValid, qcprPriorityValid, qcprSizeValid, qcprRevNotesValid, qcprFabNotesValid].includes("❌")) {
              fabDataValid = "⚠️";
            }
          } else {
            fabDataValid = "⚠️";
          }
        }
      } catch (e) {}
    }

    newData.push([
      mmtId, qcprId, existingInvId, clientMan, clientAuto,
      mmtValid, qcprValid,
      mmtMatDataValid, mmtTotalReqValid, mmtQtyKittedValid,
      vistaBomValid, vistaDescValid, vistaBomIdValid, vistaQtyOrderedValid, vistaQtyRecvValid, vistaQtyDueValid, vistaPoValid, vistaPoLineValid,
      itemFabValid, fabItemNoValid, fabQtyValid, fabCombMatValid, fabBomIdValid, fabKittedValid, fabDrawingValid,
      itemAssmValid, assmKittedValid, assmDrawingValid, assmQtyValid, assmCombMatValid, assmBomIdValid,
      fabDataValid, spoolValid, inchValid, fabIssuedValid, qcprPriorityValid, qcprSizeValid, qcprRevNotesValid, qcprFabNotesValid
    ]);
  }

  gsidSheet.getRange(2, 2, newData.length, 39).setValues(newData);
}

// Programmatically verifies and constructs a daily nightly background trigger for the GSID database
function ensureDailyGSIDTrigger() {
  const triggerName = "updateGSIDDatabase";
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === triggerName);
  
  if (!exists) {
    ScriptApp.newTrigger(triggerName)
      .timeBased()
      .everyDays(1)
      .atHour(2) // Run nightly at 2:00 AM (safe, non-disruptive time)
      .create();
  }
}
