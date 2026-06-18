// CONFIG V1.02

// ==========================================
// CONFIGURATION & GLOBAL CONSTANTS
// ==========================================
// Constants for Tool Header rows and Table start rows, logic-based sorting order, and cross-sheet header aliases.
const RT_SHEET_NAME = "Receiving Tool";
const RT_HEADER_ROW = 9;
const RT_START_ROW  = 10;

const KT_SHEET_NAME = "Kitting Tool";
const KT_HEADER_ROW = 9;
const KT_START_ROW  = 10;

// Category sorting order.
const CAT_SORT_ORDER = {
  "Pipe": 1, "Flange": 2, "Grayloc": 7, "Fittings": 3, "Valve": 4,
  "Instrument": 8, "Support": 5, "Bolt-Up & Gaskets": 6, "Misc": 12,
  "Electrical": 9, "Structural": 11, "Flange Protector": 10
};

// Centralized cross-sheet header alias map.
const HEADER_ALIASES = {
  "Item Description": ["Item Description", "Item"],
  "Qty":              ["Qty", "Qnty", "Quantity"],
  "Heat #":           ["Heat #", "Heat Number"],
  "Date Logged":      ["Date Logged", "Date Entered", "Date Form"],
  "Location":         ["Location", "Full Location"],
  "Receiver/Kitter":  ["Receiver/Kitter", "Receiver"]
};

// ==========================================
// WEB APP GATEWAY & CONFIG
// ==========================================
// Houses the Master Log ID and serves the front-end WMS_Dashboard HTML while allowing for server-side file includes.
const MASTER_LOG_ID = "1sjbmxsfP18LrZreYse0ii2EAJlV_v4ps6c6HvB8SkJ0";

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('WMS_Dashboard');
  return template.evaluate()
      .setTitle('Materials Tool Suite - WMS')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// CUSTOM UI MENU (SAFE ADDITION)
// ==========================================
// Builds the dropdown menu in the Google Sheet UI for manual admin triggers like GSID updates or history healing.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ Admin Tools')
    .addItem('Update GSID Database', 'updateGSIDDatabase')
    .addItem('Bidirectional Heal (MMT ⇆ KT History)', 'syncKittingHistoryAndMMTs')
    .addToUi();
}

// ==========================================
// CORE DRY HELPERS
// ==========================================
// Lightweight, reusable utilities used globally for formatting headers, finding safe rows, fetching inputs, and triggering alerts.
function parseCoordStr(str) {
  if (!str || str === "❌" || !str.includes("R")) return null;
  const rMatch = str.match(/R(\d+)/);
  const cMatch = str.match(/\((\d+)\)/);
  if (rMatch && cMatch) {
    return {
      headerRow:    parseInt(rMatch[1], 10),
      dataRowStart: parseInt(rMatch[1], 10) + 1,
      colIdx:       parseInt(cMatch[1], 10) - 1
    };
  }
  return null;
}

function getSafeNextRow(sheet, colIdx, startRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow;
  const numRows = lastRow - startRow + 1;
  const data = sheet.getRange(startRow, colIdx, numRows, 1).getValues();
  for (let r = data.length - 1; r >= 0; r--) {
    if (data[r][0] !== null && data[r][0].toString().trim() !== "") {
      return startRow + r + 1;
    }
  }
  return startRow;
}

function showAlert(message) {
  if (SpreadsheetApp.getUi) {
    try { SpreadsheetApp.getUi().alert(message); } catch (e) {}
  }
}

function sanitizeHeaders(row) {
  return row.map(h => h.toString().replace(/\s+/g, ' ').trim());
}

function getInputValue(sheet, rangeName) {
  try {
    const range = sheet.getRange(rangeName);
    return range ? range.getDisplayValue().trim() : "";
  } catch (e) {
    return "";
  }
}

function clearNamedRanges(sheet, rangeNames) {
  rangeNames.forEach(name => {
    try { sheet.getRange(name).clearContent(); } catch (e) {}
  });
}

function findCol(headers, possibleNames) {
  for (const name of possibleNames) {
    const idx = headers.indexOf(name);
    if (idx > -1) return idx;
  }
  return -1;
}

// ==========================================
// UPGRADED: DYNAMIC GSID PARSER
// ==========================================
// Parses the string coordinates from the GSID Database into usable objects so scripts can dynamically locate moving columns.
function getJobCoordinatesFromGSID(jobNum) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gsidSheet = ss.getSheetByName("GSID Database");
  if (!gsidSheet) return null;
 
  const data = gsidSheet.getDataRange().getDisplayValues();
  if (data.length < 2) return null;
 
  const headers = sanitizeHeaders(data[0]);
  const cleanJob = jobNum.toString().trim().toUpperCase();

  // Dynamically map the columns by header name!
  const colMap = {
    vistaDesc:       findCol(headers, ["MMT VISTA Col: Item Desc"]),
    vistaBom:        findCol(headers, ["MMT VISTA Col: BOM ID"]),
    vistaQtyOrdered: findCol(headers, ["MMT VISTA Col: QTY Ordered"]), // NEW: Added Ordered
    vistaQtyRecv:    findCol(headers, ["MMT VISTA Col: QTY Recv"]),
    vistaQtyDue:     findCol(headers, ["MMT VISTA Col: QTY Due"]),
    vistaPo:         findCol(headers, ["MMT VISTA Col: PO #"]),
    vistaPoLine:     findCol(headers, ["MMT VISTA Col: PO Line #"]),
    fabItem:         findCol(headers, ["MMT FAB Col: Item No."]),
    fabQty:          findCol(headers, ["MMT FAB Col: Qty_mm"]),
    fabDesc:         findCol(headers, ["MMT FAB Col: Combined Material"]),
    fabBom:          findCol(headers, ["MMT FAB Col: BOM ID"]),
    fabKitted:       findCol(headers, ["MMT FAB Col: Kitted"]),
    fabDraw:         findCol(headers, ["MMT FAB Col: Drawing #"]),
    assmKitted:      findCol(headers, ["MMT ASSM Col: Kitted"]),
    assmDraw:        findCol(headers, ["MMT ASSM Col: Drawing #"]),
    assmQty:         findCol(headers, ["MMT ASSM Col: Qty_mm"]),
    assmDesc:        findCol(headers, ["MMT ASSM Col: Combined Material"]),
    assmBom:         findCol(headers, ["MMT ASSM Col: BOM ID"]),
    qcprSpool:       findCol(headers, ["QCPR Fab Col: Spool #"]),
    qcprInch:        findCol(headers, ["QCPR Fab Col: Inch Count"]),
    qcprIssued:      findCol(headers, ["QCPR Fab Col: Issued to Shop"]),
    qcprPriority:    findCol(headers, ["QCPR Col: Priority"]),
    qcprSize:        findCol(headers, ["QCPR Col: Size"]),
    qcprRevNotes:    findCol(headers, ["QCPR Fab Col: Revision Notes"]),
    qcprFabNotes:    findCol(headers, ["QCPR Fab Col: Fabrication Notes"])
  };

  for (let r = 1; r < data.length; r++) {
    if (data[r][0].toString().trim().toUpperCase() === cleanJob) {
      return {
        vistaDesc:       colMap.vistaDesc > -1 ? parseCoordStr(data[r][colMap.vistaDesc]) : null,
        vistaBom:        colMap.vistaBom > -1 ? parseCoordStr(data[r][colMap.vistaBom]) : null,
        vistaQtyOrdered: colMap.vistaQtyOrdered > -1 ? parseCoordStr(data[r][colMap.vistaQtyOrdered]) : null,
        vistaQtyRecv:    colMap.vistaQtyRecv > -1 ? parseCoordStr(data[r][colMap.vistaQtyRecv]) : null,
        vistaQtyDue:     colMap.vistaQtyDue > -1 ? parseCoordStr(data[r][colMap.vistaQtyDue]) : null,
        vistaPo:         colMap.vistaPo > -1 ? parseCoordStr(data[r][colMap.vistaPo]) : null,
        vistaPoLine:     colMap.vistaPoLine > -1 ? parseCoordStr(data[r][colMap.vistaPoLine]) : null,
        fabItem:         colMap.fabItem > -1 ? parseCoordStr(data[r][colMap.fabItem]) : null,
        fabQty:          colMap.fabQty > -1 ? parseCoordStr(data[r][colMap.fabQty]) : null,
        fabDesc:         colMap.fabDesc > -1 ? parseCoordStr(data[r][colMap.fabDesc]) : null,
        fabBom:          colMap.fabBom > -1 ? parseCoordStr(data[r][colMap.fabBom]) : null,
        fabKitted:       colMap.fabKitted > -1 ? parseCoordStr(data[r][colMap.fabKitted]) : null,
        fabDraw:         colMap.fabDraw > -1 ? parseCoordStr(data[r][colMap.fabDraw]) : null,
        assmKitted:      colMap.assmKitted > -1 ? parseCoordStr(data[r][colMap.assmKitted]) : null,
        assmDraw:        colMap.assmDraw > -1 ? parseCoordStr(data[r][colMap.assmDraw]) : null,
        assmQty:         colMap.assmQty > -1 ? parseCoordStr(data[r][colMap.assmQty]) : null,
        assmDesc:        colMap.assmDesc > -1 ? parseCoordStr(data[r][colMap.assmDesc]) : null,
        assmBom:         colMap.assmBom > -1 ? parseCoordStr(data[r][colMap.assmBom]) : null,
        qcprSpool:       colMap.qcprSpool > -1 ? parseCoordStr(data[r][colMap.qcprSpool]) : null,
        qcprInch:        colMap.qcprInch > -1 ? parseCoordStr(data[r][colMap.qcprInch]) : null,
        qcprIssued:      colMap.qcprIssued > -1 ? parseCoordStr(data[r][colMap.qcprIssued]) : null,
        qcprPriority:    colMap.qcprPriority > -1 ? parseCoordStr(data[r][colMap.qcprPriority]) : null,
        qcprSize:        colMap.qcprSize > -1 ? parseCoordStr(data[r][colMap.qcprSize]) : null,
        qcprRevNotes:    colMap.qcprRevNotes > -1 ? parseCoordStr(data[r][colMap.qcprRevNotes]) : null,
        qcprFabNotes:    colMap.qcprFabNotes > -1 ? parseCoordStr(data[r][colMap.qcprFabNotes]) : null
      };
    }
  }
  return null;
}