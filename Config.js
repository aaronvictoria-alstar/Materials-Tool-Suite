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

// ==========================================
// CORE DRY HELPERS
// ==========================================
// Lightweight, reusable utilities used globally for formatting headers, finding safe rows, fetching inputs, and triggering alerts.

// Converts column index numbers (e.g. 1, 27) directly into standard spreadsheet A1 notation column letters (e.g. A, AA)
function getColLetter(colIdx) {
  let temp, letter = '';
  while (colIdx > 0) {
    temp   = (colIdx - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colIdx = (colIdx - temp - 1) / 26;
  }
  return letter;
}

// Normalizes and cleans drawing codes by stripping alphabetic prefixes and leading/trailing whitespace
function cleanDrawingNumber(rawDraw) {
  if (!rawDraw && rawDraw !== 0) return "";
  return rawDraw.toString().toUpperCase().trim().replace(/^[A-Z]+\s+/i, "");
}

// Converts millimeter measurements from VISTA databases directly into yards/feet
function convertMmToFt(mmVal) {
  return (parseFloat(mmVal) || 0) / 304.8;
}

// Applies the standardized 10% Pipe Tolerance Rule to reconcile yard receiving totals against MMT records
function isWithinPipeTolerance(val1, val2) {
  const v1 = parseFloat(val1) || 0;
  const v2 = parseFloat(val2) || 0;
  const diff = Math.abs(v1 - v2);
  const tolerance = Math.max(v1, v2) * 0.10 + 0.05;
  return diff <= tolerance;
}

// Normalizes and formats raw multi-line strings into standard newline-delimited lists
function formatMultiLine(str) {
  if (!str && str !== 0) return "";
  return str.toString().split(/[\n,;]+/).map(s => s.trim()).filter(Boolean).join("\n");
}

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
    mmtTotalReq:     findCol(headers, ["MMT Material Col: Total Req"]),
    mmtQtyKitted:    findCol(headers, ["MMT Material Col: QTY Kitted"]),
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
        mmtTotalReq:     colMap.mmtTotalReq > -1 ? parseCoordStr(data[r][colMap.mmtTotalReq]) : null,
        mmtQtyKitted:    colMap.mmtQtyKitted > -1 ? parseCoordStr(data[r][colMap.mmtQtyKitted]) : null,
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

// Returns true for inventory row types that represent a physical receipt from a vendor.
// Used in both PT (grossRecv accumulation) and RT (localRecvThisPo accumulation).
// NOTE: KITTED is intentionally excluded — it is handled separately via effectiveRecv
// in the PT display layer (max(grossRecv, kittedQty)) to avoid double-counting when
// both a REC row and a KITTED row exist for the same item.
function isGrossReceipt(type) {
  const t = type ? type.toString().trim().toUpperCase() : "";
  return t === "REC" || t === "SURPLUS" || t === "TRANSFER IN" || t === "RETURN";
}

// Returns true for inventory row types that represent a kitted/issued release.
// Centralised here so the multi-variant type string is defined in one place.
function isKittedType(type) {
  const t = type ? type.toString().trim().toUpperCase() : "";
  return t === "KITTED // ISSUED" || t === "KITTED" || t === "KIT" || t === "ISSUED" || t === "ISSUE";
}

// ==========================================
// UPGRADED INTELLIGENT MATCHING & DE-DUPLICATION HELPERS
// ==========================================

// Normalizes descriptions by removing drawing tags, circuit IDs, context suffixes (C/W, WITH, FOR) and double spacing
function normalizeDescription(desc) {
  if (!desc && desc !== 0) return "";
  let d = desc.toString().toUpperCase().trim();
  // Replace multiple spaces with single space
  d = d.replace(/\s+/g, " ");
  // Strip parenthetical pure-numeric values (e.g., wall thickness "(21.41)" or "(0.500)")
  d = d.replace(/\(\s*[\d.]+\s*\)/g, "");
  // Remove TAG / CIRCUIT ID patterns: TAG #123, TAG 123, CIRCUIT ID: 123, etc.
  d = d.replace(/(?:TAG|CIRCUIT ID)[\s#:]+([A-Z0-9\-]+)/g, "");
  // Remove context suffixes starting with C/W, W/, WITH, FOR
  d = d.replace(/(?:\bC\/W\b|\bW\/\b|\bWITH\b|\bFOR\b)(.*)/, "");
  // Standardize spaces again after replacements
  d = d.replace(/\s+/g, " ");
  return d.trim();
}

// Generates a unified matching key. Groups by BOM ID if present, otherwise groups by normalized description.
function getUnifiedItemKey(bomId, description) {
  const cleanBom = (bomId || "").toString().trim().toUpperCase();
  const cleanDesc = normalizeDescription(description);
  const safeDesc = cleanDesc.replace(/[^A-Z0-9]/g, "_");
  if (cleanBom) {
    return cleanBom + "_" + safeDesc;
  }
  // If no BOM ID, group by normalized description
  return "NOBOM_" + safeDesc;
}

// ==========================================
// CATEGORY RULE LOADERS
// ==========================================
// Two lazy-loaded lookup tables, each read once per script execution and reused across all
// getCategoryLogic calls. Learned Rules are permanent promoted corrections; Category Overrides
// are the nightly-wiped inbox for newly detected discrepancies.
let _learnedRules     = null;
let _categoryOverrides = null;

function _loadLearnedRules() {
  if (_learnedRules !== null) return _learnedRules;
  _learnedRules = new Map();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Learned Rules");
    if (sheet && sheet.getLastRow() >= 2) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
      for (const row of data) {
        const key = row[0] ? row[0].toString().trim() : "";
        const cat = row[3] ? row[3].toString().trim() : "";
        const sub = row[4] ? row[4].toString().trim() : "";
        if (key && cat) _learnedRules.set(key, { category: cat, subcat: sub });
      }
    }
  } catch(e) {}
  return _learnedRules;
}

function _loadCategoryOverrides() {
  if (_categoryOverrides !== null) return _categoryOverrides;
  _categoryOverrides = new Map();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Category Overrides");
    if (sheet && sheet.getLastRow() >= 2) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
      for (const row of data) {
        const key = row[0] ? row[0].toString().trim() : "";
        const cat = row[3] ? row[3].toString().trim() : "";
        const sub = row[4] ? row[4].toString().trim() : "";
        if (key && cat) _categoryOverrides.set(key, { category: cat, subcat: sub });
      }
    }
  } catch(e) {}
  return _categoryOverrides;
}

// ==========================================
// INTELLIGENT CATEGORY LOGIC
// ==========================================
// Massively expanded categorization engine. Parses raw material descriptions and BOM IDs to intelligently assign master categories, subcategories, and pipe sizes.
function getCategoryLogic(bomId, description) {
  const bUpper = (bomId || "").toString().toUpperCase().trim();
  const dUpper = (description || "").toString().toUpperCase().trim();
  let cat = "Misc";
  let sub = "Misc";




  const sizeMatch = dUpper.match(/(?:^|[^0-9])(\d+(?:\.\d+)?)(?:\s+\d+\/\d+)?\s*(?:\"|IN|INCH)/);
  const size = sizeMatch ? parseFloat(sizeMatch[1]) : 0;

  // Lookup order: Learned Rules (permanent) → Category Overrides (nightly inbox) → algorithm
  const _key = bUpper || ("NOBOM_" + normalizeDescription(description).replace(/[^A-Z0-9]/g, "_"));
  const _lr = _loadLearnedRules();
  if (_lr.has(_key)) {
    const m = _lr.get(_key);
    return { category: m.category, subcat: m.subcat || "Misc", size };
  }
  const _ov = _loadCategoryOverrides();
  if (_ov.has(_key)) {
    const m = _ov.get(_key);
    return { category: m.category, subcat: m.subcat || "Misc", size };
  }

  let tag = "";
  const tagMatch = dUpper.match(/(?:TAG|CIRCUIT ID)[\s#:]+([A-Z0-9\-]+)/);
  if (tagMatch) tag = tagMatch[1];




  let primaryDesc = dUpper;
  const contextMatch = dUpper.match(/(.*?)(?:\bC\/W\b|\bW\/\b|\bWITH\b|\bFOR\b)(.*)/);
  if (contextMatch && contextMatch[1].trim().length > 2) primaryDesc = contextMatch[1].trim();




  const isSW   = dUpper.match(/\bSW\b/) || dUpper.match(/\bSOCK/) || dUpper.includes("RFSW") || dUpper.includes("RTJSW") || dUpper.includes("FFSW");
  const isTH   = dUpper.match(/\bTH\b/) || dUpper.match(/\bNPT\b/) || dUpper.match(/\bMNPT\b/) || dUpper.match(/\bFNPT\b/) || dUpper.includes("THREAD") || dUpper.includes("THD") || dUpper.includes("RFTH") || dUpper.includes("RTJTH") || dUpper.includes("FFTH");
  const face   = (dUpper.includes("FLAT FACE") || dUpper.match(/\bFF\b/)) ? "FF" : (dUpper.includes("RTJ") ? "RTJ" : "RF");
  const isTapped = dUpper.includes("TAPPED") || dUpper.includes("TAP TO") || dUpper.match(/\bTAP\b/);




  const getInstSub = () => {
    if (dUpper.includes("ROTOM") || dUpper.includes("ROTAM")) return "Rotometer";
    else if (dUpper.includes("H2S") || dUpper.includes("LEL") || dUpper.includes("DETECTOR") || tag.match(/^(H2S|LEL)/)) return "Detector";
    else if (dUpper.includes("PRESSURE TRANSMITTER") || tag.includes("-PT-") || tag.includes("-PIT-") || tag.match(/^(PT|PIT)/)) return "PT";
    else if (dUpper.includes("TEMPERATURE TRANSMITTER") || tag.includes("-TT-") || tag.includes("-TIT-") || tag.match(/^(TT|TIT)/)) return "TT";
    else if (dUpper.includes("FLOW METER") || tag.includes("-FT-") || tag.includes("-FIT-") || tag.includes("-FE-") || tag.match(/^(FT|FIT|FE)/)) return "FIT";
    else if (dUpper.includes("THERMOWELL") || tag.includes("-TW-") || tag.match(/^TW/)) return "TW";
    else if (dUpper.includes("PRESSURE GAUGE") || dUpper.includes("KPA") || dUpper.includes("MPA") || tag.includes("-PG-") || tag.match(/^PG/)) return "PG";
    else if (dUpper.includes("TEMP GAUGE") || dUpper.includes("TEMPERATURE GAUGE") || tag.includes("-TG-") || tag.match(/^TG/)) return "TG";
    else if (dUpper.includes("FLOW GAUGE") || tag.includes("-FG-") || tag.match(/^FG/)) return "FG";
    else if (dUpper.includes("RTD") || dUpper.includes("SENSOR") || tag.includes("-TE-") || tag.match(/^TE/)) return "TE";
    else if (dUpper.includes("SWITCH") || tag.includes("PSHH") || tag.match(/^PSHH/)) return "Pressure Switch";
    else if (dUpper.includes("RUPTURE") || tag.includes("PSE") || tag.match(/^PSE/)) return "PSE";
    else if (dUpper.includes("MANIFOLD")) return "Manifold";
    else if (dUpper.includes("DIAPHRAGM SEAL")) return "PT";
    return "Misc";
  };

  const getElecSub = () => {
    if (dUpper.includes("BEACON") || tag.startsWith("BCN")) return "Beacon";
    else if (dUpper.includes("HORN") || tag.startsWith("HRN")) return "Horn";
    else if (dUpper.includes("TRANSFORMER") || dUpper.includes("XFM") || tag.startsWith("XFM")) return "Transformer";
    else if (dUpper.includes("PHOTOCELL") || dUpper.includes("PHOTOCONTROL") || tag.startsWith("PC-")) return "Photocell";
    else if (dUpper.includes("RECEPTACLE") || tag.startsWith("REC-") || dUpper.includes("HUBHBL5262W")) return "Receptacle";
    else if (dUpper.includes("PANEL") || tag.startsWith("IOP") || tag.startsWith("DP-")) return "Panel";
    else if (dUpper.includes("JUNCTION BOX") || dUpper.includes("MIJB") || dUpper.includes("JBS-") || dUpper.includes("JBM-") || dUpper.includes("JBHV")) return "Junction Box";
    else if (dUpper.includes("STRAP") || dUpper.includes("PS-10") || dUpper.includes("PS-20") || dUpper.includes("PS-03")) return "Strap";
    else if (dUpper.includes("TIE WIRE") || dUpper.includes("SSTW")) return "Wire";
    else if (dUpper.includes("LABEL") || dUpper.includes("ETL-")) return "Labels";
    else if (dUpper.includes("BRACKET") || dUpper.includes("MIQ-BR") || dUpper.includes("UMB-T")) return "Bracket";
    else if (dUpper.includes("TAPE") || dUpper.includes("GS54") || dUpper.includes("TAPES-GS54")) return "Tape";
    else if (tag.startsWith("HTP")) return "HTP";
    else if (tag.startsWith("ET-") || dUpper.includes("CIRCUIT ID:")) return "ET";
    else if (tag.startsWith("PE-")) return "PE";
    else if (tag.startsWith("HS-")) return "HS";
    else if (tag.startsWith("GP-")) return "GP";
    else if (dUpper.includes("MOTOR")) return "Motor";
    else if (dUpper.includes("CABLE") || dUpper.includes("EHT") || dUpper.includes("XTV") || dUpper.includes("E-100") || dUpper.includes("NGC-")) return "Heat Trace";
    else if (dUpper.includes("CONDUIT") || dUpper.includes("CRHFS2SA") || dUpper.includes("INTERCONNECT")) return "Conduit";
    else if (dUpper.includes("LED ") || dUpper.includes("LUMINAIRE")) return "Lighting";
    return "Misc";
  };

  // PHASE 1: STRONG DESCRIPTION TEXT MATCHING
  if (bUpper.startsWith("PIPE") || dUpper.includes("PIPE SMLS") || dUpper.includes("BE X BE")) {
    cat = "Pipe";
    const schMatch = dUpper.match(/SCH(?:EDULE)?\s*(\d+)/);
    if (dUpper.includes("XXS")) sub = "Sch XXS";
    else if (dUpper.includes("XS") || dUpper.includes("XH") || (schMatch && schMatch[1] === "80")) {
      sub = ((dUpper.includes("XS") || dUpper.includes("XH")) && size >= 10) ? "Sch XS" : "Sch 80";
    }
    else if (dUpper.includes("STD") || (schMatch && schMatch[1] === "40")) {
      sub = (dUpper.includes("STD") && size > 10) ? "Sch STD" : "Sch 40";
    }
    else if (schMatch) sub = "Sch " + schMatch[1];
    else if (dUpper.match(/\b160\b/)) sub = "Sch 160";
    else if (dUpper.match(/\b120\b/)) sub = "Sch 120";
    else if (dUpper.match(/\b80\b/))  sub = "Sch 80";
    else if (dUpper.match(/\b60\b/))  sub = "Sch 60";
    else if (dUpper.match(/\b40\b/))  sub = "Sch 40";
    else if (dUpper.match(/\b20\b/))  sub = "Sch 20";
    else if (dUpper.match(/\b10\b/))  sub = "Sch 10";
  }
  else if (bUpper.includes("COVER") || dUpper.includes("PROTECTOR") || dUpper.includes("FLANGE COVER") ||
           (dUpper.includes("FLANGE") && dUpper.includes("COVER")) || dUpper.match(/\bEND CAP\b/)) {
    cat = "Flange Protector";
    sub = "Protector/Cover";
  }
  else if (dUpper.includes("GRAYLOC") || dUpper.includes("GRAYLOCK") || dUpper.includes("GREYLOC") || dUpper.includes("BLUESKY") ||
           dUpper.includes("ROTABALL") || dUpper.includes("ROTA BALL") || dUpper.includes("FLEXBALL") || dUpper.includes("FLEX BALL") || dUpper.includes("UNIBALL") ||
           bUpper.match(/^R\d+/) || dUpper.match(/\bR\d+\s+(HUB|SEAL|RING|BLIND)/) ||
           (bUpper.includes("GR") && (dUpper.includes("HUB") || dUpper.includes("CLAMP") || dUpper.includes("SEALRING")))) {
    cat = "Grayloc";
    if (dUpper.includes("HUB"))            sub = (dUpper.includes("BLIND") || dUpper.includes("BLND")) ? "Hub - Blind" : "Hub";
    else if (dUpper.includes("SEAL") || dUpper.includes("RING")) sub = "Seal Ring";
    else if (dUpper.includes("CLAMP"))     sub = "Clamp";
    else if (dUpper.includes("BLIND"))     sub = "Blind";
    else if (dUpper.includes("FLEX") || dUpper.includes("ROTA") || dUpper.includes("UNI") || dUpper.includes("BALL")) sub = "Flexball";
  }
  else if (dUpper.includes("TRANSMITTER") || dUpper.includes("METER") || dUpper.includes("THERMOWELL") ||
           dUpper.includes("GAUGE") || (dUpper.includes("SWITCH") && !dUpper.includes("LIMIT SWITCH")) ||
           dUpper.includes("RUPTURE DISK") || dUpper.includes("RTD") || dUpper.includes("SENSOR") ||
           dUpper.includes("DIAPHRAGM SEAL") || dUpper.includes("MANIFOLD") || dUpper.includes("DETECTOR") ||
           dUpper.includes("ROTOM") || dUpper.includes("ROTAM") || dUpper.includes("KPA") || dUpper.includes("MPA")) {
    cat = "Instrument";
    sub = getInstSub();
  }
  else if (dUpper.includes("EHT") || dUpper.includes("PANEL") || (dUpper.includes("CABLE") && !dUpper.includes("CABLE KIT")) ||
           dUpper.includes("VOLT") || dUpper.match(/\bVAC\b/) || dUpper.includes("MOTOR") ||
           dUpper.includes("CONDUIT") || dUpper.includes("JUNCTION BOX") || dUpper.includes("MIQ-") || dUpper.includes("E-100") ||
           dUpper.includes("JBS-") || dUpper.includes("JBM-") || dUpper.includes("STRAP") || dUpper.includes("TIE WIRE") ||
           dUpper.includes("LABEL") || dUpper.includes("BRACKET") || dUpper.includes("NGC-") || dUpper.includes("LED ") ||
           dUpper.includes("BEACON") || dUpper.includes("HORN") || dUpper.includes("XFM") || dUpper.includes("TRANSFORMER") ||
           dUpper.includes("INTERCONNECT") || dUpper.includes("PHOTOCELL") || dUpper.includes("PHOTOCONTROL") ||
           dUpper.includes("RECEPTACLE") || dUpper.includes("XTV") || dUpper.includes("CRHFS2SA")) {
    cat = "Electrical";
    sub = getElecSub();
  }
  else if (dUpper.includes("VALVE") || dUpper.includes("VLV") || dUpper.match(/\bPSV\b/) || dUpper.match(/\bPRV\b/) || dUpper.match(/\bTCV\b/) ||
           dUpper.includes("CHAIN OPERATOR") || dUpper.includes("CONNECTOR LINK") || dUpper.includes("ACTUATOR") || dUpper.includes("LIMIT SWITCH") ||
           dUpper.includes("SOLENOID") || dUpper.includes("REGULATOR") || dUpper.includes("BLOCK AND BLEED") || dUpper.includes("IV202") ||
           dUpper.includes("WIKA") || dUpper.includes("ADAPTER SWAGELOK") || dUpper.includes("HANDWHEEL") || dUpper.includes("HDWH")) {
    cat = "Valve";
    if      (primaryDesc.includes("GATE"))                                    sub = "Gate";
    else if (primaryDesc.includes("GLOBE"))                                   sub = "Globe";
    else if (primaryDesc.includes("CHECK") || primaryDesc.includes("CHK") || primaryDesc.includes("SWING")) sub = "Check";
    else if (primaryDesc.includes("BALL"))                                    sub = "Ball";
    else if (dUpper.includes("CONTROL") || dUpper.match(/\bTCV\b/) || dUpper.includes("ACTUATOR") || dUpper.includes("LIMIT SWITCH") || dUpper.includes("SOLENOID") || dUpper.includes("REGULATOR") || dUpper.includes("ADAPTER KIT") || dUpper.includes("WIKA") || dUpper.includes("AS BUILT") || dUpper.includes("SWAGELOK")) sub = "Control";
    else if (primaryDesc.includes("NEEDLE") || primaryDesc.includes("BLOCK AND BLEED") || primaryDesc.includes("IV202") || primaryDesc.includes("HEX NIPPLE")) sub = "Needle";
    else if (primaryDesc.match(/\bPSV\b/) || primaryDesc.match(/\bPRV\b/))   sub = "PSV";
    else if (primaryDesc.includes("CHOKE"))                                   sub = "Choke";
    else if (dUpper.includes("CHAIN") || dUpper.includes("LINK") || dUpper.includes("CABLE KIT") || dUpper.includes("HANDWHEEL") || dUpper.includes("HDWH")) sub = "Chain Operator";
    else if (primaryDesc.includes("VENTURI"))                                 sub = "Venturi";
    else if (dUpper.includes("GATE"))                                         sub = "Gate";
    else if (dUpper.includes("GLOBE"))                                        sub = "Globe";
    else if (dUpper.includes("CHECK") || dUpper.includes("CHK") || dUpper.includes("SWING")) sub = "Check";
    else if (dUpper.includes("BALL"))                                         sub = "Ball";
    else if (dUpper.includes("NEEDLE"))                                       sub = "Needle";
  }
  else if (dUpper.includes("SUPPORT") || dUpper.includes("SHOE") || dUpper.includes("GUIDE") || dUpper.includes("ANCHOR") ||
           dUpper.includes("WEAR PAD") || dUpper.includes("RE-PAD") || dUpper.includes("REPAD") ||
           dUpper.includes("GUSSET") || dUpper.includes("END PLATE") || dUpper.includes("TRUNNION") ||
           dUpper.includes("DUMMY") || dUpper.includes("BASE PLATE") || dUpper.includes("BASE SUPPORT") || dUpper.includes("CANTILEVER") ||
           dUpper.includes("LIFTING LUG") || dUpper.includes("SLIDE PLATE") || dUpper.includes("SIDE PLATE") ||
           dUpper.includes("REINFORCING PAD") || dUpper.includes("U-BOLT") || dUpper.includes("U BOLT") || dUpper.includes("PLATE ONLY") || dUpper.includes("FIXED PLATE")) {
    cat = "Support";
    if      (dUpper.includes("END PLATE ONLY") || dUpper.includes("PLATE ONLY"))   sub = "End Plate";
    else if (dUpper.includes("REPAD ONLY") || dUpper.includes("RE-PAD ONLY"))      sub = "Reinforcing Pad";
    else if (primaryDesc.includes("ANCHOR"))                                        sub = "Anchor";
    else if (primaryDesc.includes("GUIDE"))                                         sub = "Guide";
    else if (primaryDesc.includes("SHOE"))                                          sub = "Shoe";
    else if (primaryDesc.includes("WEAR PAD"))                                      sub = "Wear Pad";
    else if (primaryDesc.includes("TRUNNION"))                                      sub = "Support - Trunnion";
    else if (primaryDesc.includes("END PLATE"))                                     sub = "End Plate";
    else if (primaryDesc.includes("DUMMY"))                                         sub = "Dummy";
    else if (primaryDesc.includes("BASE PLATE") || primaryDesc.includes("BASE SUPPORT")) sub = "Support - Base";
    else if (primaryDesc.includes("CANTILEVER"))                                    sub = "Cantilever";
    else if (primaryDesc.includes("LIFTING LUG") || primaryDesc.includes("LUG") || primaryDesc.includes("SHEAR LUG")) sub = "Lug";
    else if (primaryDesc.includes("U-BOLT") || primaryDesc.includes("U BOLT"))     sub = "U-Bolt";
    else if (primaryDesc.includes("SLIDE PLATE") || primaryDesc.includes("SIDE PLATE")) sub = "Slide Plate";
    else if (primaryDesc.includes("PIPE SUPPORT") || primaryDesc.includes("FP4-A") || primaryDesc.includes("VERTICAL PIPE")) sub = "Support - Pipe";
    else if (primaryDesc.includes("RE-PAD") || primaryDesc.includes("REPAD") || primaryDesc.includes("REINFORCING PAD")) sub = "Reinforcing Pad";
    else if (primaryDesc.includes("GUSSET"))                                        sub = "Gusset";
    else if (primaryDesc.includes("FIXED PLATE"))                                   sub = "Fixed Plate";
    else if (dUpper.includes("ANCHOR"))                                             sub = "Anchor";
    else if (dUpper.includes("GUIDE"))                                              sub = "Guide";
    else if (dUpper.includes("SHOE"))                                               sub = "Shoe";
    else if (dUpper.includes("TRUNNION"))                                           sub = "Support - Trunnion";
    else if (dUpper.includes("END PLATE"))                                          sub = "End Plate";
    else if (dUpper.includes("DUMMY"))                                              sub = "Dummy";
    else if (dUpper.includes("BASE PLATE") || dUpper.includes("BASE SUPPORT"))     sub = "Support - Base";
    else if (dUpper.includes("RE-PAD") || dUpper.includes("REPAD") || dUpper.includes("REINFORCING PAD")) sub = "Reinforcing Pad";
    else if (dUpper.includes("FIXED PLATE"))                                        sub = "Fixed Plate";
  }
  else if (bUpper.startsWith("FLG") || bUpper.startsWith("BR") || bUpper.startsWith("SB") || bUpper.startsWith("PB") ||
           dUpper.includes("FLANGE") || dUpper.includes("BLEED RING") || dUpper.includes("SPECTACLE BLIND") ||
           dUpper.includes("PADDLE BLIND") || dUpper.includes("ORIF") || dUpper.includes("SPACER") ||
           dUpper.includes("WELD NECK") || dUpper.match(/\bWN\b/) || dUpper.includes("TAPPED BLIND") || dUpper.includes("LAP JOINT")) {
    cat = "Flange";
    const wnFormat = face === "RF" ? "RFWN" : `${face} WN`;
    if      (primaryDesc.includes("SPECTACLE") || bUpper.startsWith("SB"))    sub = "Spec Blind";
    else if (primaryDesc.includes("PADDLE") || bUpper.startsWith("PB"))       sub = "Paddle Blind";
    else if (primaryDesc.includes("BLIND") || primaryDesc.includes("BLND")) {
      if (isTapped) sub = isSW ? `${face} Tapped Blind SW` : `${face} Tapped Blind TH`;
      else          sub = `${face} Blind`;
    }
    else if (primaryDesc.includes("WELD NECK") || primaryDesc.includes("WN") || primaryDesc.match(/\bWN\b/)) {
      if      (dUpper.includes("ORIF") || dUpper.includes("ORIFICE"))            sub = `Orifice ${wnFormat}`;
      else if (dUpper.includes("JACKSCREW") || dUpper.includes("JACKSCREWS"))    sub = `${wnFormat} W/ Jackscrews`;
      else                                                                        sub = wnFormat;
    }
    else if (primaryDesc.includes("LAP JOINT") || primaryDesc.includes("LAP"))  sub = "Lap Joint";
    else if (primaryDesc.includes("ORIFICE"))                                    sub = "Orifice Plate";
    else if (primaryDesc.includes("SPACER"))                                     sub = "Spacer";
    else if (primaryDesc.includes("BLEED RING") || bUpper.startsWith("BR"))     sub = "Bleed Ring";
    else if (isSW)                                                               sub = `${face} SW`;
    else if (isTH)                                                               sub = `${face} TH`;
    else if (dUpper.includes("BLEED RING"))                                      sub = "Bleed Ring";
    else if (sub === "Misc" && (dUpper.includes("WELD NECK") || dUpper.match(/\bWN\b/))) sub = wnFormat;
    else if (sub === "Misc" && dUpper.includes("BLIND"))                         sub = `${face} Blind`;
  }
  else if (bUpper.startsWith("ELL") || bUpper.startsWith("TEE") || bUpper.startsWith("CAP") ||
           bUpper.startsWith("NIP") || bUpper.startsWith("WOL") || bUpper.startsWith("SOL") ||
           bUpper.startsWith("TOL") || bUpper.startsWith("FOL") || bUpper.startsWith("ER") || bUpper.startsWith("CR") ||
           bUpper.startsWith("CPLG") || bUpper.startsWith("UN") || bUpper.startsWith("HPLUG") || bUpper.startsWith("RPLUG") ||
           bUpper.startsWith("SWC") || bUpper.startsWith("SWE") || bUpper.startsWith("STUB") || bUpper.startsWith("CROSS") ||
           dUpper.includes("ELBOW") || dUpper.match(/\bTEE\b/) || dUpper.includes("REDUCER") ||
           dUpper.includes("NIPPLE") || dUpper.includes("OLET") || dUpper.includes("COUPLING") ||
           dUpper.includes("SWAGE") || dUpper.includes("STUB END") || dUpper.includes("PLUG") ||
           dUpper.includes("STRAINER") || dUpper.includes("INJECTION QUILL") || dUpper.includes("TRANSITION") ||
           dUpper.match(/\bCAPS\b/) ||
           dUpper.match(/\bWOL\b/) || dUpper.match(/\bSOL\b/) || dUpper.match(/\bTOL\b/) || dUpper.match(/\bFOL\b/) || dUpper.match(/\bCROSS\b/)) {
    cat = "Fittings";
    if (primaryDesc.includes("ELBOW") || bUpper.startsWith("ELL")) {
      if      (isSW)                  sub = "Elbow - SW";
      else if (isTH)                  sub = "Elbow - TH";
      else if (primaryDesc.includes("45")) sub = "Elbow - 45";
      else                            sub = "Elbow - 90";
    }
    else if (primaryDesc.match(/\bTEE\b/) || bUpper.startsWith("TEE")) {
      const red = (primaryDesc.match(/\bRED\b/) || primaryDesc.includes("REDUCING") || bUpper.includes("-RED")) ? " Red" : "";
      if      (primaryDesc.includes("BARRED")) sub = "Tee - Barred";
      else if (isSW)  sub = "Tee - SW" + red;
      else if (isTH)  sub = "Tee - TH" + red;
      else if (red)   sub = "Tee - Red";
      else            sub = "Tee";
    }
    else if (primaryDesc.match(/\bCROSS\b/) || bUpper.startsWith("CROSS")) {
      if      (isSW) sub = "Tee - Cross - SW";
      else if (isTH) sub = "Tee - Cross - TH";
      else           sub = "Tee - Cross";
    }
    else if (primaryDesc.includes("CAP") || bUpper.startsWith("CAP") || primaryDesc.match(/\bCAPS\b/)) {
      if      (isSW) sub = "Cap - SW";
      else if (isTH) sub = "Cap - TH";
      else           sub = "Cap - BW";
    }
    else if (primaryDesc.includes("NIPPLE") || bUpper.startsWith("NIP")) {
      if      (primaryDesc.includes("TOE") || primaryDesc.includes("POE")) sub = "Nipple - POE/TOE";
      else if (primaryDesc.includes("TBE"))                                sub = "Nipple - TBE";
      else                                                                  sub = "Nipple - PBE";
    }
    else if (primaryDesc.includes("REDUCER") || bUpper.startsWith("ER") || (bUpper.startsWith("CR") && !bUpper.startsWith("CROSS"))) {
      sub = (primaryDesc.includes("ECC") || bUpper.startsWith("ER")) ? "Reducer - Ecc" : "Reducer - Conc";
    }
    else if (primaryDesc.includes("SWAGE") || bUpper.startsWith("SWC") || bUpper.startsWith("SWE")) {
      sub = (primaryDesc.includes("ECC") || bUpper.startsWith("SWE")) ? "Swage - Ecc" : "Swage - Conc";
    }
    else if (bUpper.startsWith("WOL") || primaryDesc.includes("WELDOLET") || primaryDesc.match(/\bWOL\b/))   sub = "Olets - WOL";
    else if (bUpper.startsWith("SOL") || primaryDesc.includes("SOCKOLET") || primaryDesc.match(/\bSOL\b/))   sub = "Olets - SOL";
    else if (bUpper.startsWith("TOL") || primaryDesc.includes("THREADOLET") || primaryDesc.match(/\bTOL\b/)) sub = "Olets - TOL";
    else if (bUpper.startsWith("FOL") || primaryDesc.includes("FLATOLET") || primaryDesc.includes("FLAT-OLET") || primaryDesc.match(/\bFOL\b/)) sub = "Olets";
    else if (primaryDesc.includes("ELBOLET") || primaryDesc.includes("LATROLET")) sub = isSW ? "Olets - SW" : (isTH ? "Olets - TH" : "Olets");
    else if (primaryDesc.includes("COUPLING") || bUpper.startsWith("CPLG")) {
      if      (dUpper.includes("FNPT") && isSW) sub = "Coupling - SW FNPT";
      else if (isSW) sub = "Coupling - SW";
      else if (isTH) sub = "Coupling - TH";
      else           sub = "Coupling";
    }
    else if (primaryDesc.includes("UNION") || bUpper.startsWith("UN")) sub = isSW ? "Union - SW" : (isTH ? "Union - TH" : "Union");
    else if (primaryDesc.includes("PLUG") || bUpper.startsWith("HPLUG") || bUpper.startsWith("RPLUG")) {
      sub = (dUpper.includes("RND") || dUpper.includes("ROUND") || bUpper.startsWith("RPLUG")) ? "Plug - Rnd" : "Plug - Hex";
    }
    else if (primaryDesc.includes("STRAINER"))       sub = "Strainer";
    else if (primaryDesc.includes("INJECTION QUILL")) sub = "Injection Quill";
    else if (primaryDesc.includes("TRANSITION"))      sub = "Transition";
    if (sub === "Misc" && dUpper.match(/\bTEE\b/))    sub = "Tee";
    if (sub === "Misc" && dUpper.includes("ELBOW"))    sub = "Elbow - 90";
  }
  else if (dUpper.includes("A325") || dUpper.includes("STRUCTURAL BOLT") || dUpper.match(/\bGALV\b/) || dUpper.includes("GALVANIZED") || dUpper.match(/\bHDG\b/) ||
           (primaryDesc.match(/\bBOLT(S)?\b/) && !dUpper.includes("STUD") && !dUpper.includes("U-BOLT") && !dUpper.includes("U BOLT")) ||
           dUpper.includes("GRATING") || dUpper.includes("GRATING CLIPS")) {
    cat = "Structural";
    sub = (dUpper.includes("GRATING") || dUpper.includes("GRATING CLIPS")) ? "Grating" : "Bolts";
  }
  else if (bUpper.startsWith("GASK") || dUpper.includes("GASKET") || dUpper.includes("RING SOFT IRON")) {
    cat = "Bolt-Up & Gaskets";
    sub = "Gaskets";
  }
  else if (bUpper.startsWith("STUD") || bUpper.startsWith("BOLT") ||
           dUpper.includes("STUD ") || dUpper.includes("STUDS") || dUpper.includes("BOLT") ||
           dUpper.match(/\bNUT\b/) || dUpper.match(/\bNUTS\b/) ||
           dUpper.match(/\bWASHER\b/) || dUpper.match(/\bWASHERS\b/)) {
    cat = "Bolt-Up & Gaskets";
    sub = "Studs/Nuts";
  }
  else if ((bUpper.startsWith("W") && !bUpper.startsWith("WOL") && !bUpper.startsWith("WN")) ||
           bUpper.startsWith("HP") || bUpper.startsWith("HSS") ||
           bUpper.startsWith("L") || bUpper.startsWith("PL") || dUpper.includes("FLAT BAR") || dUpper.includes("ANGLE")) {
    cat = "Structural";
    if      ((bUpper.startsWith("W") && !bUpper.startsWith("WOL") && !bUpper.startsWith("WN")) || bUpper.startsWith("HP")) sub = "Beam";
    else if (bUpper.startsWith("HSS"))                                              sub = "HSS";
    else if (bUpper.startsWith("L") || dUpper.includes("ANGLE"))                   sub = "Angle";
    else if (bUpper.startsWith("PL") || dUpper.includes("FLAT BAR"))               sub = "Plate";
  }
  else if (dUpper.includes("SAFETY GATE") || dUpper.includes("CAMLOCK") || dUpper.includes("CAM-LOCK") ||
           dUpper.includes("HOSE") || dUpper.includes("EXPANSION JOINT") ||
           dUpper.includes("INSULATION KIT") || dUpper.includes("CORROSION COUPON") ||
           dUpper.includes("BLANKET") || dUpper.includes("FLUSH RING") || dUpper.includes("SYPHON") ||
           dUpper.includes("BRAKE CLEAN") || dUpper.includes("RUST INHIBITOR") || dUpper.includes("CYCLONE STEAM SEPARATOR")) {
    cat = "Misc";
    if      (dUpper.includes("SAFETY GATE"))      sub = "Safety Gate";
    else if (dUpper.includes("BLANKET"))           sub = "Blanket";
    else if (dUpper.includes("FLUSH RING"))        sub = "Flush Ring";
    else if (dUpper.includes("SYPHON"))            sub = "Syphon";
    else if (dUpper.includes("CAMLOCK") || dUpper.includes("CAM-LOCK")) sub = "Camlock";
    else if (dUpper.includes("HOSE"))              sub = "Hose";
    else if (dUpper.includes("EXPANSION JOINT"))   sub = "Expansion Joint";
    else if (dUpper.includes("INSULATION KIT"))    sub = "Insulation Kit";
    else if (dUpper.includes("CORROSION COUPON"))  sub = "Corrosion Coupon";
    else if (dUpper.includes("BRAKE CLEAN") || dUpper.includes("RUST INHIBITOR")) sub = "Consumables";
    else if (dUpper.includes("CYCLONE STEAM SEPARATOR")) sub = "Separator";
  }
  else if (tag) {
    if (tag.includes("-PT-") || tag.includes("-PIT-") || tag.includes("-PG-") || tag.includes("-TG-") || tag.includes("-FG-") ||
        tag.includes("-TE-") || tag.includes("-TW-") || tag.includes("-FT-") || tag.includes("-FIT-") || tag.includes("-FE-") ||
        tag.includes("-TIT-") || tag.includes("-TT-") || tag.match(/^(PT|PIT|FIT|FT|FE|TG|PG|TE|TW|PSE|PSHH|H2S|LEL)/)) {
       cat = "Instrument";
       const subInst = getInstSub();
       if (subInst !== "Misc") sub = subInst;
    }
    else if (tag.includes("-EHT-") || tag.includes("HTP") || tag.startsWith("ET-") || tag.startsWith("HS-") || tag.startsWith("GP-") ||
             tag.startsWith("PE-") || tag.startsWith("IOP-") || tag.startsWith("BCN") || tag.startsWith("HRN") || tag.startsWith("XFM") ||
             tag.startsWith("PC-") || tag.startsWith("REC-") || tag.startsWith("DP-")) {
       cat = "Electrical";
       const subElec = getElecSub();
       if (subElec !== "Misc") sub = subElec;
    }
  }




  return { category: cat, subcat: sub, size: size };
}
