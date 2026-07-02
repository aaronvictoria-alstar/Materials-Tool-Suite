// WMS DASHBOARD V1.02

// ==========================================
// WEB APP GATEWAY
// ==========================================
// Serves the front-end WMS_Dashboard HTML and allows server-side file includes.
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
// 1. NEW DELIVERY PROCESSOR (Forms & Drive)
// ==========================================
// Intercepts form data and base64 image payloads from the web app, saves the file to Drive, and appends the metadata to RecLog2.
function initializeDriveFolder() {
  const folderName = "Vendor BOLs";
  const folders = DriveApp.getFoldersByName(folderName);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  return folder.getId();
}

function processNewDelivery(formData, attachmentPayload) {
  try {
    let fileUrl = "No Attachment";




    // 1. Process Attachment if it exists
    if (attachmentPayload && attachmentPayload.base64Data) {
      const folderId = initializeDriveFolder();
      const folder = DriveApp.getFolderById(folderId);
     
      // Strip the metadata header from the Base64 string (e.g., "data:image/jpeg;base64,")
      const base64String = attachmentPayload.base64Data.split(',')[1];
      const decoded = Utilities.base64Decode(base64String);
      const blob = Utilities.newBlob(decoded, attachmentPayload.mimeType, attachmentPayload.filename);
     
      const file = folder.createFile(blob);
      fileUrl = file.getUrl();
    }




    // 2. Append to RecLog2 in the Master Log
    const ss = SpreadsheetApp.openById(MASTER_LOG_ID);
    let recLogSheet = ss.getSheetByName("RecLog2") || ss.getSheetByName("RecLog");
   
    if (!recLogSheet) throw new Error("Could not find 'RecLog2' tab in the Master Log.");




    // Structure the row array. You can adjust this order to perfectly match your RecLog2 columns!
    const rowToAppend = [
      formData.timestamp,      // A: Date/Time
      formData.unloader,       // B: Receiver
      formData.job,            // C: Job #
      formData.po,             // D: PO #
      formData.vendor,         // E: Vendor
      formData.pl,             // F: Packing List
      formData.type,           // G: Load Type
      formData.isFreeIssue ? "YES" : "NO", // H: Free Issue
      formData.clientPo,       // I: Client PO
      fileUrl,                 // J: BOL Link
      "OPEN"                   // K: Status (Placeholder for 'not submitted')
    ];




    recLogSheet.appendRow(rowToAppend);
    return { status: "Success", message: "Delivery logged and saved." };




  } catch (error) {
    throw new Error("Backend Error: " + error.message);
  }
}

// ==========================================
// 2. DASHBOARD DATA FETCHER
// ==========================================
// Queries the Master Log tabs (RecLog2, Expected Receiving, OS&D) and returns a JSON payload to populate the Web App tables.
function getReceivingDashboardData() {
  const ss = SpreadsheetApp.openById(MASTER_LOG_ID);
 
  const payload = {
    expected: [],
    open: [],
    osd: []
  };




  // Helper: Converts a sheet to an array of objects based on header names
  function getSheetDataAsObjects(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];
   
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase().trim());
    const rows = [];
   
    for (let i = 1; i < data.length; i++) {
      let obj = {};
      for (let j = 0; j < headers.length; j++) {
        obj[headers[j]] = data[i][j];
      }
      rows.push(obj);
    }
    return { rows: rows, headers: headers };
  }




  try {
    // --- 1. FETCH OPEN RECEIVING (RecLog2) ---
    // We look for rows that don't have a "submitted" date/flag.
    const recData = getSheetDataAsObjects("RecLog2");
    if (recData.rows) {
      // Reversing so newest are at the top. Grabbing top 50 Open rows for speed.
      const reversed = recData.rows.reverse();
      for (const r of reversed) {
        // Assume 'status' or 'submitted' determines if it's open.
        // If your logic relies on a blank cell in a specific column, adjust the key here!
        payload.open.push(r);
        if (payload.open.length >= 50) break;
      }
    }




    // --- 2. FETCH EXPECTED RECEIVING ---
    const expData = getSheetDataAsObjects("Expected Receiving");
    if (expData.rows) {
      const today = new Date();
      today.setHours(0,0,0,0);
     
      const dayOfWeek = today.getDay(); // 0 = Sun, 4 = Thu, 5 = Fri
      let cutoffDate = new Date(today);
     
      // Logic: If Thu/Fri, lookahead to next Monday. Else, lookahead to this Friday.
      if (dayOfWeek === 4) cutoffDate.setDate(cutoffDate.getDate() + 4); // Thu to Mon
      else if (dayOfWeek === 5) cutoffDate.setDate(cutoffDate.getDate() + 3); // Fri to Mon
      else cutoffDate.setDate(cutoffDate.getDate() + (5 - dayOfWeek)); // Mon/Tue/Wed to Fri
     
      for (const r of expData.rows) {
        // Look for common header names for the expected date
        let dateVal = r["date"] || r["expected date"] || r["eta"] || r["arrival"];
        if (dateVal instanceof Date) {
          let expectedDate = new Date(dateVal);
          expectedDate.setHours(0,0,0,0);
         
          if (expectedDate >= today && expectedDate <= cutoffDate) {
            payload.expected.push(r);
          }
        }
      }
    }




    // --- 3. FETCH OS&D / ERROR LOG ---
    const osdData = getSheetDataAsObjects("ErrorLog") || getSheetDataAsObjects("OS&D");
    if (osdData && osdData.rows) {
      payload.osd = osdData.rows.reverse().slice(0, 50); // Top 50 recent errors
    }




  } catch (err) {
    Logger.log("Error fetching dashboard data: " + err.message);
  }




  return payload;
}