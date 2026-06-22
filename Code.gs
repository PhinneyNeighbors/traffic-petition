// ============================================================
// STEP 1: Paste this entire file into Google Apps Script
// STEP 2: Change SHEET_ID below to your Google Sheet's ID
// STEP 3: Click Deploy > New Deployment > Web App
//         - Execute as: Me
//         - Who has access: Anyone
// STEP 4: Copy the deployment URL into index.html
// ============================================================

const SHEET_ID = "YOUR_GOOGLE_SHEET_ID_HERE"; // <-- replace this

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Signatures") 
                  || SpreadsheetApp.openById(SHEET_ID).getActiveSheet();

    // Add header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Full Name", "Address", "Email", "Signature (image)"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    }

    sheet.appendRow([
      new Date().toLocaleString(),
      data.name || "",
      data.address || "",
      data.email || "",
      data.signature ? "✓ signed" : "—"
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Signatures")
                || SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const count = Math.max(0, sheet.getLastRow() - 1); // subtract header
  return ContentService
    .createTextOutput(JSON.stringify({ count: count }))
    .setMimeType(ContentService.MimeType.JSON);
}
