/*
 * Google Apps Script — Cuppilo Campaign Data Receiver
 *
 * Deploy this as a Web App (Execute as: Me, Who has access: Anyone)
 * Then paste the web app URL into YOUR_SCRIPT_URL in index.html
 *
 * Expected JSON body:
 * {
 *   "timestamp": "...",
 *   "name": "...",
 *   "frequency": "...",
 *   "cafeType": "...",
 *   "atmosphere": "...",
 *   "drinks": "...",
 *   "food": "...",
 *   "priceRange": "...",
 *   "ideas": "...",
 *   "contact": "..."
 * }
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // If the sheet is empty, write the header row first
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp',
        'Name',
        'Cafe Frequency',
        'Cafe Type',
        'Atmosphere',
        'Drinks',
        'Food',
        'Price Range',
        'Ideas',
        'Contact'
      ]);
    }

    sheet.appendRow([
      data.timestamp   || new Date().toISOString(),
      data.name        || '',
      data.frequency   || '',
      data.cafeType    || '',
      data.atmosphere  || '',
      data.drinks      || '',
      data.food        || '',
      data.priceRange  || '',
      data.ideas       || '',
      data.contact     || ''
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
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Cuppilo Campaign API is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}
