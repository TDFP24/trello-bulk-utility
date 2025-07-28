function setupAttachmentManagerSheet() {
  const sheetName = "Trello Attachment Manager";
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  }

  // Set header row
  const headers = ["✅", "Card Title", "Attachment Name", "Attachment URL", "Status"];
  sheet.getRange("A1:E1").setValues([headers]);

  // Label F1
  sheet.getRange("F1").setValue("Paste Trello List ID here →");

  // Apply checkboxes in Column A
  const maxRows = 100; // adjust as needed
  const checkboxRange = sheet.getRange(2, 1, maxRows);
  checkboxRange.insertCheckboxes();

  SpreadsheetApp.getUi().alert("✅ 'Trello Attachment Manager' initialized. Paste List ID in F1 and run Sync.");
}
