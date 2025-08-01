const SHEET_ATTACHMENT_MANAGER = "Trello Attachment Manager";
const DRIVE_FOLDER_ID = "1yVcoI4SeXp4dF96WRLo6w2HE_DKTLRjE";

/**
 * 📦 Initialize the sheet layout with headers, checkboxes and dropdowns
 */
function initializeNewAttachmentSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName(SHEET_ATTACHMENT_MANAGER);
  if (existing) ss.deleteSheet(existing);
  const sheet = ss.insertSheet(SHEET_ATTACHMENT_MANAGER);

  const headers = [
    ["Card Title", "Trello Short URL", "Attachment 1", "Attachment 2", "Attachment 3", "Attachment 4", "Attachment 5", "File Type 1", "File Type 2", "File Type 3", "Status"]
  ];
  sheet.getRange("A1:K1").setValues(headers);
  sheet.setFrozenRows(1);

  const maxRows = 100;

  // ✅ Static dropdown for file types (includes .zip, .png, .jpg)
  const fileTypes = [".zip", ".png", ".jpg", ".pdf", ".cdr", ".ai"];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(fileTypes, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 8, maxRows, 3).setDataValidation(rule); // Columns H, I, J

  SpreadsheetApp.getUi().alert("✅ Sheet has been initialized.");
}

/**
 * 🔄 Sync attachments from card short URLs
 */
function syncAttachmentsFromCardUrls() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTACHMENT_MANAGER);
  if (!sheet) return;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues(); // A2:B
  const { key, token } = getTrelloCredentials();

  // Clear old data from C–G and Status column (K)
  sheet.getRange(2, 3, sheet.getLastRow() - 1, 5).clearContent(); // C-G
  sheet.getRange(2, 11, sheet.getLastRow() - 1, 1).clearContent(); // K (Status)

  for (let i = 0; i < data.length; i++) {
    const [title, url] = data[i];
    if (!url) continue;

    const row = i + 2;
    try {
      const cardId = extractCardIdFromUrl(url);
      const apiUrl = `https://api.trello.com/1/cards/${cardId}/attachments?fields=name,url&key=${key}&token=${token}`;
      const response = UrlFetchApp.fetch(apiUrl);
      const attachments = JSON.parse(response.getContentText());

      if (!attachments.length) {
        sheet.getRange(row, 11).setValue("ℹ️ No attachments");
        continue;
      }

      const names = attachments.map(a => a.name);
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(names, true).build();

      for (let j = 0; j < 5; j++) {
        const cell = sheet.getRange(row, 3 + j);
        cell.setValue(names[j] || "");
        cell.setDataValidation(rule);
      }

      sheet.getRange(row, 11).setValue(`✅ ${attachments.length} found`);

    } catch (e) {
      sheet.getRange(row, 11).setValue(`❌ ${e.message}`);
    }
  }
}

/**
 * ⬇️ Download attachments based on checkboxes + selection
 */
function downloadSelectedAttachments() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTACHMENT_MANAGER);
  if (!sheet) return;

  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const { key, token } = getTrelloCredentials();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues(); // A2:K

  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 2;
    const [cardTitle, shortUrl, att1, att2, att3, att4, att5, fileType1, fileType2, fileType3] = data[i];
    const statusCell = sheet.getRange(rowNum, 11); // Column K

    statusCell.setValue(""); // clear previous status

    if (!shortUrl) {
      statusCell.setValue("❌ Missing URL");
      continue;
    }

    // Extract card ID
    let cardId;
    try {
      cardId = extractCardIdFromUrl(shortUrl);
    } catch (e) {
      statusCell.setValue("❌ Invalid card URL");
      continue;
    }

    try {
      const url = `https://api.trello.com/1/cards/${cardId}/attachments?key=${key}&token=${token}`;
      const response = UrlFetchApp.fetch(url);
      const attachments = JSON.parse(response.getContentText());

      if (!attachments || attachments.length === 0) {
        statusCell.setValue("ℹ️ No attachments found");
        continue;
      }

      // Collect all specified file types (H, I, J)
      const selectedFileTypes = [fileType1, fileType2, fileType3]
        .filter(Boolean)
        .map(type => type.toLowerCase().trim()); // e.g., [".cdr", ".zip", ".png"]
      
      const selectedNames = [att1, att2, att3, att4, att5].filter(Boolean).map(name => name.trim());

      let targets = [];

      if (selectedFileTypes.length > 0) {
        // 🔍 Match ALL Trello attachments with ANY of the specified extensions
        targets = attachments.filter(att => {
          const ext = att.name.slice(att.name.lastIndexOf(".")).toLowerCase();
          return selectedFileTypes.includes(ext);
        });
      } else {
        // 🎯 Match only those listed in C–G
        targets = attachments.filter(att => selectedNames.includes(att.name));
      }

      Logger.log(`🔎 Matching files (Row ${rowNum}): ${targets.map(f => f.name).join(", ")}`);

      if (targets.length === 0) {
        statusCell.setValue("❌ No matching attachments");
        continue;
      }

      let downloaded = 0;

      for (const file of targets) {
        const fileUrl = file.url;

        const fetchResponse = UrlFetchApp.fetch(fileUrl, {
          headers: {
            "Authorization": `OAuth oauth_token="${token}", oauth_consumer_key="${key}"`
          },
          muteHttpExceptions: true,
          followRedirects: true
        });

        if (fetchResponse.getResponseCode() !== 200) {
          Logger.log(`❌ Failed to fetch ${file.name}: ${fetchResponse.getResponseCode()}`);
          statusCell.setValue(`❌ ${file.name}: ${fetchResponse.getResponseCode()}`);
          continue;
        }

        const blob = fetchResponse.getBlob().setName(file.name);
        folder.createFile(blob);
        downloaded++;
      }

      statusCell.setValue(`✅ ${downloaded} file(s)`);

    } catch (err) {
      Logger.log(`❌ Error on row ${rowNum}: ${err.message}`);
      statusCell.setValue(`❌ ${err.message}`);
    }
  }

  SpreadsheetApp.getUi().alert("✅ Download process complete.");
}

/**
 * 🗑️ Clear attachment sheet data while preserving dropdowns
 */
function clearAttachmentSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ATTACHMENT_MANAGER);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ 'Trello Attachment Manager' sheet not found.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("ℹ️ Sheet is already empty.");
    return;
  }

  // Clear data while preserving dropdowns
  // Columns to clear: A (Card Title), B (URL), C-G (Attachments), K (Status)
  const numRows = lastRow - 1;
  
  // Clear Card Title and URL (A, B)
  sheet.getRange(2, 1, numRows, 2).clearContent();
  
  // Clear Attachment columns (C-G)
  sheet.getRange(2, 3, numRows, 5).clearContent();
  
  // Clear Status column (K)
  sheet.getRange(2, 11, numRows, 1).clearContent();

  SpreadsheetApp.getUi().alert(`✅ Cleared ${numRows} rows from 'Trello Attachment Manager'. Dropdowns preserved.`);
}