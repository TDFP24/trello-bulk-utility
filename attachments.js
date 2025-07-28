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
    ["Card Title", "Trello Short URL", "✅ Download", "Attachment 1", "Attachment 2", "Attachment 3", "Attachment 4", "Attachment 5", "File Type", "Status"]
  ];
  sheet.getRange("A1:J1").setValues(headers);
  sheet.setFrozenRows(1);

  const maxRows = 100;
  sheet.getRange(2, 3, maxRows).insertCheckboxes(); // Column C

  // ✅ Static dropdown for file types (includes .zip, .png, .jpg)
  const fileTypes = [".zip", ".png", ".jpg", ".pdf", ".cdr", ".ai"];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(fileTypes, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 9, maxRows).setDataValidation(rule); // Column I

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

  // Clear old data from D–I and Status column
  sheet.getRange(2, 4, sheet.getLastRow() - 1, 6).clearContent();

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
        sheet.getRange(row, 10).setValue("ℹ️ No attachments");
        continue;
      }

      const names = attachments.map(a => a.name);
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(names, true).build();

      for (let j = 0; j < 5; j++) {
        const cell = sheet.getRange(row, 4 + j);
        cell.setValue(names[j] || "");
        cell.setDataValidation(rule);
      }

      sheet.getRange(row, 10).setValue(`✅ ${attachments.length} found`);

    } catch (e) {
      sheet.getRange(row, 10).setValue(`❌ ${e.message}`);
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
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues(); // A2:J

  for (let i = 0; i < data.length; i++) {
    const rowNum = i + 2;
    const [cardTitle, shortUrl, toDownload, att1, att2, att3, att4, att5, fileTypeRaw] = data[i];
    const statusCell = sheet.getRange(rowNum, 10); // Column J

    statusCell.setValue(""); // clear previous status

    if (toDownload !== true) continue;
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

      // Normalize inputs
      const selectedFileType = (fileTypeRaw || "").toLowerCase().trim(); // e.g., ".cdr"
      const selectedNames = [att1, att2, att3, att4, att5].filter(Boolean).map(name => name.trim());

      let targets = [];

      if (selectedFileType) {
        // 🔍 Match ALL Trello attachments with the given extension (Column I trumps dropdowns)
        targets = attachments.filter(att => {
          const ext = att.name.slice(att.name.lastIndexOf(".")).toLowerCase();
          return ext === selectedFileType;
        });
      } else {
        // 🎯 Match only those listed in D–H
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
