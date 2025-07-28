/************** PROCESS ARCHIVE & DELETE (Trello Card Manager) **************/

function processTrelloActions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CARD_MANAGER);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${SHEET_CARD_MANAGER}" not found.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No data to process.");
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // A–D
  const statusOutput = [];

  const { key, token } = getTrelloCredentials();
  if (!key || !token) {
    SpreadsheetApp.getUi().alert("Trello API credentials are missing.");
    return;
  }

  for (let i = 0; i < data.length; i++) {
    const [title, shortUrl, actionRaw, labelName] = data[i];
    const action = String(actionRaw || "").toLowerCase().trim();

    if (!shortUrl || ![ACTION_DELETE, ACTION_ARCHIVE].includes(action)) {
      statusOutput.push([""]);
      continue;
    }

    const cardId = extractCardIdFromUrl(shortUrl);
    if (!cardId) {
      statusOutput.push(["❌ Invalid URL"]);
      continue;
    }

    try {
      let result = "";

      if (action === ACTION_DELETE) {
        const url = `https://api.trello.com/1/cards/${cardId}?key=${key}&token=${token}`;
        const res = UrlFetchApp.fetch(url, { method: "delete" });
        result = res.getResponseCode() === 200 ? "✅ Deleted" : "❌ Delete failed";

      } else if (action === ACTION_ARCHIVE) {
        const url = `https://api.trello.com/1/cards/${cardId}?closed=true&key=${key}&token=${token}`;
        const res = UrlFetchApp.fetch(url, { method: "put" });
        result = res.getResponseCode() === 200 ? "✅ Archived" : "❌ Archive failed";
      }

      statusOutput.push([result]);

    } catch (err) {
      statusOutput.push([`❌ ${err.message}`]);
    }
  }

  sheet.getRange(2, COL_STATUS, statusOutput.length).setValues(statusOutput);
}

/************** PROCESS LABEL ACTIONS (Trello Label Manager) **************/

function processLabelActions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LABEL_MANAGER);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${SHEET_LABEL_MANAGER}" not found.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No label actions to process.");
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // A–D
  const statusOutput = [];

  const { key, token } = getTrelloCredentials();
  if (!key || !token) {
    SpreadsheetApp.getUi().alert("Trello API credentials are missing.");
    return;
  }

  for (let i = 0; i < data.length; i++) {
    const [title, shortUrl, actionRaw, labelName] = data[i];
    const action = String(actionRaw || "").toLowerCase().trim();

    if (!shortUrl || ![ACTION_ADD_LABEL, ACTION_REMOVE_LABEL, ACTION_REMOVE_ALL_LABELS].includes(action)) {
      statusOutput.push([""]);
      continue;
    }

    const cardId = extractCardIdFromUrl(shortUrl);
    if (!cardId) {
      statusOutput.push(["❌ Invalid URL"]);
      continue;
    }

    try {
      let result = "";

      if (action === ACTION_ADD_LABEL) {
        const boardId = getBoardIdForCard(cardId, key, token);
        const labelId = findLabelIdByName(boardId, labelName, key, token);
        if (!labelId) throw new Error("Label not found on board");
        addLabelToCard(cardId, labelId, key, token);
        result = "✅ Label added";

      } else if (action === ACTION_REMOVE_LABEL) {
        const labelId = getLabelIdByName(cardId, labelName, key, token);
        if (!labelId) throw new Error("Label not found on card");
        removeLabelFromCard(cardId, labelId, key, token);
        result = "✅ Label removed";

      } else if (action === ACTION_REMOVE_ALL_LABELS) {
        removeAllLabelsFromCard(cardId, key, token);
        result = "✅ All labels removed";
      }

      statusOutput.push([result]);

    } catch (err) {
      statusOutput.push([`❌ ${err.message}`]);
    }
  }

  sheet.getRange(2, COL_LABEL_STATUS, statusOutput.length).setValues(statusOutput);
}

/************** CLEAR COLUMNS A–E (Trello Card Manager) **************/

function clearActionsAndStatuses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // ✅ ACTIVE SHEET
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("Nothing to clear.");
    return;
  }

  const numRows = lastRow - 1;
  const range = sheet.getRange(2, 1, numRows, 5);  // Columns A–E
  range.clearContent();

  SpreadsheetApp.getUi().alert(`✅ Cleared ${numRows} rows from "${sheet.getName()}"`);
}

/************** MENU **************/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Trello Tools")
    .addSubMenu(
      ui.createMenu("📋 Card Manager")
        .addItem("Run Card Actions", "processTrelloActions")
    )
    .addSubMenu(
      ui.createMenu("🏷️ Label Manager")
        .addItem("Run Label Actions", "processLabelActions")
        .addItem("Sync Label Dropdowns", "updateLabelManagerDropdowns")
        .addItem("♻️ Refresh Label Editor Sheet", "syncBoardLabels_LabelEditor")
    )
    .addSubMenu(
      ui.createMenu("📎 Attachment Manager")
        .addItem("Sync Attachments", "syncAttachmentsFromCardUrls")
        .addItem("Download Selected Attachments", "downloadSelectedAttachments")
    )
    .addSubMenu(
      ui.createMenu("🧹 Label Editor")
        .addItem("📃 List All Board Labels", "listAllBoardLabelsForCleanup")
        .addItem("❌ Delete Checked Labels", "deleteSelectedBoardLabelsFromSheet")
    )
    .addSubMenu(
      ui.createMenu("🔁 Trello Sync")
        .addItem("Refresh Boards", "fetchTrelloBoards")
        .addItem("Refresh Lists", "fetchTrelloListsForAllBoards")
        .addItem("Update List Dropdowns", "onBoardSelectionUpdate")
    )
    .addSubMenu(
      ui.createMenu("🚚 Card Mover")
        .addItem("Move Cards to Selected List", "moveCardsToSelectedList")
    )
    .addToUi(); // ✅ This applies the menu to the UI
}







