/************** TRELLO LABEL UTILITIES **************/

function getBoardIdForCard(cardId, key, token) {
  const url = `https://api.trello.com/1/cards/${cardId}?fields=idBoard&key=${key}&token=${token}`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());
  return data.idBoard;
}

function getLabelsForBoard(boardId, key, token) {
  const limit = 1000;
  const url = `https://api.trello.com/1/boards/${boardId}/labels?fields=name,color&limit=${limit}&key=${key}&token=${token}`;
  const response = UrlFetchApp.fetch(url);
  return JSON.parse(response.getContentText());
}

function getAllLabelsForBoard(boardId, key, token) {
  const allLabels = [];
  let limit = 50;
  let before = null;
  let hasMore = true;

  while (hasMore) {
    let url = `https://api.trello.com/1/boards/${boardId}/labels?fields=name,color&limit=${limit}&key=${key}&token=${token}`;
    if (before) url += `&before=${before}`;

    const res = UrlFetchApp.fetch(url);
    const labels = JSON.parse(res.getContentText());
    allLabels.push(...labels);

    if (labels.length < limit) {
      hasMore = false;
    } else {
      before = labels[labels.length - 1].id;
    }
  }

  return allLabels;
}

function getLabelsForCard(cardId, key, token) {
  const url = `https://api.trello.com/1/cards/${cardId}/labels?fields=name,color&key=${key}&token=${token}`;
  const response = UrlFetchApp.fetch(url);
  return JSON.parse(response.getContentText());
}

function addLabelToCard(cardId, labelId, key, token) {
  const url = `https://api.trello.com/1/cards/${cardId}/idLabels?value=${labelId}&key=${key}&token=${token}`;
  UrlFetchApp.fetch(url, { method: "post" });
}

function removeLabelFromCard(cardId, labelId, key, token) {
  const url = `https://api.trello.com/1/cards/${cardId}/idLabels/${labelId}?key=${key}&token=${token}`;
  UrlFetchApp.fetch(url, { method: "delete" });
}

function removeAllLabelsFromCard(cardId, key, token) {
  const labels = getLabelsForCard(cardId, key, token);
  for (let label of labels) {
    removeLabelFromCard(cardId, label.id, key, token);
  }
}

function findLabelIdByName(boardId, labelName, key, token) {
  const labels = getLabelsForBoard(boardId, key, token);
  const match = labels.find(l => l.name.toLowerCase() === labelName.toLowerCase());
  return match ? match.id : null;
}

function getLabelIdByName(cardId, labelName, key, token) {
  const labels = getLabelsForCard(cardId, key, token);
  const match = labels.find(l => l.name.toLowerCase() === labelName.toLowerCase());
  return match ? match.id : null;
}

/************** DROPDOWN REFRESH (Generalized) **************/

function syncBoardLabelsToSheet(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert(`❌ Sheet '${sheetName}' not found.`);
    return;
  }

  const response = ui.prompt("Paste Trello Board ID", "Enter the Board ID (e.g., 67a2631c799b998134b3e6cd):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("❌ Cancelled. No labels were synced.");
    return;
  }

  const boardId = response.getResponseText().trim();
  const { key, token } = getTrelloCredentials();
  Logger.log("📋 Board ID entered: " + boardId);

  let labels;
  try {
    labels = getAllLabelsForBoard(boardId, key, token);
    Logger.log("🎯 Raw label response: " + JSON.stringify(labels));
  } catch (err) {
    ui.alert("❌ Failed to fetch labels: " + err.message);
    return;
  }

  const labelNames = labels.map(l => l.name).filter(Boolean);
  Logger.log("✅ Extracted label names: " + labelNames.join(", "));

  if (!labelNames.length) {
    ui.alert("⚠️ No labels found on this board.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  let updatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const actionCell = data[i][COL_LABEL_ACTION - 1];
    const action = String(actionCell || "").toLowerCase().trim();
    if (action !== ACTION_ADD_LABEL) continue;

    const labelCell = sheet.getRange(i + 1, COL_LABEL_NAME);
    const statusCell = sheet.getRange(i + 1, COL_LABEL_STATUS);

    try {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(labelNames, true)
        .setAllowInvalid(false)
        .build();

      labelCell.clearDataValidations();
      labelCell.setDataValidation(rule);
      labelCell.setNote(`Labels: ${labelNames.join(", ")}`);
      statusCell.setValue("✅ Labels refreshed");
      updatedCount++;
    } catch (err) {
      statusCell.setValue("❌ Dropdown error");
      Logger.log(`❌ Row ${i + 1}: ` + err.message);
    }
  }

  Logger.log(`🔁 Total rows updated: ${updatedCount}`);
  ui.alert(`✅ Updated ${updatedCount} rows with latest board labels.`);
}

/************** WRAPPERS FOR MENU **************/

function syncBoardLabels_TrelloLabelManager() {
  syncBoardLabelsToSheet("Trello Label Manager");
}

function syncBoardLabels_LabelEditor() {
  syncBoardLabelsToSheet("Label Editor");
}
