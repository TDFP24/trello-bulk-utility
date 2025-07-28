/************** UTILITY FUNCTIONS **************/

/**
 * Extract Trello card ID from its short URL
 */
function extractCardIdFromUrl(url) {
  if (!url || typeof url !== "string") return null;
  const match = url.match(/\/c\/([a-zA-Z0-9]+)/);
  return match ? match[1] : null;
}

/**
 * Apply the Action dropdown in Column C of Trello Card Manager
 */
function setCardManagerDropdown() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CARD_MANAGER);
  if (!sheet) return;

  const maxRows = 100;
  const actions = [ACTION_ARCHIVE, ACTION_DELETE];
  const range = sheet.getRange(2, COL_ACTION, maxRows);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(actions, true)
    .setAllowInvalid(false)
    .build();

  range.setDataValidation(rule);
  SpreadsheetApp.getUi().alert("✅ 'Trello Card Manager' dropdown updated.");
}

/**
 * Apply the Label Action dropdown in Column C of Trello Label Manager
 */
function setLabelManagerDropdown() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LABEL_MANAGER);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("❌ Sheet 'Trello Label Manager' not found.");
    return;
  }

  const maxRows = 100;
  const labelActions = [ACTION_ADD_LABEL, ACTION_REMOVE_LABEL, ACTION_REMOVE_ALL_LABELS];
  const range = sheet.getRange(2, COL_LABEL_ACTION, maxRows);  // Column C

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(labelActions, true)
    .setAllowInvalid(false)
    .build();

  range.setDataValidation(rule);
  SpreadsheetApp.getUi().alert("✅ 'Trello Label Manager' dropdown updated.");
}

/**
 * Sync contextual label dropdowns in Column D of Trello Label Manager
 */
function updateLabelManagerDropdowns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LABEL_MANAGER);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`❌ Sheet '${SHEET_LABEL_MANAGER}' not found.`);
    return;
  }

  const data = sheet.getDataRange().getValues();
  const { key, token } = getTrelloCredentials();

  if (!key || !token) {
    SpreadsheetApp.getUi().alert("❌ Trello API credentials are missing.");
    return;
  }

  for (let i = 1; i < data.length; i++) {
    const shortUrl = data[i][COL_SHORT_URL - 1];
    const action = (data[i][COL_LABEL_ACTION - 1] || "").toLowerCase().trim();
    const labelCell = sheet.getRange(i + 1, COL_LABEL_NAME);  // D
    const statusCell = sheet.getRange(i + 1, COL_LABEL_STATUS);  // E

    const cardId = extractCardIdFromUrl(shortUrl);
    if (!cardId) {
      labelCell.clearDataValidations();
      labelCell.setNote("❌ Invalid URL");
      continue;
    }

    try {
      let labelNames = [];

      if (action === ACTION_ADD_LABEL) {
        const boardId = getBoardIdForCard(cardId, key, token);
        labelNames = getLabelsForBoard(boardId, key, token).map(l => l.name).filter(Boolean);
      } else if (action === ACTION_REMOVE_LABEL) {
        labelNames = getLabelsForCard(cardId, key, token).map(l => l.name).filter(Boolean);
      } else {
        labelCell.clearDataValidations();
        labelCell.setNote("⛔ No label needed");
        continue;
      }

      if (labelNames.length > 0) {
        const rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(labelNames, true)
          .setAllowInvalid(false)
          .build();
        labelCell.setDataValidation(rule);
        labelCell.setNote(`Available: ${labelNames.join(", ")}`);
      } else {
        labelCell.clearDataValidations();
        labelCell.setNote("❌ No labels found");
      }

      statusCell.setValue("✅ Labels synced");

    } catch (err) {
      labelCell.clearDataValidations();
      labelCell.setNote("❌ Error");
      statusCell.setValue(`❌ ${err.message}`);
    }
  }

  SpreadsheetApp.getUi().alert("✅ Label dropdowns synced for Trello Label Manager.");
}

/**
 * Optional: Log Trello API credentials
 */
function checkTrelloCredentials() {
  const props = PropertiesService.getScriptProperties();
  Logger.log("Key: " + props.getProperty("TRELLO_API_KEY"));
  Logger.log("Token: " + props.getProperty("TRELLO_TOKEN"));
}
