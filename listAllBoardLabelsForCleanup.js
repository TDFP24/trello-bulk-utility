function listAllBoardLabelsForCleanup() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Label Editor");
  if (!sheet) {
    ui.alert("❌ Sheet 'Label Editor' not found.");
    return;
  }

  const response = ui.prompt("Enter the Trello Board ID", "Paste the Board ID:", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const boardId = response.getResponseText().trim();
  const { key, token } = getTrelloCredentials();

  let labels;
  try {
    labels = getAllLabelsForBoard(boardId, key, token);
  } catch (err) {
    ui.alert("❌ Failed to fetch labels: " + err.message);
    return;
  }

  sheet.clearContents();
  sheet.getRange("A1:C1").setValues([["Label Name", "Label ID", "❌ Delete?"]]);

  const rows = labels.map(l => [l.name, l.id, false]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    const checkCol = sheet.getRange(2, 3, rows.length);
    checkCol.insertCheckboxes();
    ui.alert(`✅ ${rows.length} labels listed. Tick 'Delete?' column for any you wish to remove, then run the delete script.`);
  } else {
    ui.alert("⚠️ No labels found.");
  }
}
function deleteSelectedBoardLabelsFromSheet() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Label Editor");
  if (!sheet) {
    ui.alert("❌ Sheet 'Label Editor' not found.");
    return;
  }

  const { key, token } = getTrelloCredentials();
  const data = sheet.getDataRange().getValues();

  let deleted = 0;
  for (let i = 1; i < data.length; i++) {
    const [name, labelId, shouldDelete] = data[i];
    if (shouldDelete === true && labelId) {
      try {
        const success = deleteLabelById(labelId, key, token);
        if (success) deleted++;
        sheet.getRange(i + 1, 4).setValue("✅ Deleted");
      } catch (err) {
        Logger.log(`Failed to delete ${name}: ${err.message}`);
        sheet.getRange(i + 1, 4).setValue("❌ Failed");
      }
    }
  }

  ui.alert(`🗑️ Deleted ${deleted} labels from the board.`);
}
function deleteLabelById(labelId, key, token) {
  const url = `https://api.trello.com/1/labels/${labelId}?key=${key}&token=${token}`;
  const res = UrlFetchApp.fetch(url, { method: "delete" });
  return res.getResponseCode() === 200;
}
