// === 1. One-time setup to create the "Trello Card Mover" sheet ===
function setupTrelloCardMoverSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Trello Card Mover");
  if (!sheet) sheet = ss.insertSheet("Trello Card Mover");
  else sheet.clearContents();

  const headers = [
    "Card Title",       // A
    "Card Short URL",   // B
    "Target Board",     // C
    "Target List",      // D
    "Position",         // E
    "Move Status"       // F
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Pre-fill position dropdown
  const positionRange = sheet.getRange("E2:E100");
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["top", "bottom"], true)
    .build();
  positionRange.setDataValidation(rule);
}

// === 2. Create/clear config sheet to store boards and lists ===
function setupTrelloConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let config = ss.getSheetByName("Trello Config");
  if (!config) config = ss.insertSheet("Trello Config");
  config.clearContents();
  config.hideSheet();

  config.getRange("A1:B1").setValues([["Board Name", "Board ID"]]);
}

// === 3. Fetch Trello boards and populate column C dropdown ===
function fetchTrelloBoards() {
  const { key, token } = getTrelloCredentials();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName("Trello Config");
  const moverSheet = ss.getSheetByName("Trello Card Mover");

  if (!config || !moverSheet) throw new Error("Missing required sheets.");

  const url = `https://api.trello.com/1/members/me/boards?key=${key}&token=${token}&fields=name,id&filter=open`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());

  const boardMap = data.map(board => [board.name, board.id]);
  config.getRange(2, 1, boardMap.length, 2).setValues(boardMap);

  // ✅ Safe trimming
  const numExtraRows = config.getMaxRows() - (boardMap.length + 1);
  if (numExtraRows > 0) {
    config.deleteRows(boardMap.length + 2, numExtraRows);
  }

  // Create named range
  const boardNamesRange = config.getRange(2, 1, boardMap.length, 1);
  ss.setNamedRange("BoardNames", boardNamesRange);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getRangeByName("BoardNames"), true)
    .setAllowInvalid(false)
    .build();

  moverSheet.getRange("C2:C100").setDataValidation(rule);
}
function fetchTrelloListsForAllBoards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Trello Config");
  const listSheet = ss.getSheetByName("Trello Lists") || ss.insertSheet("Trello Lists");

  // Clear and setup header
  listSheet.clearContents();
  listSheet.getRange("A1:C1").setValues([["Board Name", "List Name", "List ID"]]);

  const { key, token } = getTrelloCredentials();

  const boardData = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
  let row = 2;

  for (let [boardName, boardId] of boardData) {
    if (!boardId) continue;
    try {
      const url = `https://api.trello.com/1/boards/${boardId}/lists?fields=name&key=${key}&token=${token}`;
      const response = UrlFetchApp.fetch(url);
      const lists = JSON.parse(response.getContentText());

      const listRows = lists.map(list => [boardName, list.name, list.id]);
      if (listRows.length) {
        listSheet.getRange(row, 1, listRows.length, 3).setValues(listRows);
        row += listRows.length;
      }
    } catch (err) {
      Logger.log(`Error fetching lists for ${boardName}: ${err.message}`);
    }
  }

  SpreadsheetApp.getUi().alert("✅ Trello Lists sheet updated.");
}
function setupTrelloMoverAndConfigSheets() {
  setupTrelloCardMoverSheet();
  setupTrelloConfigSheet();
  SpreadsheetApp.getUi().alert("✅ Mover and Config sheets are ready.");
}
function onBoardSelectionUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const moverSheet = ss.getSheetByName("Trello Card Mover");
  const configSheet = ss.getSheetByName("Trello Config");
  const listsSheet = ss.getSheetByName("Trello Lists");

  if (!moverSheet || !configSheet || !listsSheet) {
    SpreadsheetApp.getUi().alert("❌ One or more required sheets are missing.");
    return;
  }

  const configData = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues(); // [Board Name, Board ID]
  const listsData = listsSheet.getRange(2, 1, listsSheet.getLastRow() - 1, 2).getValues();     // [Board Name, List Name]

  const boardIdMap = Object.fromEntries(configData); // boardName → boardId
  const listMap = {}; // boardName → [list1, list2, ...]

  for (const [bName, lName] of listsData) {
    if (!listMap[bName]) listMap[bName] = [];
    listMap[bName].push(lName);
  }

  const numRows = moverSheet.getLastRow() - 1;
  const boardValues = moverSheet.getRange(2, 3, numRows).getValues(); // Column C

  for (let i = 0; i < numRows; i++) {
    const boardName = boardValues[i][0];
    const cell = moverSheet.getRange(i + 2, 4); // Column D

    if (!boardName || !listMap[boardName]) {
      cell.clearDataValidations();
      cell.setNote("⚠️ No lists found");
      continue;
    }

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(listMap[boardName], true)
      .setAllowInvalid(false)
      .build();

    cell.setDataValidation(rule);
    cell.setNote(`✅ ${listMap[boardName].length} lists found`);
  }

  SpreadsheetApp.getUi().alert("✅ List dropdowns updated for 'Trello Card Mover'.");
}
function moveCardsToSelectedList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Trello Card Mover");
  const configSheet = ss.getSheetByName("Trello Config");
  const listSheet = ss.getSheetByName("Trello Lists");

  if (!sheet || !configSheet || !listSheet) {
    SpreadsheetApp.getUi().alert("❌ One or more required sheets are missing.");
    return;
  }

  const { key, token } = getTrelloCredentials();

  const boardData = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues(); // [Board Name, Board ID]
  const boardMap = Object.fromEntries(boardData); // Board Name → ID

  const listData = listSheet.getRange(2, 1, listSheet.getLastRow() - 1, 3).getValues(); // [Board Name, List Name, List ID]
  // Build a map: listId -> { boardName, boardId }
  const listIdToBoard = {};
  for (let [bName, lName, lId] of listData) {
    listIdToBoard[lId] = { boardName: bName, boardId: boardMap[bName] };
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues(); // A2:F

  for (let i = 0; i < rows.length; i++) {
    const [cardTitle, shortUrl, boardName, listName, position, status] = rows[i];
    const statusCell = sheet.getRange(i + 2, 6); // Column F

    if (!shortUrl || !boardName || !listName || !position) {
      statusCell.setValue("⚠️ Missing input");
      continue;
    }

    const cardId = extractCardIdFromUrl(shortUrl);
    if (!cardId) {
      statusCell.setValue("❌ Invalid card URL");
      continue;
    }

    // Find the listId and destination boardId
    let listId = null;
    let destBoardId = null;
    for (let [bName, lName, lId] of listData) {
      if (bName === boardName && lName === listName) {
        listId = lId;
        destBoardId = boardMap[bName];
        break;
      }
    }
    if (!listId || !destBoardId) {
      statusCell.setValue("❌ List or Board not found");
      continue;
    }

    // Ensure the undo log sheet exists
    let undoSheet = ss.getSheetByName("Trello Undo Log");
    if (!undoSheet) {
      undoSheet = ss.insertSheet("Trello Undo Log");
      undoSheet.getRange(1, 1, 1, 4).setValues([["Card ID", "Board ID", "List ID", "Pos"]]);
    }

    // Fetch the card's current board, list, and position before moving
    try {
      const cardInfoUrl = `https://api.trello.com/1/cards/${cardId}?fields=idBoard,idList,pos&key=${key}&token=${token}`;
      const cardInfoRes = UrlFetchApp.fetch(cardInfoUrl);
      const cardInfo = JSON.parse(cardInfoRes.getContentText());
      undoSheet.appendRow([cardId, cardInfo.idBoard, cardInfo.idList, cardInfo.pos]);
    } catch (err) {
      // If logging fails, still attempt the move, but note the error in the status cell
      statusCell.setValue(`⚠️ Undo log failed: ${err.message}`);
    }

    try {
      const moveUrl = `https://api.trello.com/1/cards/${cardId}?idBoard=${destBoardId}&idList=${listId}&pos=${position}&key=${key}&token=${token}`;
      const res = UrlFetchApp.fetch(moveUrl, { method: "put" });
      const code = res.getResponseCode();
      if (code === 200) {
        statusCell.setValue("✅ Moved");
      } else {
        statusCell.setValue(`❌ Error ${code}`);
      }
    } catch (err) {
      statusCell.setValue(`❌ ${err.message}`);
    }
  }
}