/************** CONFIGURATION CONSTANTS **************/

const SHEET_CARD_MANAGER  = "Trello Card Manager";
const SHEET_LABEL_MANAGER = "Trello Label Manager";

// Trello Card Manager (Sheet 1)
const COL_TITLE     = 1;  // A
const COL_SHORT_URL = 2;  // B
const COL_ACTION    = 3;  // C
const COL_LABEL     = 4;  // D
const COL_STATUS    = 5;  // E

// Trello Label Manager (Sheet 2)
const COL_LABEL_ACTION = 3;  // C
const COL_LABEL_NAME   = 4;  // D
const COL_LABEL_STATUS = 5;  // E

const ACTION_DELETE            = "delete";
const ACTION_ARCHIVE           = "archive";
const ACTION_ADD_LABEL         = "add label";
const ACTION_REMOVE_LABEL      = "remove label";
const ACTION_REMOVE_ALL_LABELS = "remove all labels";

const ACTION_OPTIONS = [
  ACTION_ARCHIVE,
  ACTION_DELETE,
  ACTION_ADD_LABEL,
  ACTION_REMOVE_LABEL,
  ACTION_REMOVE_ALL_LABELS
];
