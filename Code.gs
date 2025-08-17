function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Gavel Strike")
    .addItem("Open calling order", "showSidebar")
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createTemplateFromFile("call-order")
    .evaluate()
    .setTitle("Calling Order")
  SpreadsheetApp.getUi().showSidebar(html);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Gets the ordered list of names for the selected range and mode
 */
function getCallingOrder(rangeA1, sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];

  const range = sheet.getRange(rangeA1);
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  const fontLines = range.getFontLines();

  const startRow = range.getRow();
  const startCol = range.getColumn();

  const numRows = values.length;
  const numCols = values[0].length;

  let order = [];

  // Precompute column letters to avoid calling getRange().getA1Notation()
  const colLetters = [];
  for (let c = 0; c < numCols; c++) {
    colLetters.push(columnToLetter(startCol + c));
  }

  for (let col = 0; col < numCols; col++) {
    if (!values[0][col]) continue;

    for (let row = 0; row < numRows; row++) {
      const value = values[row][col];
      if (!value) continue;

      const bg = backgrounds[row][col];
      const isBlack = bg && bg.toLowerCase() === "#000000";
      const isStrikethrough = fontLines[row][col] === "line-through";

      if (isBlack || isStrikethrough) continue;

      // Build A1 notation manually: column letter + row number
      const cellRef = colLetters[col] + (startRow + row);

      order.push({
        text: value,
        cell: cellRef,
        sheetName,
        row: row + 1, // Spreadsheet start from one, not zero
        col: col + 1,
        rangeA1
      });
    }
  }

  return order;
}

/**
 * Convert column number to letter(s) (1 = A, 27 = AA)
 */
function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
/**
 * "Click" action for a list item
 */
function listItemClick(metadata) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(metadata.sheetName);
  if (!sheet) return false;

  const range = sheet.getRange(metadata.rangeA1);
  const startRow = range.getRow();
  const startCol = range.getColumn();
  const endRow = range.getLastRow();
  const endCol = range.getLastColumn();

  // Absolute coordinates of original cell. Subtract by one because of different counting systems :)
  const absRow = startRow + metadata.row - 1;
  const absCol = startCol + metadata.col - 1;

  if (metadata.blackout) sheet.getRange(absRow, absCol).setBackground("#000000");
  else sheet.getRange(absRow, absCol).setFontLine('line-through');

  // Determine next column
  const nextCol = absCol + 1;
  if (nextCol > endCol) return false;

  // Read entire next column in one go
  const colValues = sheet.getRange(startRow, nextCol, endRow - startRow + 1, 1).getValues();

  // Find topmost empty cell
  let targetIndex = colValues.findIndex(r => r[0] === "");
  if (targetIndex === -1) return false; // no empty cell

  // Write value in one operation
  sheet.getRange(startRow + targetIndex, nextCol).setValue(
    sheet.getRange(absRow, absCol).getValue()
  );

  return true;
}

function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}
