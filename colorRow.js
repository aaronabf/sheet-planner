function format() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var fixRow = 26;
  var fixCol = 8;
  var maxRow = sheet.getMaxRows();
  var maxCol = sheet.getMaxColumns();

  // Clears the sheet of all formatting
  sheet.clearFormats();

  // Ensures the sheet is of the correct number of rows and columns
  if (maxRow > fixRow)
    sheet.deleteRows(fixRow + 1, maxRow - fixRow);
  else if (maxRow < fixRow)
    sheet.insertRows(maxRow, fixRow - maxRow);

  if (maxCol > fixCol)
    sheet.deleteColumns(fixCol + 1, maxCol - fixCol);
  else if (maxCol < fixCol)
    sheet.insertColumns(maxCol, fixCol - maxCol);

  // Sets the first row and column frozen
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  // Sets heights and width for each row and column, respectively
  for (var r = 1; r < maxRow; r++)
    sheet.setRowHeight(r, 22);

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 310);
  sheet.setColumnWidth(3, 60);
  sheet.setColumnWidth(4, 35);
  sheet.setColumnWidth(5, 35);
  sheet.setColumnWidth(6, 75);
  sheet.setColumnWidth(7, 70);
  sheet.setColumnWidth(8, 280);

  // Sets top row to bold and sets correct values
  sheet.getRange(1, 1, 1, fixCol).setFontWeight("bold");
  sheet.getRange(1, 1).setValue("Course");
  sheet.getRange(1, 2).setValue("Assignment Name");
  sheet.getRange(1, 3).setValue("DL");
  sheet.getRange(1, 4).setValue("Prio");
  sheet.getRange(1, 5).setValue("We");
  sheet.getRange(1, 6).setValue("Date Due");
  sheet.getRange(1, 7).setValue("Time");
  sheet.getRange(1, 8).setValue("Notes");


  // Sets alignment for each columns (and format for 6th column)
  for (var c = 1; c <= fixCol; c++) {
    var range = sheet.getRange(1, c, fixRow, 1);
    //range.getValue() = 5;
    if (c === 1)
      range.setHorizontalAlignment("right");
    else if (c === 3 || c === 4 || c === 5 || c === 6 || c === 7)
      range.setHorizontalAlignment("center");

    if (c === 6)
      range.setNumberFormat("M/d");
  }
}

// Colors the row ranging from columns 1 to 7
function colorRow(r) {
  var dataRange = SpreadsheetApp.getActiveSheet().getRange(r, 1, 1, 7);
  var row = dataRange.getValues()[0];

  switch(row[3]) {
    case "x":
      dataRange.setBackgroundColor("red");
      dataRange.setFontColor("black");
      break;
    case "?":
      dataRange.setBackgroundColor("yellow");
      dataRange.setFontColor("black");
      break;
    case "e":
      dataRange.setBackgroundRGB(0, 255, 0);
      dataRange.setFontColor("black");
      break;
    case "o":
      dataRange.setBackgroundRGB(77, 77, 77);
      dataRange.setFontColor("black");
      break;
    case "xx":
      dataRange.setBackgroundColor("blue");
      dataRange.setFontColor("white");
      break;
    case "xxx":
      dataRange.setBackgroundColor("black");
      dataRange.setFontColor("white");
      break;
    default:
      if (row[2] === 2 || row[2] === 1)
        dataRange.setBackgroundColor("red");
      else
        dataRange.setBackgroundColor("white");

      dataRange.setFontColor("black");
      break;
  }
}

// Runs colorRow on every row of the sheet (aside from the header)
function formatEntireSheet() {
  var startRow = 2;
  var endRow = SpreadsheetApp.getActiveSheet().getMaxRows();

  for (var r = startRow; r < endRow; r++)
    colorRow(r);
}

// Formats sheet on edit
function onEdit(event) {
  // Runs colorRow on the edited row first (for speed),
  // then formats the entire sheet
  colorRow(event.source.getActiveRange().getRowIndex());
  SpreadsheetApp.flush();

  formatEntireSheet();
  SpreadsheetApp.flush();
}

// Formats sheet on load
function onOpen() {
  formatEntireSheet();
  SpreadsheetApp.flush();
}
