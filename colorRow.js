// Colors columns between 1 and 7 in the row r
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

// Formats sheet on edit; runs colorRow on the edited row first (so the
// use can see the change quickly), then formats the entire sheet
function onEdit(event) {
  // Only run the script on the first sheet in the spreadsheet
  if (event.source.getActiveSheet().getIndex() !== 1)
    return;

  colorRow(event.source.getActiveRange().getRowIndex());
  SpreadsheetApp.flush();

  formatEntireSheet();
  SpreadsheetApp.flush();
}

// Formats sheet on open
function onOpen() {
  formatEntireSheet();
  SpreadsheetApp.flush();
}
