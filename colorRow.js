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
  var endRow = SpreadsheetApp.getActiveSheet().getMaxRows()-1;

  for (var r = startRow; r <= endRow; r++)
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
