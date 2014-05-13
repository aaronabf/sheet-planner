function colorRow(r)
{
  var dataRange = SpreadsheetApp.getActiveSheet().getRange(r, 1, 1, 7);

  var row = dataRange.getValues()[0];

  if(row[3] === "x"){
    dataRange.setBackgroundColor("red");
  }
  else if(row[3] === "?"){
    dataRange.setBackgroundColor("yellow");
  }
  else if(row[3] === "e"){
    dataRange.setBackgroundRGB(0, 255, 0);
  }
  else if(row[3] === "o"){
    dataRange.setBackgroundRGB(77, 77, 77);
  }
  else if(row[3] === "xx"){
    dataRange.setBackgroundColor("blue");
    dataRange.setFontColor("white");
  }
  else if(row[3] === "xxx"){
    dataRange.setBackgroundColor("black");
    dataRange.setFontColor("white");
  }
  else if(row[2] === 2 || row[2] === 1){
    dataRange.setBackgroundColor("red");
  }
  else{
    dataRange.setBackgroundColor("white");
    dataRange.setFontColor("black");
  }
}

function formatEntireSheet()
{
  var startRow = 2;
  var endRow = SpreadsheetApp.getActiveSheet().getMaxRows()-1;

  for (var r = startRow; r <= endRow; r++){
    colorRow(r);
  }
}

function onEdit(event)
{
  // We want to colorRow on the edited row for
  // speed, and then run the entire sheet
  colorRow(event.source.getActiveRange().getRowIndex());
  formatEntireSheet();
}

function onOpen()
{
  formatEntireSheet();
  SpreadsheetApp.flush();
}
