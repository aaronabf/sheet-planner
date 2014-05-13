function quickSortCourse() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.sort(3);
  sheet.sort(1);
}

function quickSortTime() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.sort(1);
  sheet.sort(3);
}
