// Sorts the sheet by time then the course; this results in groups
// of all entries from one course and sorts these groups by time
function sortCourse() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.sort(3);
  sheet.sort(1);
}

// Sorts the sheet by course then the time; this results in the entire sheet
// sorted by time with entries on the same day sorted by course.
function sortTime() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.sort(1);
  sheet.sort(3);
}
