// Color config
var OFF_CELL_COLOR = "#efefef";
var SECTION_HEADER_COLOR = "#cfe2f3";

// Column header config
var PRIORITY_HEADER = "Priority";

// Priority column annotation config
var UNFORMATTED_BACKLOG_INDICATOR = "*";
var FORMATTED_BACKLOG_INDICATOR = "•";
var UNFORMATTED_DONE_INDICATOR = "!";
var FORMATTED_DONE_INDICATOR = "✓";
var ALLOW_ARBITRARY_PRIORITIES = false;

// Buffer space: this determines the number of rows between the Backlog and Done sections
var BUFFER_SPACE = 8;

// This function parses the existing to-do cells, categorizes and sorts them, and then writes the results 
// back onto the sheet.
//
// Note: Data is zero-indexed, sheet ranges are 1-indexed.
// Index variables with "Sheet" in their names are therefore 1-indexed.
function sortTasks() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var totalRowCount = data.length;

  var toDoTasks = [];
  var backlogTasks = [];
  var doneTasks = [];

  var headerRowSheetIndex = 1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] != '' && data[i][0] == PRIORITY_HEADER) {
      headerRowSheetIndex = i + 1;
      break;
    }
  }

  for (var i = headerRowSheetIndex + 1; i < data.length; i++) {
    row = data[i];
    key = row[0];
    if ((typeof key == "string" && isBacklogIndicator(key))
    || (key == '' && (row[1] != '' || row[2] != ''))) {
      row[0] = FORMATTED_BACKLOG_INDICATOR;
      backlogTasks.push(row);
    } else if (typeof key == "string" && isDoneIndicator(key)) {
      row[0] = FORMATTED_DONE_INDICATOR;
      doneTasks.push(row);
    } else if (isNumber(row[0])) {
      toDoTasks.push(row);
    }
  }

  // Sort task row numerically, and move empty columns last
  toDoTasks.sort(function(inp1, inp2) {
    if (inp1[0] == '') {
      return 1;
    } else if (inp2[0] == '') {
      return -1;
    }
    return inp1[0] - inp2[0];
  });

  // Write to do tasks
  toDoRow = sheet.getRange(headerRowSheetIndex + 1, 1, 1, 3);
  clearRange(toDoRow);
  toDoRow.setBackground(SECTION_HEADER_COLOR);
  toDoCell = sheet.getRange(headerRowSheetIndex + 1, 1);
  toDoCell.setFontWeight("bold");
  toDoCell.setValue("To Do");

  var currentRowCounter = headerRowSheetIndex + 2;
  var oddCounter = 0;
  for (var i = 0; i < toDoTasks.length; i++) {
    var priority = toDoTasks[i][0];
    var task = toDoTasks[i][1];
    var note = toDoTasks[i][2];

    // Only write the row if there is a priority or task.
    if (priority != '' || task != '') {
      var row = sheet.getRange(i + currentRowCounter, 1, 1, 3);
      clearRange(row);
      sheet.getRange(i + currentRowCounter, 2, 1, 2).setWrap(true);
      if (oddCounter % 2 == 1) {
        row.setBackground(OFF_CELL_COLOR);
      }

      row.setValues([[ALLOW_ARBITRARY_PRIORITIES ? priority : i + 1, task, note]]);

    }
    oddCounter += 1;
  }

  currentRowCounter += toDoTasks.length;
  clearRange(sheet.getRange(currentRowCounter, 1, 1, 3));
  currentRowCounter += 1;

  // Write backlog tasks
  backlogRow = sheet.getRange(currentRowCounter, 1, 1, 3);
  clearRange(backlogRow);
  backlogRow.setBackground(SECTION_HEADER_COLOR);
  backlogCell = sheet.getRange(currentRowCounter, 1);
  backlogCell.setFontWeight("bold");
  backlogCell.setValue("Backlog");

  currentRowCounter += 1;
  oddCounter = 0;
  for (var i = 0; i < backlogTasks.length; i++) {
    var backlog_marker = backlogTasks[i][0];
    var task = backlogTasks[i][1];
    var note = backlogTasks[i][2];

    if (priority != '' || task != '') {
      var row = sheet.getRange(i + currentRowCounter, 1, 1, 3);
      clearRange(row);

      sheet.getRange(i + currentRowCounter, 2, 1, 2).setWrap(true);
      if (oddCounter % 2 == 1) {
        row.setBackground(OFF_CELL_COLOR);
      }

      row.setValues([[backlog_marker, task, note]]);
    }
    oddCounter += 1;
  }

  currentRowCounter += backlogTasks.length;
  clearRange(sheet.getRange(currentRowCounter, 1, BUFFER_SPACE, 3));
  currentRowCounter += BUFFER_SPACE;

  // Write done tasks
  doneRow = sheet.getRange(currentRowCounter, 1, 1, 3);
  clearRange(doneRow);
  doneRow.setBackground(SECTION_HEADER_COLOR);
  doneCell = sheet.getRange(currentRowCounter, 1);
  doneCell.setFontWeight("bold");
  doneCell.setValue("Done");

  currentRowCounter += 1 ;
  oddCounter = 0;
  for (var i = 0; i < doneTasks.length; i++) {
    var done_marker = doneTasks[i][0];
    var task = doneTasks[i][1];
    var note = doneTasks[i][2];

    if (priority != '' || task != '') {
      var row = sheet.getRange(i + currentRowCounter, 1, 1, 3);
      clearRange(row);

      sheet.getRange(i + currentRowCounter, 2, 1, 2).setFontLine('line-through');
      sheet.getRange(i + currentRowCounter, 2, 1, 2).setWrap(true);
      if (oddCounter % 2 == 1) {
        row.setBackground(OFF_CELL_COLOR);
      }

      row.setValues([[done_marker, task, note]]);
    }
    oddCounter += 1;
  }

  currentRowCounter += doneTasks.length;

  // Delete any leftover rows
  if (currentRowCounter <= totalRowCount) {
    clearRange(sheet.getRange(currentRowCounter, 1, totalRowCount - currentRowCounter + 1, 3));
  }
}

function clearRange(range) {
  range.setValue("");
  range.clearFormat();
}

function isNumber(num) {
  return typeof num === 'number' && isFinite(num);
}

function isBacklogIndicator(key) {
  return key.trim() == UNFORMATTED_BACKLOG_INDICATOR || key.trim() == FORMATTED_BACKLOG_INDICATOR;
}

function isDoneIndicator(key) {
  return key.trim() == UNFORMATTED_DONE_INDICATOR || key.trim() == FORMATTED_DONE_INDICATOR;
}

// This function will be called every time the sheet is opened.
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive()
  var menuItems = [
    {name: 'Sort Tasks', functionName: 'sortTasks'}
  ];
  spreadsheet.addMenu('Functions', menuItems);
}
