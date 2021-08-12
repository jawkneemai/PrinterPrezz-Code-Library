/*
function testEmail() {
  MailApp.sendEmail('catherine.mai@printerprezz.com', 'test', 'triggered');
}
*/

function activateTrigger(e){
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var sheetName = currentSheet.getName();
  var rowStart = e.range.rowStart;
  var rowEnd = e.range.rowEnd;
  var scheduleSheet = masterSheet.getSheetByName('Schedule');

  const GREEN = '#b6d7a8';
  const YELLOW = '#ffe599';
  const RED = '#ea9999';
  

  if (sheetName == 'Schedule') {
    console.log('In Schedule, do nothing');
    return
  } else if (e.range.columnStart == 1 && e.range.columnEnd == 1 && e.range.rowStart <= 2000 && sheetName == 'Jobs') { // JOBS ~~~~~~~~
    //console.log('We are in row ' + rowStart);
    //console.log('We are in sheet ' + sheet.getName());
    [printer, orderNum, travNum, printTime,owner] = currentSheet.getSheetValues(rowStart, 2, 1, 5)[0];
    
    if (e.value == 'TRUE') {
      console.log('Sending data to schedule');
      masterSheet.setActiveSheet(scheduleSheet);
      var tempSheet = masterSheet.getActiveSheet();
      var beginCell = tempSheet.getRange(getMachineRow(printer), 2, 1, 1);
      fillTime(scheduleSheet, tempSheet.getRange(getMachineRow(printer), getNextAvailableTime(tempSheet, beginCell), 1, 1), printTime, YELLOW);
      masterSheet.setActiveSheet(currentSheet);

    } else if (e.value == 'FALSE') {

      masterSheet.setActiveSheet(scheduleSheet);
      var tempSheet = masterSheet.getActiveSheet();
      var beginCell = tempSheet.getRange(getMachineRow(printer), 2, 1, 1);
      console.log('Removing data from schedule');
      removeTime(scheduleSheet, tempSheet.getRange(getMachineRow(printer), getNextAvailableTime(tempSheet, beginCell)-1, 1, 1), printTime);
      masterSheet.setActiveSheet(currentSheet);

    } else { console.log('Something wrong with checkbox value') }

    return
    //scheduleSheet.activate(); switches to scheduleSheet





  } else if (e.range.columnStart == 1 && e.range.columnEnd == 1 && e.range.rowStart <= 2000 && sheetName == 'NCMR, ECO, DCO') { // NCMR ~~~~~~~~
    //console.log('We are in row ' + rowStart);
    //console.log('We are in sheet ' + sheet.getName());
    [printer, orderNum, travNum, printTime,owner] = currentSheet.getSheetValues(rowStart, 2, 1, 5)[0];
    
    if (e.value == 'TRUE') {

      console.log('Sending data to schedule');
      masterSheet.setActiveSheet(scheduleSheet);
      var tempSheet = masterSheet.getActiveSheet();
      var beginCell = tempSheet.getRange(getMachineRow(printer), 2, 1, 1);
      fillTime(scheduleSheet, tempSheet.getRange(getMachineRow(printer), getNextAvailableTime(tempSheet, beginCell), 1, 1), printTime, RED);
      masterSheet.setActiveSheet(currentSheet);

    } else if (e.value == 'FALSE') {
      console.log('Removing data from schedule');

      masterSheet.setActiveSheet(scheduleSheet);
      var tempSheet = masterSheet.getActiveSheet();
      var beginCell = tempSheet.getRange(getMachineRow(printer), 2, 1, 1);
      removeTime(scheduleSheet, tempSheet.getRange(getMachineRow(printer), getNextAvailableTime(tempSheet, beginCell)-1, 1, 1), printTime);
      masterSheet.setActiveSheet(currentSheet);

    } else { console.log('Something wrong with checkbox value') }

    return
    //scheduleSheet.activate(); switches to scheduleSheet





  } else if (e.range.columnStart == 1 && e.range.columnEnd == 1 && e.range.rowStart <= 2000 && sheetName == 'Facility, PM') { // FACILITY, PM ~~~~~~~~
    //console.log('We are in row ' + rowStart);
    //console.log('We are in sheet ' + sheet.getName());
    [printer, orderNum, travNum, printTime,owner] = currentSheet.getSheetValues(rowStart, 2, 1, 5)[0];
    
    if (e.value == 'TRUE') {
      console.log('Sending data to schedule');

      masterSheet.setActiveSheet(scheduleSheet);
      var tempSheet = masterSheet.getActiveSheet();
      var beginCell = tempSheet.getRange(getMachineRow(printer), 2, 1, 1);
      fillTime(scheduleSheet, tempSheet.getRange(getMachineRow(printer), getNextAvailableTime(tempSheet, beginCell), 1, 1), printTime, RED);
      masterSheet.setActiveSheet(currentSheet);

    } else if (e.value == 'FALSE') {
      console.log('Removing data from schedule');

      masterSheet.setActiveSheet(scheduleSheet);
      var tempSheet = masterSheet.getActiveSheet();
      var beginCell = tempSheet.getRange(getMachineRow(printer), 2, 1, 1);
      console.log('Removing data from schedule');
      removeTime(scheduleSheet, tempSheet.getRange(getMachineRow(printer), getNextAvailableTime(tempSheet, beginCell)-1, 1, 1), printTime);
      masterSheet.setActiveSheet(currentSheet);

    } else { console.log('Something wrong with checkbox value') }

    return
    //scheduleSheet.activate(); switches to scheduleSheet






  } else { 
    console.log("Something weird happened. We're not in the right tabs.")
    return
  }
}

// ENUM for Printers
function getMachineRow(machineName) {
  if (machineName == 'LM47') {
    return 7
  } else if (machineName == 'LM74') {
    return 8
  } else if (machineName == 'LM36') {
    return 9
  } else if (machineName == 'M290') { 
    return 10
  } else if (machineName == 'Arcam A2') {
    return 11
  } else {
    console.log('Machine Name not valid.')
    return 0
  }
}

// Gets next available time slot on the schedule for given start cell (row)
function getNextAvailableTime(sheet, startCell) {
  var tempCell = startCell;
  var row = startCell.getRow();
  var col = startCell.getColumn();
  var endCol = col; // This column will be white on schedule
  
  isOccupied = true;
  while (isOccupied) {
    if (tempCell.getBackground() == '#ffffff') {
      isOccupied = false;
      break;
    } else {
      col += 1;
      tempCell = sheet.getRange(row, col, 1, 1);
    }
  }

  endCol = col;
  return endCol
}

// Fills schedule with designated time slot
// TODO: lock columns in current sheet so can't take away time from a different printer
function fillTime(sheet, startCell, time, statusColor) {
  var row = startCell.getRow();
  var startCol = startCell.getColumn();
  var currentCol = startCol;
  var tempCell = startCell;
  var timeBlock = Math.round(time*2);
  var endCol = startCol + timeBlock;
  if (endCol > 29) {
    Browser.msgBox('Time slot goes past current schedule!');
    return
  }

  isDone = true;
  while (isDone) {
    if (currentCol == endCol) {
      tempCell = sheet.getRange(row, currentCol, 1, 1);
      tempCell.setBorder(null, true, null, null, null, null);
      isDone = false;
    } else {
      tempCell = sheet.getRange(row, currentCol, 1, 1);
      tempCell.setBackground(statusColor);
      currentCol += 1;
    }
  }
  return
}

function removeTime(sheet, startCell, time) {
  var row = startCell.getRow();
  var startCol = startCell.getColumn();
  var currentCol = startCol;
  var tempCell = startCell;
  var timeBlock = Math.round(time*2);
  var endCol = startCol - timeBlock;
  if (endCol < 1) {
    Browser.msgBox('Time slot goes before current schedule!');
    return
  }

  isDone = true;
  while (isDone) {
    if (currentCol == endCol) {
      isDone = false;
    } else {
      tempCell = sheet.getRange(row, currentCol, 1, 1);
      tempCell.setBackground('#ffffff');
      currentCol -= 1;
    }
  }
  return
}








