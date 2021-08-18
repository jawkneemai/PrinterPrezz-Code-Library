/*
function testEmail() {
  MailApp.sendEmail('catherine.mai@printerprezz.com', 'test', 'triggered');
}
*/

// Listens to edits on the MFG Meeting Schedule for edits in the Jobs, NCMR/ECO/DCO, Facility/PM tabs for activating items onto the schedule or closing the items out into the log. Sends each item to a google doc for the log.

// Global Variables
const GREEN = '#b6d7a8';
const YELLOW = '#ffe599';
const RED = '#ea9999';

function activateTrigger(e){
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var sheetName = currentSheet.getName();
  var rowStart = e.range.rowStart;
  var scheduleSheet = masterSheet.getSheetByName('Schedule');
  

  if (sheetName == 'Schedule') {
    console.log('In Schedule, do nothing');
    return
  } else if (sheetName == 'Jobs') { // JOBS ~~~~~~~~

    // Data from row
    console.log(currentSheet);
    [printer, orderNum, travNum, startDay, amPm, printTime, owner, notes] = currentSheet.getSheetValues(rowStart, 2, 1, 8)[0];

    if (e.range.columnStart == 1 && e.range.columnEnd == 1 && e.range.rowStart <= 2000) { // ACTIVATE SCHEDULE CHECKBOXES

      if (e.value == 'TRUE') {
        console.log('Sending data to schedule');
        fillTime(scheduleSheet, getMachineRow(printer), travNum, startDay, printTime, amPm, YELLOW, notes);
        var protection = currentSheet.getRange(e.range.rowStart, (e.range.columnStart+1), 1, 8).protect();
        protection.setDescription('Row locked into schedule');
        for (let i = 0; i < protection.getEditors().length; i++) {
          console.log(protection.getEditors()[i].getEmail());
          protection.removeEditor(protection.getEditors()[i]);
        } 

      } else if (e.value == 'FALSE') {

        console.log('Removing data from schedule');
        removeTime(scheduleSheet, getMachineRow(printer), startDay, printTime, amPm);
        var protections = currentSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        for (let i = 0; i < protections.length; i++) {
          if (protections[i].getRange().getRow() == e.range.rowStart) {
            protections[i].remove();
          }
        }

      } else { console.log('Something wrong with checkbox value') }
      return


    } else if (e.range.columnStart == 10 && e.range.columnEnd == 10) { // CLOSE PRINT CHECKBOXES
      if (e.value == 'TRUE') {
        console.log('Sending row to log.');
        var rowRange = currentSheet.getRange(e.range.rowStart, 1, 1, 10);
        rowRange.clearContent();
        removeTime(scheduleSheet, getMachineRow(printer), startDay, printTime, amPm);
        var now = new Date();
        var log_string = now.toString() + ',' + printer + ',' + orderNum + ',' + travNum + ',' + startDay + ',' + amPm + ',' + printTime + ',' + owner + ',' + notes;
        console.log(log_string);

        var log = DocumentApp.openById('1FT9fU3cGBO0a5wG9Zhx_L1jjMm2f2995z22ucOHtZwU');
        var body = log.getBody();
        body.appendParagraph(log_string);

      } else { console.log('Invalid checkbox value.') }
      return
    }




  } else if (e.range.columnStart == 1 && e.range.columnEnd == 1 && e.range.rowStart <= 2000 && sheetName == 'NCMR, ECO, DCO') { // NCMR ~~~~~~~~

    // Data from Rows
    [itemType, itemNum, travNum, orderNum, startDate, amPm, time, owner, printer, notes] = currentSheet.getSheetValues(rowStart, 2, 1, 10)[0];
    
    if (e.value == 'TRUE') {

      console.log('Sending data to schedule');
      var description = itemType + ' ' + itemNum;
      if (printer == 'Metal Printers') {
        fillTime(scheduleSheet, getMachineRow('LM74'), description, startDate, time, amPm, RED, notes);
        fillTime(scheduleSheet, getMachineRow('LM36'), description, startDate, time, amPm, RED, notes);
        fillTime(scheduleSheet, getMachineRow('LM47'), description, startDate, time, amPm, RED, notes);
      } else {
        fillTime(scheduleSheet, getMachineRow(printer), description, startDate, time, amPm, RED, notes);
      }
      var protection = currentSheet.getRange(e.range.rowStart, (e.range.columnStart+1), 1, 10).protect();
        protection.setDescription('Row locked into schedule');
        for (let i = 0; i < protection.getEditors().length; i++) {
          console.log(protection.getEditors()[i].getEmail());
          protection.removeEditor(protection.getEditors()[i]);
        }

    } else if (e.value == 'FALSE') {
      console.log('Removing data from schedule');
      if (printer == 'Metal Printers') {
        removeTime(scheduleSheet, getMachineRow('LM74'), startDate, time, amPm);
        removeTime(scheduleSheet, getMachineRow('LM47'), startDate, time, amPm);
        removeTime(scheduleSheet, getMachineRow('LM36'), startDate, time, amPm);
      } else {
        removeTime(scheduleSheet, getMachineRow(printer), startDate, time, amPm);
      }
      var protections = currentSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (let i = 0; i < protections.length; i++) {
        if (protections[i].getRange().getRow() == e.range.rowStart) {
          protections[i].remove();
        }
      }

    } else { console.log('Something wrong with checkbox value') }

    return




  } else if (e.range.columnStart == 1 && e.range.columnEnd == 1 && e.range.rowStart <= 2000 && sheetName == 'Facility, PM') { // FACILITY, PM ~~~~~~~~

    // Data from Rows
    [printer, description, owner, startDay, amPm, time, notes] = currentSheet.getSheetValues(rowStart, 2, 1, 7)[0];
    
    if (e.value == 'TRUE') {
      console.log('Sending data to schedule');

      if (printer == 'Metal Printers') {
        fillTime(scheduleSheet, getMachineRow('LM74'), description, startDay, time, amPm, RED, notes);
        fillTime(scheduleSheet, getMachineRow('LM36'), description, startDay, time, amPm, RED, notes);
        fillTime(scheduleSheet, getMachineRow('LM47'), description, startDay, time, amPm, RED, notes);
      } else {
        fillTime(scheduleSheet, getMachineRow(printer), description, startDay, time, amPm, RED, notes);
      }
      var protection = currentSheet.getRange(e.range.rowStart, (e.range.columnStart+1), 1, 7).protect();
        protection.setDescription('Row locked into schedule');
        for (let i = 0; i < protection.getEditors().length; i++) {
          console.log(protection.getEditors()[i].getEmail());
          protection.removeEditor(protection.getEditors()[i]);
        }


    } else if (e.value == 'FALSE') {
      console.log('Removing data from schedule');
      if (printer == 'Metal Printers') {
        removeTime(scheduleSheet, getMachineRow('LM74'), startDay, time, amPm);
        removeTime(scheduleSheet, getMachineRow('LM47'), startDay, time, amPm);
        removeTime(scheduleSheet, getMachineRow('LM36'), startDay, time, amPm);
      } else {
        removeTime(scheduleSheet, getMachineRow(printer), startDay, time, amPm);
      }
      var protections = currentSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (let i = 0; i < protections.length; i++) {
        if (protections[i].getRange().getRow() == e.range.rowStart) {
          protections[i].remove();
        }
      }

    } else { console.log('Something wrong with checkbox value') }

    return


  } else { 
    console.log("Not in the scheduling tabs.")
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
  var endCol = col; // This column will be green on schedule AKA will end up one column after the time slot ends.
  
  isOccupied = true;
  while (isOccupied) {
    if (tempCell.getBackground() == GREEN) {
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
function fillTime(sheet, machineRow, travNum, startDate, time, amPm, statusColor, note) {
  var dateRow = sheet.getSheetValues(6,2,1,56)[0];
  var startCol = 0;
  for (let i = 0; i < dateRow.length; i++) {
    if (String(startDate) == String(dateRow[i])) {
      if (amPm == 'AM') {
        startCol = i+2;
      } else if (amPm == 'PM') {
        startCol = i+3;
      } else {
        console.log('Error with time slot.');
        return;
      }
    }
  }
  var timeBlock = Math.round(time*2);
  var endCol = startCol + timeBlock;
  
  if (endCol > 57) {
    Browser.msgBox('Time slot goes past current schedule!');
    return
  }

  var slot = sheet.getRange(machineRow, startCol, 1, timeBlock);
  slot.merge();
  slot.setBackground(statusColor);
  slot.setValue(travNum);
  slot.setHorizontalAlignment('center');
  slot.setVerticalAlignment('middle');
  slot.setNote(note);
  return
}

// Removes time slot from schedule.
function removeTime(sheet, machineRow, startDate, time, amPm) {
  var dateRow = sheet.getSheetValues(6,2,1,56)[0];
  var startCol = 0;
  var timeBlock = Math.round(time*2);
  for (let i = 0; i < dateRow.length; i++) {
    if (String(startDate) == String(dateRow[i])) {
      if (amPm == 'AM') {
        startCol = i+2;
      } else if (amPm == 'PM') {
        startCol = i+3;
      } else {
        console.log('Error with time slot.');
        return;
      }
    }
  }

  if (startCol < 1) {
    Browser.msgBox('Time slot goes before current schedule!');
    return
  }
  var slot = sheet.getRange(machineRow, startCol, 1, timeBlock);
  
  slot.setNote('');
  slot.setBackground(GREEN);
  slot.clearContent();
  slot.breakApart();
  return
}

// Highlights today whenever sheets opens, formats weekends with dotted lines
function highlightToday() {
  console.log('Sheet opened!');
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = masterSheet.getSheetByName('Printer Schedule');

  // Resetting formatting

  // Rows above and below machine rows
  scheduleSheet.getRange(1,2,4,732).setBackground('#ffffff');
  scheduleSheet.getRange(1,2,4,732).setValue('');
  scheduleSheet.getRange(5,2,2,732).setBackground('#efefef');
  scheduleSheet.getRange(12,2,2,732).setBackground('#efefef');
  scheduleSheet.getRange(14,2,5,732).setBackground('#ffffff');
  // Machine rows
  scheduleSheet.getRange(7,2,5,732);
  // for loop loop through two columns at a time(1 day at a time) and set borders appropriately

  // Highlight Today
  var dateRow = scheduleSheet.getSheetValues(6,2,1,730)[0];
  var today = new Date();
  var todayDay = today.getDay();
  var todayDate = today.getDate();
  
  todayIndex = 0;
  for (let i = 0; i < dateRow.length; i++) {
    if (dateRow[i] != '') {
      if (todayDate == dateRow[i].getDate()) {
        todayIndex = i+2;
        break;
      }
    }
  }

  scheduleSheet.getRange(1, todayIndex, 6, 2).setBackground('#d9b8df');
  scheduleSheet.getRange(1, todayIndex, 1, 1).setValue('Today');
  scheduleSheet.getRange(12, todayIndex, 7, 2).setBackground('#d9b8df');
}