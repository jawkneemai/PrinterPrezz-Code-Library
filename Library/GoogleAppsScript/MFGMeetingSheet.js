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
  var scheduleSheet = masterSheet.getSheetByName('Printer Schedule');
  

  if (sheetName == 'Printer Schedule') {
    console.log('In Schedule, do nothing');
    return




// ~~~~~~~~~~~~~~~~~~~ Print Jobs ~~~~~~~~~~~~~~~~~~~~





  } else if (sheetName == 'Print Jobs') { // JOBS ~~~~~~~~

    // Data from row
    [printer, orderNum, travNum, startDay, amPm, printTime, owner, notes] = currentSheet.getSheetValues(rowStart, 2, 1, 8)[0];
    var rowData = {
      "Printer": printer,
      "Order #": orderNum,
      "Traveler #": travNum,
      "Start Date (MM/DD)": startDay,
      "AM or PM": amPm,
      "Time (days)": printTime,
      "Owner": owner,
      "Notes": notes
    };





    if (e.range.columnStart == 1 && e.range.columnEnd == 1) { // ACTIVATE SCHEDULE CHECKBOXES

      if (e.value == 'TRUE') {
        console.log('Sending data to schedule');
        if (fillTime(scheduleSheet, getMachineRow(printer), travNum, startDay, printTime, amPm, YELLOW, notes) == 1) {
          var protection = currentSheet.getRange(e.range.rowStart, (e.range.columnStart+1), 1, 8).protect();
          protection.setDescription('Row locked into schedule');
          for (let i = 0; i < protection.getEditors().length; i++) {
            console.log(protection.getEditors()[i].getEmail());
            protection.removeEditor(protection.getEditors()[i]);
          }
        } else {
          currentSheet.getRange(e.range.rowStart, e.range.columnStart, 1, 1).uncheck();
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
        var logSheet = masterSheet.getSheetByName('Print, NCMR, ECO, DCO, DEV Log');
        sendEventToLog(logSheet, rowData);
        currentSheet.getRange(e.range.rowStart, 1, 1, 10).clearContent();
      } else { console.log('Invalid checkbox value.') }
      return
    }




// ~~~~~~~~~~~~~~~~~~~ NCMR, ECO, DCO, DEV ~~~~~~~~~~~~~~~~~~~~




  } else if (sheetName == 'NCMR, ECO, DCO, DEV') { 

    // Data from Rows
    [itemType, itemNum, travNum, orderNum, startDate, amPm, time, owner, printer, notes] = currentSheet.getSheetValues(rowStart, 2, 1, 10)[0];

    var rowData = {
      "Item #": itemNum,
      "Item Type": itemType,
      "Printer": printer,
      "Order #": orderNum,
      "Traveler #": travNum,
      "Start Date (MM/DD)": startDate,
      "AM or PM": amPm,
      "Time (days)": time,
      "Owner": owner,
      "Notes": notes
    };
    



    if (e.range.columnStart == 1 && e.range.columnEnd == 1) { // ACTIVATE BOXES

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
      }

    } else if (e.range.columnStart == 12 && e.range.columnEnd == 12) { // CLOSE PRINT CHECKBOXES
      if (e.value == 'TRUE') {
        console.log('Sending row to log.');
        var logSheet = masterSheet.getSheetByName('Print, NCMR, ECO, DCO, DEV Log');
        sendEventToLog(logSheet, rowData);
        currentSheet.getRange(e.range.rowStart, 1, 1, 12).clearContent();

      }

    } else { console.log('Something wrong with checkbox value') }
    return



// ~~~~~~~~~~~~~~~~~~~~~ FACILITY, PM ~~~~~~~~~~~~~~~~~~~~~~~~~




  } else if (sheetName == 'Facility, PM') { 

    // Data from Rows
    [printer, description, owner, startDay, amPm, time, notes] = currentSheet.getSheetValues(rowStart, 2, 1, 7)[0];

    var rowData = {
      "PM/Facility": description,
      "Printer": printer,
      "Start Date (MM/DD)": startDay,
      "AM or PM": amPm,
      "Time (days)": time,
      "Owner": owner,
      "Notes": notes
    };
  
    if (e.range.columnStart == 1 && e.range.columnEnd == 1) {

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
      }



    } else if (e.range.columnStart == 9 && e.range.columnEnd == 9) { // CLOSE PRINT CHECKBOXES
      if (e.value == 'TRUE') {
        console.log('Sending row to log.');
        var logSheet = masterSheet.getSheetByName('Print, NCMR, ECO, DCO, DEV Log');
        sendEventToLog(logSheet, rowData);
        currentSheet.getRange(e.range.rowStart, 1, 1, 9).clearContent();

      } else { console.log('Invalid checkbox value.') }
    }
    return

    



// ~~~~~~~~~~~~~~~~~~~ Open Business ~~~~~~~~~~~~~~~~~~~~~




  } else if (sheetName == 'Open Business') { 

    if (e.range.columnStart == 3 && e.range.columnEnd == 3) {

      [description, owners] = currentSheet.getSheetValues(rowStart, 1, 1, 2)[0];
      var rowData = {
        "Item Description": description,
        "Owner(s)": owners
      };
      if (e.value == 'TRUE') {
        console.log('Sending row to log.');
        var logSheet = masterSheet.getSheetByName('Open Business Log');
        sendOpenBusinessToLog(logSheet, rowData);
        currentSheet.getRange(e.range.rowStart, 1, 1, 3).clearContent();
      } else { console.log('Invalid checkbox value.') }

    }
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
  if (slot.getBackground() == GREEN) {
    slot.merge();
    slot.setBackground(statusColor);
    slot.setValue(travNum);
    slot.setHorizontalAlignment('center');
    slot.setVerticalAlignment('middle');
    slot.setNote(note);
    return 1;
  } else {
    Browser.msgBox('Time Conflict!');
    return 0;
  }
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
  var machineRows = scheduleSheet.getRange(7,2,5,732);

// Uncomment when need to reformat weekends
/* 
  for (let i = 1; i < 366; i++) {
    if (scheduleSheet.getRange(5, (2*i), 1, 1).getValue().getDay() == 0 || scheduleSheet.getRange(5, (2*i), 1, 1).getValue().getDay() == 6) {
      var tempRange = scheduleSheet.getRange(7,(2*i), 5, 2);
      tempRange.setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      tempRange.setBorder(null, null, null, null, true, true, '#000000', SpreadsheetApp.BorderStyle.DOTTED);
    }
  }
*/

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

// Sends an item in Print Jobs or NCMR, ECO, DCO, DEV to the Event Log tab
function sendEventToLog(logSheet, rowData) {
  var categories = ['Time Stamp', 'Traveler #', 'Order #', 'Item Type', 'Item #', 'PM/Facility', 'Start Date (MM/DD)', 'AM or PM', 'Time (days)', 'Owner', 'Printer', 'Notes'];
  var now = new Date();
  var timeStamp = now.toString();
  var emptyRow = getFirstEmptyRowByColumnArray(logSheet);
  logSheet.getRange(emptyRow, 1, 1, 1).setValue(timeStamp);
  for (let i = 1; i < categories.length; i++) {
    if (categories[i] in rowData) {
      logSheet.getRange(emptyRow, i+1, 1, 1).setValue(rowData[categories[i]]);
    } else {
      logSheet.getRange(emptyRow, i+1, 1, 1).setValue('N/A');
    }
  }
}

// Sends an item in Open Business tab to the Open Business Log
function sendOpenBusinessToLog(logSheet, rowData) {
  var categories = ['Item Description', 'Owner(s)'];
  var now = new Date();
  var timeStamp = now.toString();
  var emptyRow = getFirstEmptyRowByColumnArray(logSheet);
  logSheet.getRange(emptyRow, 1, 1, 1).setValue(timeStamp);
  for (let i = 0; i < categories.length; i++) {
    if (categories[i] in rowData) {
      console.log(rowData[categories[i]]);
      logSheet.getRange(emptyRow, i+2, 1, 1).setValue(rowData[categories[i]]);
    } else {
      logSheet.getRange(emptyRow, i+2, 1, 1).setValue('N/A');
    }
  }
}

// Gets the first empty row on a given sheet.
function getFirstEmptyRowByColumnArray(sheet) {
  var column = sheet.getRange('A:A');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct] && values[ct][0] != "" ) {
    ct++;
  }
  return (ct+1);
}

