// Listens to edits on the MFG Meeting Schedule for edits in the Jobs, NCMR/ECO/DCO, Facility/PM tabs for activating items onto the schedule or closing the items out into the log. Sends each item to a google doc for the log.

// Global Variables
const GREEN = '#b6d7a8';
const YELLOW = '#ffe599';
const RED = '#ea9999';

// Sheets: ['3 Month Schedule', '1: Print Jobs', '2: NCMR, ECO, DCO, DEV', '3: PM', '4: Facility', '5: Open Business', 'Printer Schedule', 'Master Log', '1: Print Job Logs', '2: NCMR, ECO, DCO, DEV Logs', '3: EHF Logs', '4: Facility Logs', '5: Open Business Log', 'Drop Down Lists', 'Notes']

function activateTrigger(e){
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = SpreadsheetApp.getActiveSheet();
  var sheetName = currentSheet.getName();
  var rowStart = e.range.rowStart;
  var scheduleSheet = masterSheet.getSheetByName('3 Month Schedule');
  var scheduleLog = masterSheet.getSheetByName('Printer Schedule');
  
  if (sheetName === 'Printer Schedule' || sheetName === '3 Month Schedule' || sheetName === '5: Open Business Log' || sheetName === 'Drop Down Lists' || sheetName === '1: Print Job Log' || sheetName === '2: NCMR, ECO, DCO, DEV Log' || sheetName === '4: EHF Log' || sheetName === '3: Facility Log' || sheetName === 'Notes') {
    console.log(sheetName);
    console.log('In a sheet where nothing is supposed to be edited.');
    return




// ~~~~~~~~~~~~~~~~~~~ Print Jobs ~~~~~~~~~~~~~~~~~~~~





  } else if (sheetName == '1: Print Jobs') { // JOBS ~~~~~~~~

    // Data from row
    [printer, orderNum, travNum, startDay, amPm, printTime, owner, notes, totalBuildTime, totalBuildHeight] = currentSheet.getSheetValues(rowStart, 3, 1, (currentSheet.getLastColumn()-3))[0];
    var rowData = {
      "Printer": printer,
      "Order #": orderNum,
      "Traveler #": travNum,
      "Start Date (MM/DD)": startDay,
      "AM or PM": amPm,
      "Time (days)": printTime,
      "Owner": owner,
      "Notes": notes,
      "Total Build Time (hours)": totalBuildTime,
      "Total Build Height (mm)": totalBuildHeight
    };





    if (e.range.columnStart == 1 && e.range.columnEnd == 1) { // ACTIVATE SCHEDULE CHECKBOXES

      if (e.value == 'TRUE') {
        console.log('Sending data to schedule');
        if (fillTime(scheduleSheet, getMachineRow(printer), travNum, startDay, printTime, amPm, YELLOW, notes) == 1) {
          
        } else {
          currentSheet.getRange(e.range.rowStart, e.range.columnStart, 1, 1).uncheck();
        }

      } else if (e.value == 'FALSE') {

        console.log('Removing data from schedule');
        removeTime(scheduleSheet, getMachineRow(printer), startDay, printTime, amPm);
        
      } else { console.log('Something wrong with checkbox value') }
      return




    } else if (e.range.columnStart == currentSheet.getLastColumn() && e.range.columnEnd == currentSheet.getLastColumn()) { // CLOSE PRINT CHECKBOXES
      if (e.value == 'TRUE') {
        console.log('Sending row to log.');
        sendEventToLog(currentSheet, masterSheet.getSheetByName('Master Log'), masterSheet.getSheetByName('1: Print Job Log'), rowData);
        fillTime(scheduleLog, getMachineRow(printer), travNum, startDay, printTime, amPm, YELLOW, notes);
        currentSheet.getRange(e.range.rowStart, 1, 1, currentSheet.getLastColumn()).clearContent();
      } else { console.log('Invalid checkbox value.') }
      return





    } else if (e.range.rowStart > 1 && 2 < e.range.columnStart < currentSheet.getLastColumn()) { // UPDATE TIMESLOT IF VALUE IS CHANGED AND ACTIVE

      if (currentSheet.getRange(e.range.rowStart, 1, 1, 1).getValue() == true) {
        var changedCat = currentSheet.getRange(1, e.range.columnStart, 1, 1).getValue();
        rowData[changedCat] = e.oldValue;
        removeTime(scheduleSheet, getMachineRow(rowData['Printer']), rowData['Start Date (MM/DD)'], rowData['Time (days)'], rowData['AM or PM']);
        fillTime(scheduleSheet, getMachineRow(printer), travNum, startDay, printTime, amPm, YELLOW, notes);
      } 
      return;
    }




// ~~~~~~~~~~~~~~~~~~~ NCMR, ECO, DCO, DEV ~~~~~~~~~~~~~~~~~~~~




  } else if (sheetName == '2: NCMR, ECO, DCO, DEV') { 

    // Data from Rows
    [itemType, itemNum, vov, travNum, orderNum, startDay, amPm, time, owner, printer, notes] = currentSheet.getSheetValues(rowStart, 2, 1, (currentSheet.getLastColumn()-3))[0];

    var description = itemType + ' ' + itemNum;
    var rowData = {
      "Item #": itemNum,
      "Item Type": itemType,
      'Verification vs. Validation': vov,
      "Printer": printer,
      "Order #": orderNum,
      "Traveler #": travNum,
      "Start Date (MM/DD)": startDay,
      "AM or PM": amPm,
      "Time (days)": time,
      "Owner": owner,
      "Notes": notes
    };
    
    console.log(rowData);



    if (e.range.columnStart == 1 && e.range.columnEnd == 1) { // ACTIVATE BOXES

      if (e.value == 'TRUE') {

        console.log('Sending data to schedule');
        if (printer == 'Metal Printers') {
          fillTime(scheduleSheet, getMachineRow('LM74'), description, startDay, time, amPm, RED, notes);
          fillTime(scheduleSheet, getMachineRow('LM36'), description, startDay, time, amPm, RED, notes);
          fillTime(scheduleSheet, getMachineRow('LM47'), description, startDay, time, amPm, RED, notes);
        } else {
          fillTime(scheduleSheet, getMachineRow(printer), description, startDay, time, amPm, RED, notes);
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

      }

    } else if (e.range.columnStart == currentSheet.getLastColumn() && e.range.columnEnd == currentSheet.getLastColumn()) { // CLOSE PRINT CHECKBOXES
      if (e.value == 'TRUE') {
        console.log('Sending row to log.');
        var logSheet = masterSheet.getSheetByName('Master Log');
        sendEventToLog(currentSheet, masterSheet.getSheetByName('Master Log'), masterSheet.getSheetByName('2: NCMR, ECO, DCO, DEV Log'), rowData);
        if (printer == 'Metal Printers') {
          fillTime(scheduleLog, getMachineRow('LM74'), description, startDay, time, amPm, RED, notes);
          fillTime(scheduleLog, getMachineRow('LM36'), description, startDay, time, amPm, RED, notes);
          fillTime(scheduleLog, getMachineRow('LM47'), description, startDay, time, amPm, RED, notes);
        } else {
          fillTime(scheduleLog, getMachineRow(printer), description, startDay, time, amPm, RED, notes);
        }
        currentSheet.getRange(e.range.rowStart, 1, 1, 12).clearContent();


      }





    } else if (e.range.rowStart > 1 && 2 < e.range.columnStart < currentSheet.getLastColumn()) { // UPDATE TIME SLOT IF VALUE IS CHANGED AND ACTIVE

      if (currentSheet.getRange(e.range.rowStart, 1, 1, 1).getValue() == true) {
        var changedCat = currentSheet.getRange(1, e.range.columnStart, 1, 1).getValue();
        rowData[changedCat] = e.oldValue;
        
        
        if (changedCat == 'Printer') { // IF CHANGING PRINTER TO AND FROM MULTIPLE AFFECTED PRINTERS
          if (printer == 'Metal Printers') { // 
            removeTime(scheduleSheet, getMachineRow(rowData['Printer']), rowData['Start Date (MM/DD)'], rowData['Time (days)'], rowData['AM or PM']);
            fillTime(scheduleSheet, getMachineRow('LM74'), description, startDay, time, amPm, RED, notes);
            fillTime(scheduleSheet, getMachineRow('LM36'), description, startDay, time, amPm, RED, notes);
            fillTime(scheduleSheet, getMachineRow('LM47'), description, startDay, time, amPm, RED, notes);
          } else if (e.oldValue == 'Metal Printers') {
            removeTime(scheduleSheet, getMachineRow('LM74'), rowData['Start Date (MM/DD)'], rowData['Time (days)'], rowData['AM or PM']);
            removeTime(scheduleSheet, getMachineRow('LM36'), rowData['Start Date (MM/DD)'], rowData['Time (days)'], rowData['AM or PM']);
            removeTime(scheduleSheet, getMachineRow('LM47'), rowData['Start Date (MM/DD)'], rowData['Time (days)'], rowData['AM or PM']);
            fillTime(scheduleSheet, getMachineRow(printer), description, startDay, time, amPm, RED, notes);
          } else {
            removeTime(scheduleSheet, getMachineRow(rowData['Printer']), rowData['Start Date (MM/DD)'], rowData['Time (days)'], rowData['AM or PM']);
            fillTime(scheduleSheet, getMachineRow(printer), description, startDay, time, amPm, RED, notes);
          } 
        } else { 
          if (printer == 'Metal Printers') {
            removeTime(scheduleSheet, getMachineRow('LM74'), rowData['Start Date (MM/DD)'], rowData['Time (days)'], rowData['AM or PM']);
            removeTime(scheduleSheet, getMachineRow('LM36'), rowData['Start Date (MM/DD)'], rowData['Time (days)'], rowData['AM or PM']);
            removeTime(scheduleSheet, getMachineRow('LM47'), rowData['Start Date (MM/DD)'], rowData['Time (days)'], rowData['AM or PM']);
            fillTime(scheduleSheet, getMachineRow('LM74'), description, startDay, time, amPm, RED, notes);
            fillTime(scheduleSheet, getMachineRow('LM36'), description, startDay, time, amPm, RED, notes);
            fillTime(scheduleSheet, getMachineRow('LM47'), description, startDay, time, amPm, RED, notes);
          } else {
            removeTime(scheduleSheet, getMachineRow(rowData['Printer']), rowData['Start Date (MM/DD)'], rowData['Time (days)'], rowData['AM or PM']);
            fillTime(scheduleSheet, getMachineRow(printer), description, startDay, time, amPm, RED, notes);
          }
        }
      } 
      return;
    }
    


// ~~~~~~~~~~~~~~~~~~~~~ FACILITY ~~~~~~~~~~~~~~~~~~~~~~~~~




  } else if (sheetName == '3: Facility') { 

    // Data from Rows
    [printer, description, owner, startDay, amPm, time, notes] = currentSheet.getSheetValues(rowStart, 2, 1, (currentSheet.getLastColumn()-3))[0];

    var rowData = {
      "Facility": description,
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
        


      } else if (e.value == 'FALSE') {
        console.log('Removing data from schedule');
        if (printer == 'Metal Printers') {
          removeTime(scheduleSheet, getMachineRow('LM74'), startDay, time, amPm);
          removeTime(scheduleSheet, getMachineRow('LM47'), startDay, time, amPm);
          removeTime(scheduleSheet, getMachineRow('LM36'), startDay, time, amPm);
        } else {
          removeTime(scheduleSheet, getMachineRow(printer), startDay, time, amPm);
        }
      }





    } else if (e.range.columnStart == currentSheet.getLastColumn() && e.range.columnEnd == currentSheet.getLastColumn()) { // CLOSE PRINT CHECKBOXES
      if (e.value == 'TRUE') {
        console.log('Sending row to log.');
        var logSheet = masterSheet.getSheetByName('Master Log');
        sendEventToLog(currentSheet, masterSheet.getSheetByName('Master Log'), masterSheet.getSheetByName('3: Facility Log'), rowData);
        if (printer == 'Metal Printers') {
          fillTime(scheduleLog, getMachineRow('LM74'), description, startDay, time, amPm, RED, notes);
          fillTime(scheduleLog, getMachineRow('LM36'), description, startDay, time, amPm, RED, notes);
          fillTime(scheduleLog, getMachineRow('LM47'), description, startDay, time, amPm, RED, notes);
        } else {
          fillTime(scheduleLog, getMachineRow(printer), description, startDay, time, amPm, RED, notes);
        }
        currentSheet.getRange(e.range.rowStart, 1, 1, 9).clearContent();

      } else { console.log('Invalid checkbox value.') }





    } else if (e.range.rowStart > 1 && 2 < e.range.columnStart < currentSheet.getLastColumn()) { // EDITING MIDDLE AREA

      if (currentSheet.getRange(e.range.rowStart, 1, 1, 1).getValue() == true) {
        var changedCat = currentSheet.getRange(1, e.range.columnStart, 1, 1).getValue();
        rowData[changedCat] = e.oldValue;
        removeTime(scheduleSheet, getMachineRow(rowData['Printer']), rowData['Start Date (MM/DD)'], rowData['Time (days)'], rowData['AM or PM']);
        fillTime(scheduleSheet, getMachineRow(printer), description, startDay, time, amPm, RED, notes);
      } 
      return;
    }
    return






// ~~~~~~~~~~~~~~~~~~~ PM ~~~~~~~~~~~~~~~~~~~~~




  } else if (sheetName == 'PM') { 

    // Data from Rows
    [printer, description, owner, startDay, amPm, time, notes, pmType, cleaning, inspectionChecks, replacementOfParts, repairRequired, calibration, adjustmentsRequired, upgrades, otherTasks, tasks, frequency, reasonObservedIssue, fseWorking] = currentSheet.getSheetValues(rowStart, 2, 1, (currentSheet.getLastColumn()-3))[0];

    var rowData = {
      "PM": description,
      "Printer": printer,
      "Start Date (MM/DD)": startDay,
      "AM or PM": amPm,
      "Time (days)": time,
      "Owner": owner,
      "Notes": notes,
      "PM Type": pmType,
      "Cleaning (Y/N)": cleaning,
      "Inspection Checks (Y/N)": inspectionChecks,
      "Replacement of Parts (Y/N)": replacementOfParts,
      "Repair Required (Y/N)": repairRequired,
      "Calibration (Y/N)": calibration,
      "Adjustments Required (Y/N)": adjustmentsRequired,
      "Upgrades (Y/N)": upgrades,
      "Other Tasks": otherTasks,
      "Tasks": tasks,
      "Frequency": frequency,
      "Reason/Observed Issue": reasonObservedIssue,
      "FSE Working": fseWorking
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
        


      } else if (e.value == 'FALSE') {
        console.log('Removing data from schedule');
        if (printer == 'Metal Printers') {
          removeTime(scheduleSheet, getMachineRow('LM74'), startDay, time, amPm);
          removeTime(scheduleSheet, getMachineRow('LM47'), startDay, time, amPm);
          removeTime(scheduleSheet, getMachineRow('LM36'), startDay, time, amPm);
        } else {
          removeTime(scheduleSheet, getMachineRow(printer), startDay, time, amPm);
        }
      }
        



    } else if (e.range.columnStart == currentSheet.getLastColumn() && e.range.columnEnd == currentSheet.getLastColumn()) { // CLOSE PRINT CHECKBOXES

      //console.log(rowData.every(element => element === null));
      if (e.value == 'TRUE') {
        console.log('Sending row to log.');

        sendEventToLog(currentSheet, masterSheet.getSheetByName('Master Log'), masterSheet.getSheetByName('4: EHF Log'), rowData);

        if (printer == 'Metal Printers') {
          fillTime(scheduleLog, getMachineRow('LM74'), description, startDay, time, amPm, RED, notes);
          fillTime(scheduleLog, getMachineRow('LM36'), description, startDay, time, amPm, RED, notes);
          fillTime(scheduleLog, getMachineRow('LM47'), description, startDay, time, amPm, RED, notes);
        } else {
          fillTime(scheduleLog, getMachineRow(printer), description, startDay, time, amPm, RED, notes);
        }
        currentSheet.getRange(e.range.rowStart, 1, 1, 22).clearContent();

      } else { console.log('Invalid checkbox value.') }





    } else if (e.range.rowStart > 1 && 2 < e.range.columnStart < 10 && 2 < e.range.columnEnd < 10) { // EDITING MIDDLE AREA

      if (currentSheet.getRange(e.range.rowStart, 1, 1, 1).getValue() == true) {
        var changedCat = currentSheet.getRange(1, e.range.columnStart, 1, 1).getValue();
        rowData[changedCat] = e.oldValue;
        removeTime(scheduleSheet, getMachineRow(rowData['Printer']), rowData['Start Date (MM/DD)'], rowData['Time (days)'], rowData['AM or PM']);
        fillTime(scheduleSheet, getMachineRow(printer), description, startDay, time, amPm, RED, notes);
      } 
      return;
    }
    return






// ~~~~~~~~~~~~~~~~~~~ Open Business ~~~~~~~~~~~~~~~~~~~~~




  } else if (sheetName == 'Open Business') { 

    if (e.range.columnStart == currentSheet.getLastColumn() && e.range.columnEnd == currentSheet.getLastColumn()) { // CLOSING OPEN BUSINESS ITEM TO LOG

      [description, owners] = currentSheet.getSheetValues(rowStart, 1, 1, 2)[0];
      var rowData = {
        "Item Description": description,
        "Owner(s)": owners
      };
      if (e.value == 'TRUE') {
        console.log('Sending row to log.');
        var logSheet = masterSheet.getSheetByName('5: Open Business Log');
        sendOpenBusinessToLog(logSheet, rowData);
        currentSheet.getRange(e.range.rowStart, 1, 1, currentSheet.getLastColumn()).clearContent();
      } else { console.log('Invalid checkbox value.') }
    }
    return







  } else {
    console.log("In a weird tab!")
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
  } else if (machineName == 'No Printer Affected') {
    return 0
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
function fillTime(sheet, machineRow, display, startDate, time, amPm, statusColor, note) {
  if (machineRow == 0 || display.includes('Dev.') || display.includes('DCO') ) {
    console.log('This event doesnt need to go on the schedule.');
    return;
  }
  var dateRow = sheet.getSheetValues(6,2,1,sheet.getLastColumn()-2)[0];
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
  
  if (endCol > 731) {
    Browser.msgBox('Time slot goes past current schedule!');
    return
  }

  var slot = sheet.getRange(machineRow, startCol, 1, timeBlock);
  
  if (slot.getBackgrounds()[0].every((background) => background == GREEN)) {
    slot.merge();
    //var centerCol = Math.round(slot.getNumColumns()/2);

    slot.setBackground(statusColor);
    slot.setValue(display);
    slot.setHorizontalAlignment('center');
    slot.setVerticalAlignment('middle');
    slot.setNote(note);
    //formatBorderFill(slot);
    return 1;
  } else {
    var ui = SpreadsheetApp.getUi();
    var conflict = ui.alert(
      'Time Conflict. Do you want to override?',
      ui.ButtonSet.YES_NO);
      
    if (conflict == ui.Button.YES) {
      conflictResolve(slot, display, statusColor, note);
      return 1;
    } else {
      return 0;
    }
  }
}

/*

// Function: Formats borders appropriately.
function formatBorderFill(slot) {
  var slotColBeg = slot.getColumn() - 2;
  var slotColEnd = slot.getLastColumn() + 1;
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = masterSheet.getSheetByName('Printer Schedule');
  var rightDate = scheduleSheet.getRange(6, slotColEnd, 1, 1).getValue();
  var leftDate = scheduleSheet.getRange(6, slotColBeg, 1, 2).getValue();

  if (rightDate == undefined || rightDate.getDay() !== 6) {
    slot.setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.DASHED);
    return;
  } else {
    slot.setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  if (leftDate == undefined || leftDate.getDay() !== 0) {
    slot.setBorder(null, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.DASHED);
    return;
  } else {
    slot.setBorder(null, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    return;
  }
}




function formatBorderRemove(slot) {
  var slotColBeg = slot.getColumn() - 2;
  var slotColEnd = slot.getLastColumn() + 1;
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleSheet = masterSheet.getSheetByName('Printer Schedule');
  var rightDate = scheduleSheet.getRange(6, slotColEnd, 1, 1).getValue();
  var leftDate = scheduleSheet.getRange(6, slotColBeg, 1, 2).getValue();

  if (rightDate == undefined || rightDate.getDay() !== 6) {
    slot.setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.DASHED);
    return;
  } else {
    slot.setBorder(null, null, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  if (leftDate == undefined || leftDate.getDay() !== 0) {
    slot.setBorder(null, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.DASHED);
    return;
  } else {
    slot.setBorder(null, true, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    return;
  }
}
*/

// Resolves Time Conflict
function conflictResolve(slot, travNum, statusColor, note) {
  var conflictCells = slot.getMergedRanges();
  var sheet = SpreadsheetApp.getActiveSheet();


  for (let i = 0; i < conflictCells.length; i++) {
    var tempValue = conflictCells[i].getValue();
    var tempCol = conflictCells[i].getLastColumn() - conflictCells[i].getColumn()+1; 
    conflictCells[i].breakApart();
    conflictCells[i].getCell(1, tempCol).setValue(tempValue);

/* FOR SOME REASON, GOOGLE SHEETS CANNOT FIND SHEET. COMMON ERROR WHEN DEALING WITH TOO MUCH DATA?
    if (slot.getColumn() > conflictCells[i].getColumn()) {
      var temp = sheet.getRange(conflictCells[i].getRow(), conflictCells[i].getColumn(), 1, (slot.getColumn()-1));
      temp.merge();
    } else {
      var temp = sheet.getRange(conflictCells[i].getRow(), conflictCells[i].getColumn(), 1, (slot.getLastColumn()+1));
      temp.merge();
    }
   */ 
  }


  slot.merge();
  slot.setBackground(statusColor);
  slot.setValue(travNum);
  slot.setHorizontalAlignment('center');
  slot.setVerticalAlignment('middle');
  slot.setNote(note);
  slot.setBorder(null, true, null, true, null, null, '#000000', SpreadsheetApp.BorderStyle.DASHED);
}




// Removes time slot from schedule.
function removeTime(sheet, machineRow, startDate, time, amPm) {
  if (machineRow == 0) {
    return;
  }
  var dateRow = sheet.getSheetValues(6,2,1,(sheet.getLastColumn()-2))[0];
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
  var shortSheet = masterSheet.getSheetByName('3 Month Schedule');

  // Resetting formatting

  // Rows above and below machine rows
  scheduleSheet.getRange(1,2,4,732).setBackground('#ffffff');
  scheduleSheet.getRange(1,2,4,732).setValue('');
  scheduleSheet.getRange(5,2,2,732).setBackground('#efefef');
  scheduleSheet.getRange(12,2,2,732).setBackground('#efefef');
  scheduleSheet.getRange(14,2,5,732).setBackground('#ffffff');
  shortSheet.getRange(1,2,4,126).setBackground('#ffffff');
  shortSheet.getRange(1,2,4,126).setValue('');
  shortSheet.getRange(5,2,2,126).setBackground('#efefef');
  shortSheet.getRange(12,2,2,126).setBackground('#efefef');
  shortSheet.getRange(14,2,5,126).setBackground('#ffffff');
  // Machine rows
  //var machineRows = scheduleSheet.getRange(7,2,5,732);
  //var machineRowsShort = shortSheet.getRange(7,2,5,126);

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
  var dateRow = scheduleSheet.getSheetValues(6,2,1,(scheduleSheet.getLastColumn()-2))[0];
  var dateRow2 = shortSheet.getSheetValues(6,2,1,(shortSheet.getLastColumn()-2))[0];
  var today = new Date();
  today.setHours(0,0,0,0);
  
  todayIndex = 0;
  for (let i = 0; i < dateRow.length; i++) {
    if (dateRow[i] != '') {
      if (String(today) == String(dateRow[i])) {
        todayIndex = i+2;
        break;
      }
    }
  }
  todayIndexShort = 0;
  for (let i = 0; i < dateRow2.length; i++) {
    if (dateRow2[i] != '') {
      if (String(today) == String(dateRow2[i])) {
        todayIndexShort = i+2;
        break;
      }
    }
  }

  scheduleSheet.getRange(1, todayIndex, 6, 2).setBackground('#d9b8df');
  scheduleSheet.getRange(1, todayIndex, 1, 1).setValue('Today');
  scheduleSheet.getRange(12, todayIndex, 7, 2).setBackground('#d9b8df');
  shortSheet.getRange(1, todayIndexShort, 6, 2).setBackground('#d9b8df');
  shortSheet.getRange(1, todayIndexShort, 1, 1).setValue('Today');
  shortSheet.getRange(12, todayIndexShort, 7, 2).setBackground('#d9b8df');
}




// Sends an item in Print Jobs or NCMR, ECO, DCO, DEV to the Event Log tab
function sendEventToLog(currentSheet, masterLogSheet, subLogSheet, rowData) {
  var sheetCategories = currentSheet.getRange(1, 1, 1, currentSheet.getLastColumn()).getValues()[0];
  sheetCategories.pop();
  sheetCategories.shift();
  console.log(rowData);

  var now = new Date();
  var timeStamp = now.toString();

  if (currentSheet.getName() == '1: Print Jobs' || currentSheet.getName() == '2: NCMR, ECO, DCO, DEV' || currentSheet.getName() == '3: Facility' || currentSheet.getName() == '4: PM') {
    var emptyRow = getFirstEmptyRowByColumnArray(subLogSheet);
    subLogSheet.getRange(emptyRow, 1, 1, 1).setValue(timeStamp);
    for (let i = 0; i < sheetCategories.length; i++) {
      if (sheetCategories[i] in rowData) {
        subLogSheet.getRange(emptyRow, i+2, 1, 1).setValue(rowData[sheetCategories[i]]);
      } else {
        subLogSheet.getRange(emptyRow, i+2, 1, 1).setValue('N/A');
      }
    }
  } else {
    console.log('Error, try to send an event from an invalid sheet.');
    return;
  }
  var masterCategories = masterLogSheet.getRange(1,1,1,masterLogSheet.getLastColumn()).getValues()[0];
  var emptyRow = getFirstEmptyRowByColumnArray(masterLogSheet);
  masterLogSheet.getRange(emptyRow, 1, 1, 1).setValue(timeStamp);
  for (let i = 1; i < masterCategories.length; i++) {
    if (masterCategories[i] in rowData) {
      masterLogSheet.getRange(emptyRow, i+1, 1, 1).setValue(rowData[masterCategories[i]]);
    } else {
      masterLogSheet.getRange(emptyRow, i+1, 1, 1).setValue('N/A');
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

