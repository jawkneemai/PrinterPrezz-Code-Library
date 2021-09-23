// Traverses through all travelers in Printing Projects (RT). 
// Currently made to grab parts lists from google slides (T4210 and before), parts lists from traveler folder PL (T4283 and above)
 

function getPartsLists() {
  // Iterates through Run Traveler folders older than T4211 to get parts lists, stored in a google slides file.

  var masterPrintFolder = DriveApp.getFolderById('1lNxvM7WIKcGfM-XR-nAbP6nT0NtYPC7T');
  var t4200To4100Folder = DriveApp.getFolderById('1YawCmPWxYXrTRCQLcz2yUbgVANpo2zH_');
  var t4100UnderFolder = DriveApp.getFolderById('1M_GWj2-4P5f65lDj6u4kkKiEpR2ioCEd');
  var slidesFolder = DriveApp.getFolderById('11M8531hNPTQOs5rLzE0bhe5PsbhCZZsk');
  var dhf = DriveApp.getFileById('1Dzi_BjU2MsJ507MYx69Zni3fjmIoNWh7uqjI3N4FNgI');
  dhf = SpreadsheetApp.open(dhf);
  var dhfMasterList = dhf.getSheetByName('Master List');
  var travelerNames = [];
  var travelerFolders = [];

  
  var folderIterator = masterPrintFolder.getFolders();
  while (folderIterator.hasNext()) {
    var temp = folderIterator.next();
    travelerFolders.push(temp);
    travelerNames.push(temp.getName());
  }
  

  /*
  folderIterator = t4200To4100Folder.getFolders();
  while (folderIterator.hasNext()) {
    temp = folderIterator.next();
    travelerNames.push(temp.getName());
  }

  folderIterator = t4100UnderFolder.getFolders();
  while (folderIterator.hasNext()) {
    temp = folderIterator.next();
    travelerNames.push(temp.getName());
  }

  
  for (var i = 0; i < travelerNames.length; i++) {
    dhfMasterList.getRange(i+2, 2, 1, 1).setValue(travelerNames[i]);
  }
  */

  // Gets all travelers in a master folder (i.e. T4100-T4200) their PPT of their build arrangement (with PL) and converts them to google slides
  /*
  var folderIterator = masterPrintFolder.getFolders();
  while(folderIterator.hasNext()) {
    temp = folderIterator.next();
    var tempName = temp.getName();
    var tempId = temp.getId();
    var tempChildrenFolders = temp.getFolders();
    var tempChildrenFiles = temp.getFiles();
    var tempPL = [];
    while (tempChildrenFolders.hasNext()) {
      var temp2 = tempChildrenFolders.next();
      if (temp2.getName() == 'Run Traveler') {
        var deeperFiles = temp2.getFiles();
        while (deeperFiles.hasNext()) {
          var temp3 = deeperFiles.next();
          if (temp3.getMimeType() == 'application/vnd.openxmlformats-officedocument.presentationml.presentation') {
            tempPL.push(temp3);
          }
        }
      }
    }
    while (tempChildrenFiles.hasNext()) {
      temp2 = tempChildrenFiles.next();
      if (temp2.getMimeType() == 'application/vnd.openxmlformats-officedocument.presentationml.presentation') {
        tempPL.push(temp2);
      }
    }

    console.log(tempName);
    for (i=0; i < tempPL.length; i++) {
      try {
        Drive.Files.insert({title: tempPL[i].getName(), mimeType: MimeType.GOOGLE_SLIDES, parents: [{id: '11M8531hNPTQOs5rLzE0bhe5PsbhCZZsk'}]}, tempPL[i].getBlob());
        console.log('converted ' + tempPL[i].getName());
      } catch (e) {
        console.log('Error: ' + e);
      }
    }
  }
  
  */

  // Gather all PPT names
  /*
  var slides = slidesFolder.getFiles();
  var slideNames = [];
  while (slides.hasNext()) {
    var temp = slides.next();
    console.log(temp.getName());
    slideNames.push(temp.getName());
  }

  for (var i = 0; i < slideNames.length; i++) {
    dhfMasterList.getRange(i+2, 4, 1, 1).setValue(slideNames[i]);
  }
  */

  // Pull information from Slides. Looks to see if there's a table element, copies table element into Master Parts List in a new tab
  /*
  var slides = slidesFolder.getFiles();
  var dhfSheets = dhf.getSheets();
  var sheetsArray = [];
  for (var m = 0; m < dhfSheets.length; m++) {
    sheetsArray.push(dhfSheets[m].getName());
  }

  while (slides.hasNext()) {
    var temp = slides.next();
    temp = SlidesApp.openById(temp.getId());
    var tempSlides = temp.getSlides();
    var traveler = temp.getName().split('_')[0];
    tempSlides.shift();

    // Creates or gets sheet for traveler
    console.log('Searching ' + traveler + '...');
    if (sheetsArray.includes(traveler)) {
      var newSheet = dhf.getSheetByName(traveler);
    } else {
      var newSheet = dhf.insertSheet(traveler);
      sheetsArray.push(traveler);
    }

    // Searches slide for tables
    for (var i = 0; i < tempSlides.length; i++) {

      // Gets all tables in traveler slide deck
      var tempTables = tempSlides[i].getTables();
      if (tempTables.length > 0) {
        for (var j = 0; j < tempTables.length; j++) {
          var numRows = tempTables[j].getNumRows();
          var numCols = tempTables[j].getNumColumns();
          
          // Inserting row by row for every table
          for (var k = 0; k < numRows; k++) {
            var emptyRow = getFirstEmptyRowByColumnArray(newSheet);
            for (var l = 0; l < numCols; l++) {
              if (tempTables[j].getCell(k, l)) {
                //console.log(tempTables[j].getCell(k, l).getText().asString());
                newSheet.getRange(k+emptyRow, l+1, 1, 1).setValue(tempTables[j].getCell(k, l).getText().asString());
              }
            }
          }
          newSheet.getRange(1, 1, newSheet.getLastRow(), newSheet.getLastColumn()).setHorizontalAlignment('left');
          newSheet.getRange(1, 1, newSheet.getLastRow(), newSheet.getLastColumn()).setVerticalAlignment('top');
        }
      }
    }

  console.log('Finished ' + traveler);
  }

  */

  /*
  // Changing T4211 - T4282 Sheet Names
  var trav2sheet = DriveApp.getFileById('169QARqsVqScquHefcLv5rtnKXEfNyIxsCvKLKrUQMIA');
  trav2sheet = SpreadsheetApp.open(trav2sheet);
  var tabNames = trav2sheet.getSheets();
  var tabArray = [];
  for (var i = 0; i < tabNames.length; i++) {
    var tempSheet = trav2sheet.getSheetByName(tabNames[i].getName());
    var tempName = tempSheet.getName().split('_')[0];
    console.log(tempName);
    tempSheet.setName(tempName);


  }
  */

  
  // Finally, travelers T4283 and above
  var travelerFolders = masterPrintFolder.getFolders();
  var t4283cutOffDate = new Date('July 22, 2021 00:00:00');
  var trav2Folders = [];
  var dhfSheets = dhf.getSheets();
  var sheetsArray = [];
  for (var m = 0; m < dhfSheets.length; m++) {
    sheetsArray.push(dhfSheets[m].getName());
  }

  // Time filter T4283 folders +
  while (travelerFolders.hasNext()) {
    var tempFolder = travelerFolders.next();
    var dateCreated = tempFolder.getDateCreated();
    if (dateCreated > t4283cutOffDate) {
      trav2Folders.push(tempFolder);
    }
  }
  
  for (var i = 0; i < trav2Folders.length; i++) {
    var tempName = trav2Folders[i].getName().split(' ')[0];
    if (sheetsArray.includes(tempName)) {
      var newSheet = dhf.getSheetByName(tempName);
    } else {
      var newSheet = dhf.insertSheet(tempName);
      sheetsArray.push(tempName);
    }
    console.log(tempName);
    var travelerFormsFolder = trav2Folders[i].getFoldersByName('Traveler Forms');
    if (travelerFormsFolder.hasNext()) {
      travelerFormsFolder = travelerFormsFolder.next();
      var children = travelerFormsFolder.getFiles();
      while (children.hasNext()) {
        var tempFile = children.next();
        if (tempFile.getName().includes('PL') || tempFile.getName().includes('Parts List')) {
          if (tempFile.getMimeType() == 'application/vnd.google-apps.spreadsheet') {
            var tempSheet = SpreadsheetApp.open(tempFile);
            var tempData = tempSheet.getActiveSheet().getRange(1, 1, tempSheet.getLastRow(), tempSheet.getLastColumn()).getValues();
            for (var j = 0; j < tempData.length; j++) {
              newSheet.getRange(getFirstEmptyRowByColumnArray(newSheet), 1, 1, tempData[j].length).setValues([tempData[j]]);
              console.log(tempData[j]);
            }
          }
        }
      }
    }
  }

  

  // Remove all blank cells, reached upper limit of cells in one sheet (5,000,000)
  var dhfSheets = dhf.getSheets();
  var sheetsArray = [];
  for (var m = 0; m < dhfSheets.length; m++) {
    sheetsArray.push(dhfSheets[m].getName());
  }
  for (var i = 1; i < dhfSheets.length; i++) {
    var tempSheet = dhfSheets[i];
    console.log(tempSheet.getName());
    tempSheet.deleteRows(tempSheet.getLastRow()+1, 1000-(tempSheet.getLastRow()+1));
    tempSheet.deleteColumns(tempSheet.getLastColumn()+1, 26-(tempSheet.getLastColumn()+1));
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

























