function onFormSubmit(e) {
  var masterSheet = SpreadsheetApp.getActiveSpreadsheet();
  console.log(e);
  console.log(e.namedValues);

  var shippingSheet = masterSheet.getSheetByName('Sheet1');
  fillShippingRow(shippingSheet, getFirstEmptyRowByColumnA(shippingSheet), e.namedValues);

}

// Gets the first empty row on a given sheet.
function getFirstEmptyRowByColumnA(sheet) {
  var row = 1;
  var isEmpty = 1;
  while (isEmpty) {
    console.log(row);
    if (sheet.getRange(row, 1, 1, sheet.getLastColumn()).isBlank()) {
      isEmpty = 0;
    } else {
      row += 1;
    }
  }
  return row;
}

function fillShippingRow(sheet, rowNum, rowData) {
  var orderNum = 1;
  var trackingNum = 3;
  var dateSent = 4;
  var sender = 12;
  var shipmentType = 20;
  var timeStamp = 1;
  var attachmentURL = 1;
  
  sheet.getRange(rowNum, orderNum, 1, 1).setValue(rowData['Order Number'][0]);
  sheet.getRange(rowNum, trackingNum, 1, 1).setValue(rowData['Tracking Number'][0]);
  sheet.getRange(rowNum, dateSent, 1, 1).setValue(rowData['Date Sent'][0]);
  sheet.getRange(rowNum, sender, 1, 1).setValue(rowData['Sender'][0]);
  sheet.getRange(rowNum, shipmentType, 1, 1).setValue(rowData['Shipment Type'][0]);
}