// This Google Sheets script will post to a slack channel when a user submits data to a Google Forms Spreadsheet


// Variables
var slackIncomingWebhookUrl = 'https://hooks.slack.com/services/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';

// Include # for public channels, omit it for private channels
var postChannel = "Cat";

var postIcon = ":mailbox_with_mail:";
var postUser = "Form Response";
var postColor = "#FFFFFF";

var messageFallback = "The attachment must be viewed as plain text.";
var messagePretext = "RUFUS has submitted something!!!!!!";

var shippingSheetId = "1UUr_hemxmOCmD3Vj_EMFkALYWOFZOGTaTS36sAN7rRU";

///////////////////////
// End customization //
///////////////////////






// In the Script Editor, run initialize() at least once to make your code execute on form submit
function initialize() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i in triggers) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger("submitValuesToSlack")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
}






// Running the code in initialize() will cause this function to be triggered this on every Form Submit
function submitValuesToSlack(e) {
  // Test code. uncomment to debug in Google Script editor
  // if (typeof e === "undefined") {
  //   e = {namedValues: {"Question1": ["answer1"], "Question2" : ["answer2"]}};
  //   messagePretext = "Debugging our Sheets to Slack integration";
  // }


  // Fill in Shipping sheet
  var file = DriveApp.getFileById(shippingSheetId);
  var masterSheet = SpreadsheetApp.open(file);
  var shippingSheet = masterSheet.getSheetByName('Data Log');
  console.log(shippingSheet.getName());
  fillShippingRow(shippingSheet, getFirstEmptyRowByColumnA(shippingSheet), e.namedValues);


  var attachments = constructAttachments(e.values);

  var payload = {
    "channel": postChannel,
    "username": postUser,
    "icon_emoji": postIcon,
    "link_names": 1,
    "attachments": attachments
  };

  var options = {
    'method': 'post',
    'payload': JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(slackIncomingWebhookUrl, options);
}






// Creates Slack message attachments which contain the data from the Google Form
// submission, which is passed in as a parameter
// https://api.slack.com/docs/message-attachments
var constructAttachments = function(values) {
  var fields = makeFields(values);

  var attachments = [{
    "fallback" : messageFallback,
    "pretext" : messagePretext,
    "mrkdwn_in" : ["pretext"],
    "color" : postColor,
    "fields" : fields
  }]

  return attachments;
}






// Creates an array of Slack fields containing the questions and answers
var makeFields = function(values) {
  var fields = [];

  var columnNames = getColumnNames();

  for (var i = 0; i < columnNames.length; i++) {
    var colName = columnNames[i];
    var val = values[i];
    fields.push(makeField(colName, val));
  }

  return fields;
}






// Creates a Slack field for your message
// https://api.slack.com/docs/message-attachments#fields
var makeField = function(question, answer) {
  var field = {
    "title" : question,
    "value" : answer,
    "short" : false
  };
  return field;
}






// Extracts the column names from the first row of the spreadsheet
var getColumnNames = function() {
  var sheet = SpreadsheetApp.getActiveSheet();

  // Get the header row using A1 notation
  var headerRow = sheet.getRange("1:1");

  // Extract the values from it
  var headerRowValues = headerRow.getValues()[0];

  return headerRowValues;
}






// Gets the first empty row on a given sheet.
function getFirstEmptyRowByColumnA(sheet) {
  var row = 1;
  var isEmpty = 1;
  while (isEmpty) {
    console.log(row);
    console.log(sheet.getName());
    if (sheet.getRange(row, 1, 1, sheet.getLastColumn()).isBlank()) {
      isEmpty = 0;
    } else {
      row += 1;
    }
  }
  return row;
}






// Fills in the custom Shipping sheet
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