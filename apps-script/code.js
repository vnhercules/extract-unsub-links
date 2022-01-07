/**
* Extract Gmail Message Metadata to Google Sheets
* By Vanesa Hercules, 2022
*/

/**
* Custom menu
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Extract Gmail Messages")
    .addItem("Extract by label", "getEmailsByLabel")
    .addItem("Extract by search", "getEmailsBySearch")
    .addToUi();
}

/**
* Get emails by label
*/
function getEmailsByLabel(){
  var label = GmailApp.getUserLabelByName("Data Interview Qs"); // update to label of interest
  var threads = label.getThreads();
  getMsgDetails(threads);
}

/**
* Get emails by query
*/
function getEmailsBySearch(){
  var threads = GmailApp.search("is:Trash is:Unread", 0, 50); // update to query of interest and limit to 50 threads
  getMsgDetails(threads);
}

/**
* Loop through emails and extract metadata
*/
function getMsgDetails(threads){
  for(var i = threads.length - 1; i >= 0; i--){
    var messages = threads[i].getMessages();

    for(var j = 0; j < messages.length; j++){
      var message = messages[j]
      extractDetails(message);
    }
  }
}

/**
* Extract metadata and write to Google Sheets
*/
function extractDetails(message){
  var dateTime = message.getDate();
  var subjectText = message.getSubject();
  var senderDetails = message.getFrom();
  var msgBody = message.getBody();

  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  activeSheet.appendRow([dateTime, senderDetails, subjectText, msgBody])
}
