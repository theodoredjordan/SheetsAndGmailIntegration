function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Ted's Menu").addItem('Get Emails', 'getGmailEmails').addToUi();
}

function getGmailEmails() {
  var label = GmailApp.getUserLabelByName('tobeprocessedbygas');
  var threads = label.getThreads();

  for (var i = threads.length -1; i >=0; i--) {
    var messages = threads[i].getMessages();
  for (var j=0; j<messages.length; j++) {
    var message = messages[j];
    extractDetails(message);
    GmailApp.markMessageRead(message);
  }
   threads[i].removeLabel(label);
  }

}

function extractDetails(message) {
  var dateTime = message.getDate();
  var subjectText = message.getSubject();
  var senderDetails = message.getFrom();
  var bodyContents = message.getPlainBody();

  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  activeSheet.appendRow([dateTime, senderDetails, subjectText, bodyContents])
}
