function appendToExistingSpreadsheetWithAttachments(fileData) {

  var spreadsheetId = ""; //ID OF EXISTING SPREADSHEETS
  var existingSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var currentDate = new Date().toLocaleDateString();
  var sheet = existingSpreadsheet.insertSheet(currentDate);
  
  var lastRow = sheet.getLastRow();
    
  for (var i = 0; i < fileData.length; i++) {
    var file = DriveApp.createFile(fileData[i].name, fileData[i].content);
    var fileUrl = file.getUrl();
    sheet.getRange("A" + (lastRow + i + 1)).setValue(fileUrl);
    sheet.getRange("B" + (lastRow + i + 1)).setValue(fileData[i].source_ip);
    sheet.getRange("C" + (lastRow + i + 1)).setValue(fileData[i].spf);
    sheet.getRange("D" + (lastRow + i + 1)).setValue(fileData[i].dkim);
  }
  
  Logger.log("Data appended to existing spreadsheet: " + existingSpreadsheet.getUrl());
}


function loadAddOn(event) {

    var threads = GmailApp.getInboxThreads(0, 10);
    var msgs = GmailApp.getMessagesForThreads(threads);
    var spreadsheetId = ""; //ID OF EXISTING SPREADSHEETS
    var existingSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var currentDate = new Date().toLocaleDateString();
    var sheet = existingSpreadsheet.insertSheet(currentDate);
    var fileData = [];
    for (var i = 0 ; i < msgs.length; i++) {
      for (var j = 0; j < msgs[i].length; j++) {
        var attachments = msgs[i][j].getAttachments();
        for (var k = 0; k < attachments.length; k++) {
            s = Utilities.ungzip(attachments[k].copyBlob()).getDataAsString();
            var file = DriveApp.createFile(dict['name'], dict['content']);
            const array = s.split(" ");
            for(var l = 0; l < array.length; l++) {
              if(array[l].includes("source_ip")) {
                  const copy = JSON.parse(JSON.stringify(dict));
                  fileData.push(copy);
                  var sheet = existingSpreadsheet.getActiveSheet();
                  var lastRow = sheet.getLastRow();
                  var fileUrl = file.getUrl();
                  sheet.getRange("A" + (lastRow + 1)).setValue(fileUrl);
                  sheet.getRange("B" + (lastRow + 1)).setValue(dict['source_ip']);
                  sheet.getRange("C" + (lastRow + 1)).setValue(dict['spf']);
                  sheet.getRange("D" + (lastRow + 1)).setValue(dict['dkim']);
                }
                start = 0;
              }
              const regex = /<source_ip>(.*?)<\/source_ip>/;
              const match = regex.exec(array[l]);
              if (match !== null) {
                const result = match[1]; 
                dict['source_ip'] = result;
              }
          }
      }
  }
  appendToExistingSpreadsheetWithAttachments(fileData);
}
