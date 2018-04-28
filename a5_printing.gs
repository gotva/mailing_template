function onOpen() {
  var menuEntries = [ {name: "A5 Envelope", functionName: "printA5Envelope"}];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("Print", menuEntries);
}

function printA5Envelope() {
  var SENDER_NAME = 'Vasily';
  var SENDER_ADDRESS = 'My Home, Sweet Home';
  
  var templateid = "1DQk7G-eNgRMz2PU-Wo5hpkv3qtURjS8Zss6GmkH72l8"; // text with %str%
  var templateDoc = DocumentApp.openById(templateid);
  var templateDocBody = templateDoc.getBody();
  
  var emptyTemplateid = "1oOOj2FlaqquncZ-3h8dL-SPwWR1X-h1uv2V5RNrLiX0"; // empty file with correct borders
  var outFolder = DriveApp.getFolderById("1MCVahL7ix9NKqwWEvQSSni9uzdaIJ8TV");
  var newMailingFile = DriveApp.getFileById(emptyTemplateid);
  var newMailingDocId = newMailingFile.makeCopy("Mailing_2018-03-18", outFolder).getId();
  var newMailingDoc = DocumentApp.openById(newMailingDocId);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues();
  
  for (var i in data){
    var replacedBody = templateDocBody.copy();
    var row = data[i];
    
    replacedBody.replaceText('%sender_name%', SENDER_NAME)
    replacedBody.replaceText('%sender_address%', SENDER_ADDRESS)
    replacedBody.replaceText("%recepient_name%", row[0]);
    replacedBody.replaceText("%recepient_address%", row[1]);
    
    appendToDoc(replacedBody, newMailingDoc)
    newMailingDoc.getBody().appendPageBreak();
  }
  
  newMailingDoc.getBody().removeChild(newMailingDoc.getBody().getChild(0)); // remove the first empty line
  
  templateDoc.saveAndClose();
  newMailingDoc.saveAndClose();
 
}

function appendToDoc(fromBody, toDoc) {
  var body = toDoc.getBody();
  var totalElements = fromBody.getNumChildren();
  
  for( var j = 0; j < totalElements; ++j ) {
    var element = fromBody.getChild(j).copy();
    var type = element.getType();
    if( type == DocumentApp.ElementType.PARAGRAPH )
      body.appendParagraph(element);
    else if( type == DocumentApp.ElementType.TABLE )
      body.appendTable(element);
    else if( type == DocumentApp.ElementType.LIST_ITEM )
      body.appendListItem(element);
    else
      throw new Error("According to the doc this type couldn't appear in the body: "+type);
  }
}