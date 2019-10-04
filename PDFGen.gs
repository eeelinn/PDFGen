function makePDF(){
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PDF").activate();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();  // Get main Database Sheet
  var values = sheet.getRange("D2:HV2").getValues()[0];   // Get Values
  var keys = sheet.getRange("D1:HV1").getValues()[0];    // Get Titles
  var template = "https://docs.google.com/document/d/1HuDSjkJ5iJciXs0U7legGE1TFeDfeH0S27r3j7gbJ7A/edit#"
  
  var newFile = fillTemplate(keys, values, template);
  
  //Get the following text from the link after 'https://drive.google.com/drive/u/1/folders/' 
  //For example this is from 'https://drive.google.com/drive/u/1/folders/1uMhtQ1nHqp3gcvMbvKuEPghX9NBkZLsN'
  var TARGET_FOLDER = '1uMhtQ1nHqp3gcvMbvKuEPghX9NBkZLsN' 
  var targetFolder = DriveApp.getFolderById(TARGET_FOLDER);
  targetFolder.addFile(newFile);
  
}

//Fills template areas with values from sheet
function fillTemplate(keys, array, template) {
  var doc = DocumentApp.openByUrl(template);
  var copyFile = DriveApp.getFileById(doc.getId()).makeCopy("copy"),
      copy = DocumentApp.openById(copyFile.getId()),
      copybody = copy.getBody();
  
  for(var i=0; i < keys.length; i++) {
    if (keys[i] == 'Full Name') {
      copy.setName('VAR: ' + array[i]); 
    }
    copybody.replaceText('{{'+keys[i]+'}}', array[i]); 
  }
  
  copy.saveAndClose();
  
  var newFile = DriveApp.createFile(copy.getBlob().getAs('application/pdf'));
  copyFile.setTrashed(true);
  
  return newFile
}
