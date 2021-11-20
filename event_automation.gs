function onOpen(){
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Create Files');
  menu.addItem('Create New Event Folder', 'createEventFolder')
      .addItem('Create New Slides', 'createSlides')
      .addItem('Create New Networking Sheet', 'createNetworkingSheet')
      .addItem('Create Q&A Doc', 'createQAdoc')
      .addToUi();
}

function createEventFolder() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  const rows = sheet.getDataRange().getValues();

  rows.forEach(function(row, index) {
    if (index === 0) return;
    if (row[9]) return;
    var primaryFolderID = config.getRange(1,2).getValues();
    var mainFolder = DriveApp.getFolderById(primaryFolderID);
    var newFolder = mainFolder.createFolder(`${row[3]} ${row[5]} ${row[7]}`);
    var newFolderID = newFolder.getId();
    var url = newFolder.getUrl();
    
    sheet.getRange(index + 1, 10).setValue(url)
    sheet.getRange(index + 1, 11).setValue(newFolderID)
  })
}

function createSlides() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  const rows = sheet.getDataRange().getValues();
  
  rows.forEach(function(row, index) {
    if (index === 0) return;
    if (row[11]) return;
    var slidesTemplateID = config.getRange(3,2).getValues();;
    var destinationFolder = DriveApp.getFolderById(`${row[10]}`);

    //Create a copy of slides template
    var slidesTemplate = DriveApp.getFileById(slidesTemplateID);
    var slidesCopy = slidesTemplate.makeCopy(`Slides for ${row[5]}'s event`, destinationFolder);
    var newSlidesID = slidesCopy.getId();


    var slidesUrl = slidesCopy.getUrl();
    
    sheet.getRange(index + 1, 12).setValue(slidesUrl)
    sheet.getRange(index + 1, 13).setValue(newSlidesID)
  })
}

function createMasterSheet(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  const rows = sheet.getDataRange().getValues();
  
  rows.forEach(function(row, index) {
    if (index === 0) return;
    if (row[11]) return;
    var masterTemplateID = config.getRange(2,2).getValues();;
    var destinationFolder = DriveApp.getFolderById(`${row[10]}`);

    //Create a copy of event master sheets template
    var masterTemplate = DriveApp.getFileById(masterTemplateID);
    var masterCopy = masterTemplate.makeCopy(`HirePhD Event Master Document ${row[3]}`, destinationFolder);
    var newMasterID = masterCopy.getId();

    var masterUrl = masterCopy.getUrl();
    
    sheet.getRange(index + 1, 12).setValue(masterUrl)
    sheet.getRange(index + 1, 13).setValue(newMasterID)
  })

}

function createNetworkingSheet(){

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  const rows = sheet.getDataRange().getValues();
  
  rows.forEach(function(row, index) {
    if (index === 0) return;
    if (row[15]) return;
    var networkingTemplateID = config.getRange(4,2).getValues();
    var destinationFolder = DriveApp.getFolderById(`${row[10]}`);

    //Create a copy of networking sheets template
    var networkingTemplate = DriveApp.getFileById(networkingTemplateID);
    var networkingCopy = networkingTemplate.makeCopy(`HirePhD Networking Document ${row[3]}`, destinationFolder);
    var newNetworkingID = networkingCopy.getId();


    var networkingUrl = networkingCopy.getUrl();
    
    sheet.getRange(index + 1, 16).setValue(networkingUrl)
    sheet.getRange(index + 1, 17).setValue(newNetworkingID)
  })

}

function createQAdoc(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  const rows = sheet.getDataRange().getValues();
  
  rows.forEach(function(row, index) {
    if (index === 0) return;
    if (row[17]) return;
    var questionTemplateID = config.getRange(5,2).getValues();;
    var destinationFolder = DriveApp.getFolderById(`${row[10]}`);

    //Create a copy of networking sheets template
    var questionTemplate = DriveApp.getFileById(questionTemplateID);
    var questionCopy = questionTemplate.makeCopy(`Q&A Document for ${row[5]}'s event`, destinationFolder);
    var newQuestionID = questionCopy.getId();


    var questionUrl = questionCopy.getUrl();
    
    sheet.getRange(index + 1, 18).setValue(questionUrl)
    sheet.getRange(index + 1, 19).setValue(newQuestionID)
  })

}
7