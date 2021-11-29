function onOpen(){
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Create Files');
  menu.addItem('Create New Event Folder', 'createEventFolder')
      .addItem('Create New Event Master Sheet', 'createMasterSheet')
      .addItem('Create New Slides', 'createSlides')
      .addItem('Create New Networking Sheet', 'createNetworkingSheet')
      .addItem('Create Q&A Doc', 'createQAdoc')
      .addItem('Create New Speaker Package Doc', 'createSpeakerPackage')
      .addItem('Create New Partner Package Doc', 'createPartnerPackage')
      .addItem('Create New Event Description Doc', 'createEventDescription')
      .addToUi();
}

function createEventFolder() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  const rows = sheet.getDataRange().getValues();
  var column = rows[0].indexOf("folder_link");

  rows.forEach(function(row, index) {
    if (index === 0) return;
    if (row[column]) return;
    var primaryFolderID = config.getRange(1,2).getValues();
    var mainFolder = DriveApp.getFolderById(primaryFolderID);
    var newFolder = mainFolder.createFolder(`${row[3]} ${row[5]} ${row[7]}`);
    var newFolderID = newFolder.getId();
    var url = newFolder.getUrl();
    
    sheet.getRange(index + 1, column + 1).setValue(url)
    sheet.getRange(index + 1, column + 2).setValue(newFolderID)

   

  })
}
  

function createMasterSheet(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  const rows = sheet.getDataRange().getValues();
  var column = rows[0].indexOf("event_master_link");
  var destinationFolderColumn = rows[0].indexOf("folder_id");
  
  rows.forEach(function(row, index) {
    if (index === 0) return;
    if (row[column]) return;
    var masterTemplateID = config.getRange(2,2).getValues();;
    var destinationFolder = DriveApp.getFolderById(`${row[destinationFolderColumn]}`);

    //Create a copy of event master sheets template
    var masterTemplate = DriveApp.getFileById(masterTemplateID);
    var masterCopy = masterTemplate.makeCopy(`HirePhD Event Master Document ${row[3]} ${row[8]}`, destinationFolder);
    var newMasterID = masterCopy.getId();
    var masterUrl = masterCopy.getUrl();
    
    sheet.getRange(index + 1, column + 1).setValue(masterUrl)
    sheet.getRange(index + 1, column + 2).setValue(newMasterID)

    var newFile = SpreadsheetApp.openById(newMasterID); //add "Series" and "Topic" info into new master sheet
    var newFileSheet = newFile.getSheetByName("Speaker Info & Agenda");
    var series = newFileSheet.getRange("E10");
    series.setValue(`${row[8]}`); //set value of cell
    var topic = newFileSheet.getRange("E11");
    topic.setValue(`${row[7]}`); //set value of cell

  })

}

function createSlides() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  const rows = sheet.getDataRange().getValues();
  var column = rows[0].indexOf("slides_link");
  
  rows.forEach(function(row, index) {
    if (index === 0) return;
    if (row[column]) return;
    var slidesTemplateID = config.getRange(3,2).getValues();;
    var destinationFolder = DriveApp.getFolderById(`${row[11]}`);

    //Create a copy of slides template
    var slidesTemplate = DriveApp.getFileById(slidesTemplateID);
    var slidesCopy = slidesTemplate.makeCopy(`Slides ${row[3]} ${row[8]}`, destinationFolder);
    var newSlidesID = slidesCopy.getId();


    var slidesUrl = slidesCopy.getUrl();
    
    sheet.getRange(index + 1, column + 1).setValue(slidesUrl)
    sheet.getRange(index + 1, column + 2).setValue(newSlidesID)
  })
}


function createNetworkingSheet(){

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  const rows = sheet.getDataRange().getValues();
  var column = rows[0].indexOf("networking_link");
  
  rows.forEach(function(row, index) {
    if (index === 0) return;
    if (row[column]) return;
    var networkingTemplateID = config.getRange(4,2).getValues();
    var destinationFolder = DriveApp.getFolderById(`${row[10]}`);

    //Create a copy of networking sheets template
    var networkingTemplate = DriveApp.getFileById(networkingTemplateID);
    var networkingCopy = networkingTemplate.makeCopy(`HirePhD Networking Document ${row[3]} ${row[8]}`, destinationFolder);
    var newNetworkingID = networkingCopy.getId();


    var networkingUrl = networkingCopy.getUrl();
    
    sheet.getRange(index + 1, column + 1).setValue(networkingUrl)
    sheet.getRange(index + 1, column + 2).setValue(newNetworkingID)
  })

}

function createQAdoc(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  const rows = sheet.getDataRange().getValues();
  var column = rows[0].indexOf("questions_link")
  
  rows.forEach(function(row, index) {
    if (index === 0) return;
    if (row[column]) return;
    var questionTemplateID = config.getRange(5,2).getValues();;
    var destinationFolder = DriveApp.getFolderById(`${row[10]}`);

    //Create a copy of networking sheets template
    var questionTemplate = DriveApp.getFileById(questionTemplateID);
    var questionCopy = questionTemplate.makeCopy(`Q&A Document ${row[3]} ${row[8]}`, destinationFolder);
    var newQuestionID = questionCopy.getId();


    var questionUrl = questionCopy.getUrl();
    
    sheet.getRange(index + 1, column + 1).setValue(questionUrl)
    sheet.getRange(index + 1, column + 2).setValue(newQuestionID)
  })

}

function createSpeakerPackage(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  const rows = sheet.getDataRange().getValues();
  var column = rows[0].indexOf("speaker_package_link");
  
  rows.forEach(function(row, index) {
    if (index === 0) return;
    if (row[column]) return;
    var speakerTemplateID = config.getRange(7,2).getValues();;
    var destinationFolder = DriveApp.getFolderById(`${row[10]}`);

    //Create a copy of networking sheets template
    var speakerTemplate = DriveApp.getFileById(speakerTemplateID);
    var speakerCopy = speakerTemplate.makeCopy(`Speaker Package ${row[3]} ${row[8]}`, destinationFolder);
    var newSpeakerID = speakerCopy.getId();


    var speakerUrl = speakerCopy.getUrl();
    
    sheet.getRange(index + 1, column + 1).setValue(speakerUrl)
    sheet.getRange(index + 1, column + 2).setValue(newSpeakerID)
  })

}

function createPartnerPackage(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  const rows = sheet.getDataRange().getValues();
  var column = rows[0].indexOf("partner_package_link");
  
  rows.forEach(function(row, index) {
    if (index === 0) return;
    if (row[column]) return;
    var partnerTemplateID = config.getRange(8,2).getValues();;
    var destinationFolder = DriveApp.getFolderById(`${row[10]}`);

    //Create a copy of networking sheets template
    var partnerTemplate = DriveApp.getFileById(partnerTemplateID);
    var partnerCopy = partnerTemplate.makeCopy(`Partner Package ${row[3]} ${row[8]}`, destinationFolder);
    var newPartnerID = partnerCopy.getId();


    var partnerUrl = partnerCopy.getUrl();
    
    sheet.getRange(index + 1, column + 1).setValue(partnerUrl)
    sheet.getRange(index + 1, column + 2).setValue(newPartnerID)
  })

}

function createEventDescription(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config');
  const rows = sheet.getDataRange().getValues();
  var column = rows[0].indexOf("event_description_link");
  
  rows.forEach(function(row, index) {
    if (index === 0) return;
    if (row[column]) return;
    var eventDescriptionTemplateID = config.getRange(6,2).getValues();;
    var destinationFolder = DriveApp.getFolderById(`${row[10]}`);

    //Create a copy of networking sheets template
    var eventDescriptionTemplate = DriveApp.getFileById(eventDescriptionTemplateID);
    var eventDescriptionCopy = eventDescriptionTemplate.makeCopy(`Event Description ${row[3]} ${row[8]}`, destinationFolder);
    var newEventDescriptionID = eventDescriptionCopy.getId();


    var eventDescriptionUrl = eventDescriptionCopy.getUrl();
    
    sheet.getRange(index + 1, column + 1).setValue(eventDescriptionUrl)
    sheet.getRange(index + 1, column + 2).setValue(newEventDescriptionID)
  })

}
