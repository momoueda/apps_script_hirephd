function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
      .addSubMenu(ui.createMenu('Main')
        .addItem('Add Event Plan Master Sheet Link', 'addEventMasterLink')
        .addItem('Add Event Folder Link', 'addFolderLink')
        .addItem('Create New Slides', 'createSlides')
        )
      .addToUi();
}

function addEventMasterLink(){
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main');
    const url = SpreadsheetApp.getActiveSpreadsheet().getUrl();

    sheet.getRange(11, 5).setValue(url)
}

function addFolderLink(){
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main');
    var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    var sheetFile = DriveApp.getFileById(sheetId);
    var folderUrl = sheetFile.getParents().next().getUrl();


    sheet.getRange(14, 5).setValue(folderUrl);

}

function createSlides() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config (leave untouched)');
  const auto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Automation (leave untouched)');
  
  //get folderId
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var sheetFile = DriveApp.getFileById(sheetId);
  var folderId = sheetFile.getParents().next().getId();

  var slidesTemplateID = config.getRange(1,2).getValues();;
  var destinationFolder = DriveApp.getFolderById(folderId);

  //get values for date, series, speaker and topic
  var date = auto.getRange("B1").getValues().toString();
  var series = auto.getRange("B2").getValues().toString();
  var speaker = auto.getRange("B3").getValues().toString();
  var topic = auto.getRange("B4").getValues().toString();

    //Create a copy of slides template
  var slidesTemplate = DriveApp.getFileById(slidesTemplateID);
  var slidesCopy = slidesTemplate.makeCopy(date + `_` + series + `_` + speaker + `_` + topic + `_` + `Slides`, destinationFolder);
  var slidesUrl = slidesCopy.getUrl();
    
  sheet.getRange(25, 5).setValue(slidesUrl);

}