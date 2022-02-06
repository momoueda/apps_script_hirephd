function onOpen(){
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Create Files');
  menu.addItem('Create New Event Folder', 'createEventFolder')
      .addItem('Create New Event Master Sheet', 'createMasterSheet')
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
    var masterCopy = masterTemplate.makeCopy(`Event_Master_Document_${row[3]}_${row[8]}_${row[5]}`, destinationFolder);
    var newMasterID = masterCopy.getId();
    var masterUrl = masterCopy.getUrl();
    
    sheet.getRange(index + 1, column + 1).setValue(masterUrl)
    sheet.getRange(index + 1, column + 2).setValue(newMasterID)

    var newFile = SpreadsheetApp.openById(newMasterID); //add "Series" and "Topic" info into new master sheet
    var newFileSheet = newFile.getSheetByName("Automation (leave untouched)");
    var date = newFileSheet.getRange("B1");
    date.setValue(`${row[3]}`); //set value of cell
    var series = newFileSheet.getRange("B2");
    series.setValue(`${row[8]}`); //set value of cell
    var speaker = newFileSheet.getRange("B3");
    speaker.setValue(`${row[5]}`); //set value of cell
    var topic = newFileSheet.getRange("B4");
    topic.setValue(`${row[7]}`); //set value of cell
    var topic = newFileSheet.getRange("B5");
    topic.setValue(`${row[9]}`); //set value of cell

  })

}

