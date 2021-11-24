function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
      .addSubMenu(ui.createMenu('Main')
        .addItem('Add Event Plan Master Sheet Link', 'addEventMasterLink'))
      .addToUi();
}

function addEventMasterLink(){
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main');
    const url = SpreadsheetApp.getActiveSpreadsheet().getUrl();

    sheet.getRange(11, 5).setValue(url)
}
