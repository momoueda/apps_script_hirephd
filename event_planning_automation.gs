function onOpen(){
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
      .addItem('Create All Event Documents', 'addAll')
      .addSubMenu(ui.createMenu('Manually Create')
        .addItem('Add Event Plan Master Sheet Link', 'addEventMasterLink')
        .addItem('Add Event Folder Link', 'addFolderLink')
        .addItem('Create New Slides', 'createSlides')
        .addItem('Create New Questions Doc', 'createQAdoc')
        .addItem('Create New Networking Doc', 'createNetworkingSheet')
        .addItem('Create New Speaker Package Doc', 'createSpeakerPackage')
        .addItem('Create New Partner Package Doc', 'createPartnerPackage')
        .addItem('Create New Event Description Doc', 'createEventDescription')
        .addItem('Create New Feedback Form', 'createFeedbackForm')
      )
      .addItem('Import Speaker Info', 'speakerInfo')
      .addItem('Notify Team in Slack', 'sendSlack')
      .addToUi();

}

function sendSlack(){
  const ss = SpreadsheetApp.getActive();
  let data = ss.getSheetByName('Speaker Info & Agenda').getRange("E5:E8").getValues();
  let payload = buildAlert(data);
  sendAlert(payload);
}

function buildAlert(data) {
  let speakerName = data[3][0];
  let eventMonth = data[0][0]
  
  let payload = {
    "blocks": [
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": ":bell: *Speaker Info Imported* :bell:"
        }
      },
      {
        "type": "divider"
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "Import of speaker info completed for " + eventMonth + "'s event"
        }
      },
       {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "Speaker name: " + speakerName
        }
      }
    ]
  };
  return payload;
}

function sendAlert(payload) {
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config (leave untouched)');
  var webhook = config.getRange("B11").getValues();
  var options = {
    "method": "post", 
    "contentType": "application/json", 
    "muteHttpExceptions": true, 
    "payload": JSON.stringify(payload) 
  };
  
  try {
    UrlFetchApp.fetch(webhook, options);
  } catch(e) {
    Logger.log(e);
  }
}

function speakerInfo(){
  const speakerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Speaker form entry');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config (leave untouched)');
  const auto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Automation (leave untouched)');
  var responsesID = config.getRange("B9").getValues();;
  var responsesSheet = SpreadsheetApp.openById(responsesID).getSheetByName('Form Responses 1');
  var responsesDataRange = responsesSheet.getDataRange();
  var values = responsesDataRange.getValues();
  var email = auto.getRange("B5").getValues();
  var speakerRowNum;

  for (var i = 0; i < values.length; i++)
  {
    if (values[i][3] == email)
    {
      speakerRowNum = i+1;
      break;
    }
  }
 
  var speakerRow = responsesSheet.getRange(speakerRowNum,1,1,20).getValues();
  speakerSheet.getRange(2,1,1,20).setValues(speakerRow);

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

  var slidesTemplateID = config.getRange(1,2).getValues();
  var destinationFolder = DriveApp.getFolderById(folderId);

  //get values for date, series, speaker and topic
  var date = auto.getRange("B1").getValues().toString();
  var series = auto.getRange("B2").getValues().toString();
  var speaker = auto.getRange("B3").getValues().toString();
  var topic = auto.getRange("B4").getValues().toString();

    //Create a copy of slides template
  var slidesTemplate = DriveApp.getFileById(slidesTemplateID);
  var slidesCopy = slidesTemplate.makeCopy(`Slides_` + date + `_` + series + `_` + speaker + `_` + topic, destinationFolder);
  var slidesUrl = slidesCopy.getUrl();
    
  sheet.getRange("E28").setValue(slidesUrl);

}

function createQAdoc(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config (leave untouched)');
  const auto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Automation (leave untouched)');
  
  //get folderId
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var sheetFile = DriveApp.getFileById(sheetId);
  var folderId = sheetFile.getParents().next().getId();

  var qaTemplateID = config.getRange("B3").getValues();;
  var destinationFolder = DriveApp.getFolderById(folderId);

  //get values for date, series, speaker and topic
  var date = auto.getRange("B1").getValues().toString();
  var series = auto.getRange("B2").getValues().toString();
  var speaker = auto.getRange("B3").getValues().toString();
  var topic = auto.getRange("B4").getValues().toString();

    //Create a copy of QA template
  var qaTemplate = DriveApp.getFileById(qaTemplateID);
  var qaCopy = qaTemplate.makeCopy(`Questions_Doc_`+ date + '_' + series + '_' + speaker + '_' + topic, destinationFolder);
  var qaUrl = qaCopy.getUrl();
    
  sheet.getRange("E29").setValue(qaUrl);

}

function createNetworkingSheet(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config (leave untouched)');
  const auto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Automation (leave untouched)');
  
  //get folderId
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var sheetFile = DriveApp.getFileById(sheetId);
  var folderId = sheetFile.getParents().next().getId();

  var networkingTemplateID = config.getRange("B2").getValues();;
  var destinationFolder = DriveApp.getFolderById(folderId);

  //get values for date, series, speaker and topic
  var date = auto.getRange("B1").getValues().toString();
  var series = auto.getRange("B2").getValues().toString();
  var speaker = auto.getRange("B3").getValues().toString();
  var topic = auto.getRange("B4").getValues().toString();

    //Create a copy of slides template
  var networkingTemplate = DriveApp.getFileById(networkingTemplateID);
  var networkingCopy = networkingTemplate.makeCopy(`Networking_Doc_` + date + '_' + series + '_' + speaker + '_' + topic, destinationFolder);
  var networkingUrl = networkingCopy.getUrl();
    
  sheet.getRange("E30").setValue(networkingUrl);

}

function createSpeakerPackage(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config (leave untouched)');
  const auto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Automation (leave untouched)');
  
  //get folderId
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var sheetFile = DriveApp.getFileById(sheetId);
  var folderId = sheetFile.getParents().next().getId();

  var speakerPackageTemplateID = config.getRange("B5").getValues();;
  var destinationFolder = DriveApp.getFolderById(folderId);

  //get values for date, series, speaker and topic
  var date = auto.getRange("B1").getValues().toString();
  var series = auto.getRange("B2").getValues().toString();
  var speaker = auto.getRange("B3").getValues().toString();
  var topic = auto.getRange("B4").getValues().toString();

    //Create a copy of slides template
  var speakerPackageTemplate = DriveApp.getFileById(speakerPackageTemplateID);
  speakerPackageTemplate.makeCopy(`Speaker_Package_` + date + '_' + series + '_' + speaker + '_' + topic, destinationFolder);
    
}

function createPartnerPackage(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config (leave untouched)');
  const auto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Automation (leave untouched)');

  
  //get folderId
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var sheetFile = DriveApp.getFileById(sheetId);
  var folderId = sheetFile.getParents().next().getId();

  var partnerPackageTemplateID = config.getRange("B6").getValues();;
  var destinationFolder = DriveApp.getFolderById(folderId);

  //get values for date, series, speaker and topic
  var date = auto.getRange("B1").getValues().toString();
  var series = auto.getRange("B2").getValues().toString();
  var speaker = auto.getRange("B3").getValues().toString();
  var topic = auto.getRange("B4").getValues().toString();

    //Create a copy of slides template
  var partnerPackageTemplate = DriveApp.getFileById(partnerPackageTemplateID);
  partnerPackageTemplate.makeCopy(`Partner_Package_`+ date + '_' + series + '_' + speaker + '_' + topic, destinationFolder);

}

function createEventDescription(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config (leave untouched)');
  const auto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Automation (leave untouched)');
  
  
  //get folderId
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var sheetFile = DriveApp.getFileById(sheetId);
  var folderId = sheetFile.getParents().next().getId();

  var eventDescriptionTemplateID = config.getRange("B4").getValues();;
  var destinationFolder = DriveApp.getFolderById(folderId);

  //get values for date, series, speaker and topic
  var date = auto.getRange("B1").getValues().toString();
  var series = auto.getRange("B2").getValues().toString();
  var speaker = auto.getRange("B3").getValues().toString();
  var topic = auto.getRange("B4").getValues().toString();

    //Create a copy of slides template
  var eventDescriptionTemplate = DriveApp.getFileById(eventDescriptionTemplateID);
  var eventDescriptionCopy = eventDescriptionTemplate.makeCopy(`Event_Descriptions_Doc_`+ date + '_' + series + '_' + speaker + '_' + topic, destinationFolder);
  var eventDescriptionUrl = eventDescriptionCopy.getUrl();
    
  sheet.getRange(17, 5).setValue(eventDescriptionUrl);
}

function createFeedbackForm(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main');
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config (leave untouched)');
  const auto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Automation (leave untouched)');
  
  //get folderId
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var sheetFile = DriveApp.getFileById(sheetId);
  var folderId = sheetFile.getParents().next().getId();

  var feedbackFormTemplateID = config.getRange("B8").getValues();;
  var destinationFolder = DriveApp.getFolderById(folderId);

  //get values for date, series, speaker and topic
  var date = auto.getRange("B1").getValues().toString();
  var series = auto.getRange("B2").getValues().toString();
  var speaker = auto.getRange("B3").getValues().toString();
  var topic = auto.getRange("B4").getValues().toString();

    //Create a copy of feedback form template
  var feedbackFormTemplate = DriveApp.getFileById(feedbackFormTemplateID);
  feedbackFormTemplate.makeCopy(`Feedback_Form_` + date + '_' + series + '_' + speaker + '_' + topic, destinationFolder);
  var form = FormApp.openById(feedbackFormTemplateID);
  var feedbackFormUrl = form.getPublishedUrl();
    
  sheet.getRange("E32").setValue(feedbackFormUrl);
}


//
function addAll(){
  //Add event master link
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main');
  const url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  sheet.getRange(11, 5).setValue(url)

  //Add Folder Link
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var sheetFile = DriveApp.getFileById(sheetId);
  var folderUrl = sheetFile.getParents().next().getUrl();
  sheet.getRange(14, 5).setValue(folderUrl);

  //Create Slides
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('config (leave untouched)');
  const auto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Automation (leave untouched)');
  
    //get folderId
  var folderId = sheetFile.getParents().next().getId();
  var destinationFolder = DriveApp.getFolderById(folderId);

  var slidesTemplateID = config.getRange(1,2).getValues();;

  //get values for date, series, speaker and topic
  var date = auto.getRange("B1").getValues().toString();
  var series = auto.getRange("B2").getValues().toString();
  var speaker = auto.getRange("B3").getValues().toString();
  var topic = auto.getRange("B4").getValues().toString();

    //Create a copy of slides template
  var slidesTemplate = DriveApp.getFileById(slidesTemplateID);
  var slidesCopy = slidesTemplate.makeCopy(`Slides_` + date + `_` + series + `_` + speaker + `_` + topic, destinationFolder);
  var slidesUrl = slidesCopy.getUrl();
    
  sheet.getRange(25, 5).setValue(slidesUrl);

  //Create Q and A Doc
  var qaTemplateID = config.getRange("B3").getValues();

    //Create a copy of QA template
  var qaTemplate = DriveApp.getFileById(qaTemplateID);
  var qaCopy = qaTemplate.makeCopy(`Questions_Doc_` + date + '_' + series + '_' + speaker + '_' + topic, destinationFolder);
  var qaUrl = qaCopy.getUrl();
    
  sheet.getRange(26, 5).setValue(qaUrl);

  //Create Networking Sheet
  var networkingTemplateID = config.getRange("B2").getValues();;

    //Create a copy of networking template
  var networkingTemplate = DriveApp.getFileById(networkingTemplateID);
  var networkingCopy = networkingTemplate.makeCopy('Networking_Doc_' + date + '_' + series + '_' + speaker + '_' + topic, destinationFolder);
  var networkingUrl = networkingCopy.getUrl();
    
  sheet.getRange(27, 5).setValue(networkingUrl);

  //Create Speaker Package
  var speakerPackageTemplateID = config.getRange("B5").getValues();;

    //Create a copy of speaker package template
  var speakerPackageTemplate = DriveApp.getFileById(speakerPackageTemplateID);
  speakerPackageTemplate.makeCopy(`Speaker_Package_` + date + '_' + series + '_' + speaker + '_' + topic, destinationFolder);
    
  //Create Partner Package
   var partnerPackageTemplateID = config.getRange("B6").getValues();;

    //Create a copy of partner package template
  var partnerPackageTemplate = DriveApp.getFileById(partnerPackageTemplateID);
  partnerPackageTemplate.makeCopy(`Partner_Package_` + date + '_' + series + '_' + speaker + '_' + topic, destinationFolder);

  //Create Event Description
  var eventDescriptionTemplateID = config.getRange("B4").getValues();;

    //Create a copy of event description template
  var eventDescriptionTemplate = DriveApp.getFileById(eventDescriptionTemplateID);
  var eventDescriptionCopy = eventDescriptionTemplate.makeCopy(`Event_Descriptions_Doc_` + date + '_' + series + '_' + speaker + '_' + topic, destinationFolder);
  var eventDescriptionUrl = eventDescriptionCopy.getUrl();
    
  sheet.getRange(17, 5).setValue(eventDescriptionUrl);

  //Create Feedback Form
  var feedbackFormTemplateID = config.getRange("B8").getValues();;

    //Create a copy of feedback form template
  var feedbackFormTemplate = DriveApp.getFileById(feedbackFormTemplateID);
  feedbackFormTemplate.makeCopy(`Feedback_Form_` + date + '_' + series + '_' + speaker + '_' + topic, destinationFolder);
  var form = FormApp.openById(feedbackFormTemplateID);
  var feedbackFormUrl = form.getPublishedUrl();
    
  sheet.getRange(29, 5).setValue(feedbackFormUrl);

}

