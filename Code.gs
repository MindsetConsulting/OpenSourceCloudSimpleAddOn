/**
 * @OnlyCurrentDoc
 *
 * CloudSimple 3.0
    by Naveen Rokkam, Paul J. Modderman, Gavin P. Quinn - Mindset Consulting
    v3.0   20181205 - Open Source Software
    v2.0   20140804 - Updates to UI, removing Forms integration
    v2.0.1 20140817 - More updates to UI, very close to re-submit to google
    v2.0.2 20140818 - Re-submitted to google
    v2.0.3 20140818 - Re-submitted to google after feedback from previous version
    v2.0.4 20140819 - First (private) publication of app
    v2.1.0 20140820 - Re-submission, all Forms stuff stripped out and some other bugs fixed
    v2.1.1 20140904 - using SheetConverter in libraries now.
                      see https://sites.google.com/site/scriptsexamples/custom-methods/sheetconverter
*/

var ADDON_TITLE = 'CloudSimple';
var GOOGLE_ANALYTICS_ID = ''; // Add Google Analytics id or leave blank

/** Runs every time the add-on is started.  Generates the menu. */
function onOpen() {
  SpreadsheetApp
    .getUi()
    .createAddonMenu()
    .addItem('Start', 'showInitialUIWrapper')
    .addItem('Delete all triggers', 'deleteAllTriggers')
    .addToUi();
}

/** Runs when the add-on is installed.  For our purposes just ensures that
 *  the menu is added live to the document when the installation is run.
 */
function onInstall() {
  PropertiesService.getDocumentProperties().deleteAllProperties();
  PropertiesService.getUserProperties().deleteAllProperties();
  Utilities.sleep(1000);
  sendContactEmail();
  onOpen();
}

function showSidebar(setupData, isRefresh, deleteFilters) {
  var existingData = JSON.parse(PropertiesService.getDocumentProperties().getProperty('SETUP_DATA'));
  if(setupData.currentUser === "cs_demo" && setupData.connectionName === "DEMO") {
    if(existingData && existingData != undefined) {
      if(existingData.selectedServiceID != "ZDEMO_CLOUDSIMPLE_SRV_0001") {  //"ZDEMO_SRV_0001"
        // A demo user tried to open a spreadsheet with an existing connection.
        // We don't want them to clear out a registered users settings.
        var ui = SpreadsheetApp.getUi();
        ui.alert(
      'Existing connections',
      'You do not have permission to change the connection settings on this document. Please create a new document to try the demo.',
      ui.ButtonSet.OK);
        return;
      }
    }
  }
  if(!isRefresh){
    PropertiesService.getDocumentProperties().setProperty('SETUP_DATA', JSON.stringify(setupData));
  }

  if(deleteFilters){
    PropertiesService.getDocumentProperties().deleteProperty('FILTERS');
  }

  var ssUi = SpreadsheetApp.getUi();
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('CloudSimple');
  ssUi.showSidebar(ui);

}

function getDataSourceSettings(){
  return JSON.parse(PropertiesService.getDocumentProperties().getProperty('SETUP_DATA'));
}

function getFields(){
  var setupData = JSON.parse(PropertiesService.getDocumentProperties().getProperty('SETUP_DATA'));
  var fieldSet = getEntityFields(setupData.selectedMetadataUrl, setupData.selectedEntity);
  return fieldSet;
}

function showInitialUIWrapper(){
  generateSessionId();
  showInitialUI();
}

function generateSessionId(){
  var date = Utilities.formatDate(new Date(), "GMT", "yyyyMMddHHmmss");
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var sessionId = date + ssId;

  try{
    PropertiesService.getDocumentProperties().deleteProperty('SESSION_ID');
  } catch(err){
  }
  PropertiesService.getDocumentProperties().setProperty('SESSION_ID', sessionId);
}

function showInitialUI(){

  var goToSetup = true;

  try{
    var setupData = JSON.parse(PropertiesService.getDocumentProperties().getProperty('SETUP_DATA'));
    var activeConnection = getActiveConnection();
    if(setupData == undefined || setupData == null ||
      activeConnection == undefined || activeConnection == null){
      goToSetup = true;
    }else{
      goToSetup = false;
    }
  } catch (err) {
    goToSetup = true;
  }

  if(goToSetup){
    showSetup();
  } else {
    showSidebar(setupData, false);
  }

}

/**
 * Used to determine the number of installations.
 */
function sendContactEmail() {
  var user = Session.getActiveUser().getEmail();
  //If we don't have a user name (happening in some gmail.com addresses)
  //try a couple other paths
  if (user == '') {
    user = Session.getActiveUser();
    if (user == '') {
      user = Session.getEffectiveUser();
      if (user == '') {
        user = PropertiesService.getUserProperties().getProperty('USER_NAME');
      }
    }
  }
  setClientId();
  var recipient = 'cloudsimpletrial@mindsetconsulting.com';
  var subject = 'CloudSimple trial user auto-setup';
  var body = 'Please note: ' + user + ' has installed CloudSimple.';
  var options = {};
  options.noReply = true;
  MailApp.sendEmail(recipient, subject, body, options);
}

function setClientId() {
  try {
    // Set client id as the active user key
    PropertiesService.getUserProperties().setProperty('CLIENT_ID', Session.getTemporaryActiveUserKey());
  } catch (err) {
    // Do nothing if this results in an error
  }
}

function getClientId() {
  var client_id = PropertiesService.getUserProperties().getProperty('CLIENT_ID');
  if (!client_id) {
    setClientId();
    client_id = Session.getTemporaryActiveUserKey();
  }
  return {
    ga_id: GOOGLE_ANALYTICS_ID,
    client_id: client_id,
    url: SpreadsheetApp.getActive().getUrl()
  };
}


//PAUL schedule build
function showSchedule(){
  var html = HtmlService.createHtmlOutputFromFile('Schedule').setHeight(400).setWidth(325);
  SpreadsheetApp.getUi().showModalDialog(html, 'Schedule a Recurring Report');
}

function showSetup(){
//  var existingData = JSON.parse(PropertiesService.getDocumentProperties().getProperty('SETUP_DATA'));
//  if(existingData != undefined && existingData.currentUser === "cs_demo" && existingData.connectionName === "DEMO") {
//     //Demo user shouldn't be able to change connection settings
//     showLicenseSplash();
//     return;
//  }
  var html = HtmlService.createHtmlOutputFromFile('Setup').setHeight(500).setWidth(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'CloudSimple Setup');
}

function confirmShowSetup() {
  if(hasExistingSchedule()) {
    var ui = SpreadsheetApp.getUi();
    var userChoice = ui.alert(
      'Remove recurring report?',
      'You have a recurring report set up for this dataset. \nChanging your dataset will remove your recurring report.',
      ui.ButtonSet.YES_NO);

      if(userChoice == ui.Button.YES) {
        deleteCurrentTrigger();
        showSetup();
      }
  } else {
    showSetup();
  }
}

function getOriginalFieldName(fieldName, fields){
  var originalFieldName;
  var field = {};
  for(var i = 0; i < fields.length; i++){
    if(fields[i].upperName == fieldName){
      field.Name = fields[i].Name;
      field.Type = fields[i].Type;
      return field;
    }
  }
}


function getMessageText(responseText){
  var startPos = responseText.lastIndexOf('<message>') + 9;
  var endPos = responseText.lastIndexOf('</message>');
  return responseText.substring(startPos, endPos);
}

function showCleanupConfirmation(ui){
  var userChoice = ui.alert(
    'Remove old data?',
    'Old SAP data sheets will be permanently deleted.',
    ui.ButtonSet.YES_NO);

  return userChoice;
}

/** Handles the user clicking the Go! button on the interface Sidebar.
 *  @param {string} action - The name of the action to take.
 *  @param {string} entitySet - The name of the entity in use.
 *  @param {string} metadata - The URL of the $metadata service for the given entity set.
 */
function handleUserAction(action, entitySet, metadata, selectFields, filterIds, encodedAuth){

  try{
    var applicationData = {};
    applicationData.action = action;
    applicationData.entitySet = entitySet;
    applicationData.metadata = metadata;
    applicationData.selectFields = selectFields;
    applicationData.filterFields = filterIds;
    PropertiesService
      .getDocumentProperties()
      .setProperty("APPLICATION_DATA", JSON.stringify(applicationData));

    var result;
    switch(action) {
      case 'createform':
        //As is in published version of the app, we can't use Forms to automatically trigger
        //backend processing. To see the way this used to work, consult build archive CloudShuttle version 1.0
        break;
      case 'connectsheet':
        result = readService(selectFields, filterIds, true, false);
        break;
    }
  }
  catch(err){
    quickLog('Error in user action: ' + err + genericErrorMessage);
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
      'Can\'t retrieve records',
      'There was an error getting your data from SAP. You can check the Log sheet for technical details.',
      ui.ButtonSet.OK);
  }

  return result;
}

function Fieldset() {
  this.fields = [];
  this.entitySetAttributes = [];
}

function onEdit(event){
  var activeSheet = SpreadsheetApp.getActiveSheet().getName();
  if(activeSheet == 'Log'){return;}

  var cloudshuttleRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('CLOUDSHUTTLE_DATA');
  var changedSheet = event.range.getSheet().getIndex();
  if(changedSheet == cloudshuttleRange.getSheet().getIndex()){
    var activeRange = SpreadsheetApp.getActiveSheet().getDataRange();
    var width = activeRange.getWidth();
    var height = activeRange.getHeight();
    for(var i = event.range.rowStart; i <= event.range.rowEnd; i++){
      var changedKey = SpreadsheetApp.getActiveSheet()
                         .getRange(i, width)
                         .getValue();// + ':updated';
      var updateCheck = changedKey.split(':');
      if(updateCheck.length < 2 && width > 1){
        changedKey += ':updated';
        SpreadsheetApp
          .getActiveSheet()
          .getRange(i, width)
          .setValue(changedKey);
      }
    }
  }
}

/**
 * This function is used for logging application errors to make them
 * easier to debug.
 *
 * @param {string} log - Text to log to the sheet
 */
function quickLog(log){
  if(log){
    var logSheet = SpreadsheetApp.getActive().getSheetByName('Log');
    if(logSheet == null){
      //If it doesn't exist, create it and hide it
      logSheet = SpreadsheetApp.getActive().insertSheet('Log');
      logSheet.hideSheet();
    }
    var nextRow = logSheet.getDataRange().getHeight();
    var logHeader = 'Application log entry: recorded at ' + Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
    logSheet.getRange(nextRow + 1, 1).setValue(logHeader);
    logSheet.getRange(nextRow + 1, 2).setValue(log);
  }
}

var genericErrorMessage = '\n\nHere are some general error sources and suggested resolutions.\n\n' +
                          '1. Filters are not specific enough. The SAP system is spending too much ' +
                          'time processing your results. This can cause "Syntax Error" messages. ' +
                          'Try being as specific as possible when you create filters, to reduce the ' +
                          'volume of data sent from SAP.\n\n' +
                          '2. Your SAP Gateway system is responsive but its connection to the SAP Business ' +
                          'Suite (ERP) system is unresponsive. Attempt to log in to the Business Suite ' +
                          'system or contact your system administrator.\n\n' +
                          '3. An error may have occurred in your SAP system that has not been passed back ' +
                          'to CloudSimple. Your system administrator may be able to assist.\n\n' +
                          '4. For general CloudSimple help, feel free to contact Mindset at info@mindsetconsulting.com';

function cleanupAllProperties(){
  PropertiesService.getDocumentProperties().deleteAllProperties();
  PropertiesService.getScriptProperties().deleteAllProperties();
  PropertiesService.getUserProperties().deleteAllProperties();
}
