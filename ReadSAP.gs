/*
* CloudSimple 3.0
   by Naveen Rokkam, Paul J. Modderman, Gavin P. Quinn - Mindset Consulting
   v3.0   20181205 - Open Source Software
*/
/** Gets a list of all the services available in a given catalog.
 *  @param {string} url - The catalog service URL.
 *  @returns {object} JSON of service results.
 */
function readAllServices() {
  try {
    var activeConnection = getActiveConnection();
    var options = {
      headers: { 'Authorization': 'Basic ' + activeConnection.base64Auth,
                'X-CSRF-Token': 'Fetch'
               },
    };

    var url = activeConnection.url + '/sap/opu/odata/IWFND/CATALOGSERVICE;v=2/ServiceCollection?$expand=EntitySets&$format=json';
    var response = UrlFetchApp.fetch(url, options);
    var jsonString = response.getContentText();

    //If the response contains a logon error, pop the user back to the setup screen
    if(jsonString.lastIndexOf('Logon Error Message') > -1){
      var ui = SpreadsheetApp.getUi();
      ui.alert(
        'Logon failed',
        'We couldn\'t log you in to your chosen SAP system. Check your connection settings.',
        ui.ButtonSet.OK);

      showSetup();
      return;
    }

    var jsonObj = JSON.parse(jsonString);

    //this sort will sort the services by description
    jsonObj.d.results.sort(function(a,b) {return (a.Description > b.Description) ? 1 : ((b.Description > a.Description) ? -1 : 0);} );

    //This will sort the entities of each service by their description
    for(var i = 0; i < jsonObj.d.results.length; i++){
      jsonObj.d.results[i].EntitySets.results
      .sort(function(a,b) {return (a.Description > b.Description) ? 1 : ((b.Description > a.Description) ? -1 : 0);} );
    }

    return jsonObj;
  }catch (error){
    quickLog('An error occurred establishing the connection: ' + error.message);
    logError('readAllServices() ' + error.message);
    if(error.message.indexOf("<title>Logon Error Message</title>") > -1){
      throw 'password';
    } else {
      throw 'error';
    }
  }
}


function getAuthenticatingResponse(resource) {
  //This function is part of the strangely frustrating CSRF process.  We may be
  //ditching using this on the backend in favor of using the client-side code to send requests.
  var encodedAuth = PropertiesService.getUserProperties().getProperty('AUTH');
  var options = {
    headers: { 'Authorization': 'Basic ' + encodedAuth,
              'X-CSRF-Token': 'Fetch'
             },
  };

  var response = UrlFetchApp.fetch(resource,options);
  var responseString = response.getContentText();

  //If the response contains a logon error, pop the user back to the setup screen
  if(responseString.lastIndexOf('Logon Error Message') > -1){
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      'Logon failed',
      'We couldn\'t log you in to your chosen SAP system. Check your connection settings.',
      ui.ButtonSet.OK);

    showSetup();
    return;
  }


  var authenticatingResponse = {};
  authenticatingResponse.csrfToken = response.getAllHeaders()['x-csrf-token'];
  var setCookie = response.getAllHeaders()['set-cookie'];
  var relevantCookie = JSON.stringify(setCookie).split(';')[0];
  authenticatingResponse.cookieName = relevantCookie.split('=')[0].replace('\"', '');
  authenticatingResponse.cookieValue = relevantCookie.split('=')[1];
  return authenticatingResponse;
}

/** Reads the current service/entity for data matching the select/filter criteria.
 *  @param {string} metadataURL - URL for service metadata.
 *  @param {string} entitySet - Name of the entityset being selected.
 *  @param {array} selectFields - Fields to be selected in the service call.
 *  @param {array} filterFields - Criteria for filtering the result set.
 */
function readService(selectFields, filterIds, showDialog, newTab) {

  //PAUL
  //Added to support the background running of reports. Save the select fields and filters
  //to persistent data
  PropertiesService.getUserProperties().setProperty('SELECT_FIELDS', JSON.stringify(selectFields));
  PropertiesService.getUserProperties().setProperty('FILTER_IDS', JSON.stringify(filterIds));
  //End added support

  var setupData = JSON.parse(PropertiesService.getDocumentProperties().getProperty('SETUP_DATA'));
  var entitySet = setupData.selectedEntity;
  var url = setupData.selectedMetadataUrl.replace('$metadata', '') + entitySet;

  var encodedAuth = PropertiesService.getUserProperties().getProperty('AUTH');
  var options = {
    headers: { 'Authorization': 'Basic ' + encodedAuth,
               'X-CSRF-Token': 'Fetch'
             },
  };

  //Delete all sheets but the Log sheet.  Then create a new sheet
  //with the entityset name
  var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();

  var selectStatement = '';
  if(selectFields != undefined){
    for(var i = 0; i < selectFields.length; i++){
      if(i == 0){
        selectStatement += selectFields[i];
      }else{
        selectStatement += ',' + selectFields[i];
      }
    }
  }

  var filterStatement = '';

  var filterFields = getFilterFields(filterIds);
  if(filterFields != null){
    for(var i = 0; i < filterFields.length; i++){
      var field = filterFields[i];
      if(filterStatement.length > 0){
        filterStatement += ' and ';
      }
      if(field.type != 'in'){
        var fieldValue = formatFilterFieldValue(field);
        filterStatement += field.name + ' ' + field.type + ' ' + fieldValue + ' ';
      }else{
        var values = field.value.split(',');
        for(var i = 0; i < values.length; i++){
          filterStatement += field.name + ' eq ' + '\'' + values[i] + '\' ';
          if(i < values.length - 1){
            filterStatement += 'or ';
          }
        }
      }
    }
  }

  var queryStatement = '';
  if(selectStatement != ''){
    queryStatement += '&$select=' + selectStatement;
  }
  if(filterStatement != ''){
    queryStatement += '&$filter=' + filterStatement;
  }

  var jsonFormat = '?$format=json';
  url = url.trim() + jsonFormat.trim() + queryStatement.trim();

  var response = UrlFetchApp.fetch(url, options);
  var jsonString = response.getContentText();

  //If the response contains a logon error, pop the user back to the setup screen
  if(jsonString.lastIndexOf('Logon Error Message') > -1){
    if(showDialog){
      var ui = SpreadsheetApp.getUi();
      ui.alert(
        'Logon failed',
        'We couldn\'t log you in to your chosen SAP system. Check your connection settings.',
        ui.ButtonSet.OK);

      showSetup();
      return;
    } else {
      return;
    }
  }

  var jsonObj = JSON.parse(jsonString);

  if(jsonObj.d.results.length < 1){
    newSheet.getRange(1,1).setValue('No results returned from SAP for given filter criteria.');
    return;
  }

  var resultData = [];
  var columnsToGray = [];
  var numberColumns = [];
  var fieldHeaders = [];
  var fieldOrder = [];
  var resultColors = [];
  var headerColors = [];

  var fieldNames = Object.keys(jsonObj.d.results[0]);
  var fieldObjTmp = JSON.parse(PropertiesService.getDocumentProperties().getProperty('FIELDS'));
  var fieldObj = {fields:[]};

  for(var i = 0; i < fieldObjTmp.fields.length; i++){
    for(var j = 0; j < fieldNames.length; j++){
      if(fieldObjTmp.fields[i].Name == fieldNames[j]){
        fieldObj.fields.push(fieldObjTmp.fields[i]);
        var fieldHeader;
        if(fieldObjTmp.fields[i].label != undefined){
          fieldHeader = fieldObjTmp.fields[i].label + ' (' + fieldObjTmp.fields[i].Name.toUpperCase() + ')';
        }else{
          fieldHeader = '(' + fieldObjTmp.fields[i].Name.toUpperCase() + ')';
        }
        fieldHeaders.push(fieldHeader);
        headerColors.push('#FFFFFF');


        if(fieldObjTmp.fields[i].updatable != undefined){
          columnsToGray.push(i);
        }

        if(fieldObjTmp.fields[i].Type == 'Edm.Decimal'){
          numberColumns.push(i);
        }

      }
    }
  }

  //This would be a good place to sort the result fields.
  fieldObj.fields.sort(function(a,b) {return (a.label > b.label) ? 1 : ((b.label > a.label) ? -1 : 0);} );

  fieldHeaders.push('CloudSimpleRecordID');
  headerColors.push('#FFFFFF');

  if(fieldHeaders.length > 1){
    resultData.push(fieldHeaders);
    resultColors.push(headerColors);
  }


  for(var i = 0; i < jsonObj.d.results.length; i++){
    var result = jsonObj.d.results[i];
    var resultItem = [];
    var resultColor = [];
    var id = result.__metadata.id.split('/');
    var recordID = id[id.length - 1];

    var propertyCount = 0;
    //for(var property in result){
    for(var j = 0; j < fieldObj.fields.length; j++){
      var property = fieldObj.fields[j].Name;
      if(property != '__metadata'){
        var itemValue = result[property];
        if(itemValue != null && itemValue != undefined && itemValue.length > 0){
          if(itemValue.indexOf('/Date(') == 0){
            var itemDateValue = new Date(parseInt(itemValue.substr(6)));
            itemValue = Utilities.formatDate(itemDateValue, 'GMT', 'yyyy-MM-dd');
          }
        }

        if(numberColumns.indexOf(propertyCount) > -1){
          resultItem.push(itemValue);
        }else{
          resultItem.push('\'' + itemValue);
        }

        if(columnsToGray.indexOf(propertyCount) > -1){
          resultColor.push('#e3e3e3');
        }else{
          resultColor.push('#FFFFFF');
        }

      }

      propertyCount++;
    }

    //Put the record ID in there
    if(resultItem.length > 0){
      resultItem.push(recordID);
      resultColor.push('#e3e3e3');
    }

    //Enter the entire record
    if(resultItem.length > 0){
      resultData.push(resultItem);
      resultColors.push(resultColor);
    }
  }

  var numRows = resultData.length;
  var numColumns = resultData[0].length;

  //This is a way faster way of setting the data in the sheet
  if(numRows > 0 && numColumns > 0){
    var range = newSheet.getRange(1,1, numRows, numColumns);
    range.setValues(resultData);
    range.setBackground(resultColors);
    var spreadsheet = newSheet.getParent();
    spreadsheet.setNamedRange('CLOUDSHUTTLE_DATA', range);
  }

  //Now set coloring to indicate which are editable
  var fields = resultData[0];

  newSheet
    .getRange(1,1,1,numColumns)
    .setBackground('LightGray')
    .setFontWeight('bold');

  newSheet.setFrozenRows(1);
  newSheet.hideColumns(numColumns);

  var doCleanup = true;
  if(newTab) {
    doCleanup = false;
  }

  if(showDialog){
    var ui = SpreadsheetApp.getUi();
    var userChoice = showCleanupConfirmation(ui);
    if(userChoice == ui.Button.NO){ doCleanup = false; }
  }
  if(doCleanup){
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var logID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log').getSheetId();
    var newSheetID = newSheet.getSheetId();
    for(var i = 0; i < sheets.length; i++){
      var sheetID = sheets[i].getSheetId();
      if(sheetID != logID && sheetID != newSheetID){
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheets[i]);
      }
    }
    newSheet.setName(entitySet);
  }else{
    // Last run is set right before calling this function when a scheduled
    // report is run. Defaults to the current timestamp.
    var nameSuffix = getFormattedDate(new Date(), Session.getScriptTimeZone());
    newSheet.setName(entitySet + ' ' + nameSuffix);
  }

  var retrievedRows = numRows -1;
  return "Retrieved " + retrievedRows + " records from SAP backend.";
}

/** Gets an object containing all metadata about entity fields.
 *  @param {string} metadataURL - The $metadata metadataURL for the service.
 *  @param {string} entityName - That's what it is.
 */
function getEntityFields(metadataURL, entitySetName) {
  var encodedAuth = PropertiesService.getUserProperties().getProperty('AUTH');
  var fieldSet = new Fieldset();
  var options = {
    headers: { 'Authorization': 'Basic ' + encodedAuth,
               'X-CSRF-Token': 'Fetch'
             },
  };

  try{
    var response = UrlFetchApp.fetch(metadataURL, options);
    var xml = response.getContentText();
    var document = XmlService.parse(xml);
    var root = document.getRootElement();
    var edmxNS = XmlService.getNamespace('http://schemas.microsoft.com/ado/2007/06/edmx');
    var schemaNS = XmlService.getNamespace('http://schemas.microsoft.com/ado/2008/09/edm');
    var dataServices = root.getChild('DataServices', edmxNS);
    var schema = dataServices.getChild('Schema', schemaNS);
    var entityContainer = schema.getChild('EntityContainer', schemaNS);
    var entitySet = entityContainer.getChildren('EntitySet', schemaNS);

    for(var i = 0; i < entitySet.length; i++) {
      var name = entitySet[i].getAttribute('Name').getValue();
      if(name == entitySetName){
        var typeName = entitySet[i].getAttribute('EntityType').getValue();
        var split = typeName.split('.');
        var entityName = split[1];
        var attribute = {};
        var attributes = entitySet[i].getAttributes();
        for(var j = 0; j < attributes.length; j++){
          var attribName = attributes[j].getName();
          var attribValue = attributes[j].getValue();
          attribute[attribName] = attribValue;
        }
        fieldSet.entitySetAttributes = attribute;
        break;  //exit the entitySet loop - only one will match
      }
    }


    var entityType = schema.getChildren('EntityType', schemaNS);
    for(var i = 0; i < entityType.length; i++){
      var name = entityType[i].getAttribute('Name').getValue();
      if(name == entityName){
        var property = entityType[i].getChildren('Property', schemaNS);
        for(var j = 0; j < property.length; j++){
          var field = {};
          var attributes = property[j].getAttributes();
          for(var k = 0; k < attributes.length; k++){
            var attribName = attributes[k].getName();
            var attribValue = attributes[k].getValue();
            field[attribName] = attribValue;
          }
          field.upperName = field.Name.toUpperCase();
          fieldSet.fields.push(field);
        }
        break;  //Exit the entityType loop, only need to find one match
      }
    }
  }catch(err){
    quickLog('Error getting entity fields: ' + err.message);
    logError('getEntityFields() ' + err.message);
    //If the response contains a logon error, pop the user back to the setup screen
    if(err.message.lastIndexOf('Logon Error Message') > -1){
      var ui = SpreadsheetApp.getUi();
      ui.alert(
        'Logon failed',
        'We couldn\'t log you in to your chosen SAP system. Check your connection settings.',
        ui.ButtonSet.OK);

      showSetup();
      return;
    }
    throw 'error';
  }

  fieldSet.fields.sort(function(a,b) {return (a.label > b.label) ? 1 : ((b.label > a.label) ? -1 : 0);} );
  PropertiesService.getDocumentProperties().setProperty('FIELDS', JSON.stringify(fieldSet));
  return fieldSet;
}
