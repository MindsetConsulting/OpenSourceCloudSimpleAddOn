/*
* CloudSimple 3.0
   by Naveen Rokkam, Paul J. Modderman, Gavin P. Quinn - Mindset Consulting
   v3.0   20181205 - Open Source Software
*/
function updateBackendMultipleTransactions() { 

  try{
    //Sheet information
    var activeRange = SpreadsheetApp.getActiveSheet().getDataRange();
    var width = activeRange.getWidth();
    var height = activeRange.getHeight();

    //Get the base url for the service, just take it out of the metadata value
    var setupData = JSON.parse(PropertiesService.getDocumentProperties().getProperty("SETUP_DATA"));
    var url = setupData.selectedMetadataUrl.replace('$metadata', '$batch');

    //Handle CSRF token
    //var authenticatingResponse = getAuthenticatingResponse(setupData.selectedMetadataUrl);

    //Necessary payload information
    var payloadBegin = '--batch\nContent-Type: multipart/mixed; boundary=changeset\n\n'
    var payload = '';
    var fullPayload;
    var payloadEnd = '\n--changeset--\n--batch--';

    //add column headers as properties
    var columnNames = {};
    var fieldSet = JSON.parse(PropertiesService.getDocumentProperties().getProperty('FIELDS'));

    for(var i = 1; i < width; i++){
      var wholeCaption = activeRange.getCell(1,i).getValue().split('(');
      var fieldName = wholeCaption[wholeCaption.length - 1].replace(/\)/,'');
      fieldName = getOriginalFieldName(fieldName, fieldSet.fields);
      columnNames[i] = fieldName;
    }

    var changedRows = [];
    var rowResults = [];
    //Here's where we look at the data.  Starting at height 2, to skip the header row
    for(var currentRow = 2; currentRow <= height; currentRow++){
      var entry = '';
      var entryData = {d:{}};
      var title = SpreadsheetApp
        .getActiveSheet()
        .getRange(currentRow,width)
        .getValue();

      //The value in the last column will have ':updated' if it's necessary to
      //send that row to SAP for update
      var titleValues = title.split(':');
      if(titleValues.length > 1){
        //assign data to the JSON object that goes to SAP
        for(var currentColumn = 1; currentColumn < width; currentColumn++){
          var cellData = activeRange.getCell(currentRow,currentColumn).getValue();
          cellData = encodeURI(cellData);
          var column = columnNames[currentColumn];
          if(column.Type.toUpperCase() == 'EDM.STRING'){
            entryData.d[column.Name] = cellData.toString();
          } else if(column.Type.toUpperCase() == 'EDM.DATETIME') {
            entryData.d[column.Name] = cellData + 'T00:00:00';
          } else {
            entryData.d[column.Name] = cellData;
          }
        }

        //Necessary headers for this particular entry.  Need the content-length
        entry = '--changeset\nContent-Type: application/http\nContent-Transfer-Encoding: binary\n\n';
        entry += 'PUT ' + titleValues[0] + ' HTTP/1.1\n'
        entry += 'Content-Type: application/json; charset=utf-8\n';
        entry += 'Content-Length: ' +  JSON.stringify(entryData).length + '\n\n';

        //Now build the request.  Payload, headers, options.
        var entryString = payloadBegin + entry + JSON.stringify(entryData) + '\n\n' + payloadEnd;
        payload = entryString;

        //var encodedAuth = PropertiesService.getDocumentProperties().getProperty('AUTH');
        var encodedAuth = PropertiesService.getUserProperties().getProperty('AUTH');

        //Now getting strange error "Unexpected exception upon serialization continuation"
        //Be mindful of that with this little snippet.
        var headers = {};
        headers.method = 'post';
        headers.Authorization = 'Basic ' + encodedAuth;
        headers['X-Requested-With'] = 'XMLHTTPRequest';
        headers['Content-Type'] = 'multipart/mixed; boundary=batch';
        //headers['X-CSRF-Token'] = authenticatingResponse.csrfToken;
        //headers[authenticatingResponse.cookieName] = authenticatingResponse.cookieValue;

        var options = {};
        options.headers = headers;
        options.payload = payload;

        //Call out to the service and build the structure of result statuses for all requests
        try{
          var lock = LockService.getPublicLock();
          lock.waitLock(30000);

          var response = UrlFetchApp.fetch(url, options);
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

          var result = {};
          result.rowNumber = currentRow;
          if(response.getResponseCode() == 202){
            if(response.getContentText().indexOf('1.1 204') > -1){
              result.color = 'LightGreen';
              var sheet = SpreadsheetApp.getActiveSheet();
              sheet.getRange(currentRow,width).setValue(titleValues[0]);
              sheet.getRange(currentRow, 1).clearNote();
            }else{
              result.color = 'Red';
              result.message = getMessageText(response.getContentText());
            }
          }else{
            result.color = 'Red';
            result.message = getMessageText(response.getContentText());
          }
//          rowResults.push(result);

          activeRange.offset(result.rowNumber - 1, 0, 1).setBackground(result.color);
          if(result.color == 'Red'){
            activeRange.offset(result.rowNumber - 1, 0, 1, 1).setNote(result.message);
          }

        }catch (err){
          quickLog('Error occurred in single mode update: ' + err.message);
          logError('updateBackendMultipleTransactions() inner catch ' + err.message);
          throw err.message;
        }finally{
          lock.releaseLock();
        }
      }
    }

//    //Flag the changed rows with their status colors
//    for(var i = 0; i < rowResults.length; i++){
//      var result = rowResults[i];
//      activeRange.offset(result.rowNumber - 1, 0, 1).setBackground(result.color);
//      if(result.color == 'Red'){
//        activeRange.offset(result.rowNumber - 1, 0, 1, 1).setNote(result.message);
//      }
//    }
  }
  catch(err){
    quickLog('Error in updating SAP: ' + err + genericErrorMessage);  //PAUL
    logError('updateBackendMultipleTransactions() outer catch ' + err.message);
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
      'Can\'t update records',
      'There was an error updating SAP. You can check the Log sheet for technical details.',
      ui.ButtonSet.OK);
  }
}
