/*
* CloudSimple 3.0
   by Naveen Rokkam, Paul J. Modderman, Gavin P. Quinn - Mindset Consulting
   v3.0   20181205 - Open Source Software
*/
function removeFilters(filtersToRemove){
  var filterSet = JSON.parse(PropertiesService.getDocumentProperties().getProperty('FILTERS'));
  if(filterSet == null){return null;}

  //gotta decrement through it, because splice() will re-index the array
  for(var i = filterSet.filters.length - 1; i >= 0; i--){
    var removeFilter = filtersToRemove.indexOf(filterSet.filters[i].id);
    if(removeFilter > -1){
      filterSet.filters.splice(i, 1);
    }
  }

  PropertiesService.getDocumentProperties().setProperty('FILTERS', JSON.stringify(filterSet));

  return filtersToRemove;
}

function getFilterFields(filterIds){
  var filterSet = JSON.parse(PropertiesService.getDocumentProperties().getProperty('FILTERS'));
  if(filterSet == null){return null;}

  var filterFields = [];
  for(var i = 0; i < filterSet.filters.length; i++){
    if(filterIds.indexOf(filterSet.filters[i].id > -1)){
      if(filterSet.filters[i].edmType.indexOf('DateTime') > -1){
        filterSet.filters[i].value = 'datetime\'' + filterSet.filters[i].value + 'T00:00:00\'';
      }
      filterFields.push(filterSet.filters[i]);
    }
  }

  return filterFields;
}

function formatFilterFieldValue(field){
  if(field.edmType.indexOf('DateTime') > -1){
    return field.value;
  } else {
    return '\'' + field.value + '\'';
  }
}


function addTempFilter(filter){
  var filterSet = JSON.parse(PropertiesService.getDocumentProperties().getProperty('FILTERS'));
  var shouldAdd = true;
  if(filterSet == null){
    filterSet = {filters:[]};
  }

  //compare type, name, and value
  //If there's a duplicate filter, don't add it.
  if(filterSet != null){
    for(var i = 0; i < filterSet.filters.length; i++){
      if(filterSet.filters[i].name == filter.name &&
         filterSet.filters[i].type == filter.type &&
         filterSet.filters[i].value == filter.value){
        shouldAdd = false;
      }
    }
  }

  if(shouldAdd){
    //if we're saving a previously-created filter, the ID is populated and
    //we should remove the previously-created filter.
    if(filter.hasOwnProperty('id')){
      for(var i = 0; i < filterSet.filters.length; i++){
        if(filterSet.filters[i].id == filter.id){
          filterSet.filters.splice(i, 1);
        }
      }
    }else{
      filter.id = Utilities.formatDate(new Date(), 'GMT', 'yyyyMMddHHmmssSSS');
    }
    filterSet.filters.push(filter);
    PropertiesService.getDocumentProperties().setProperty('FILTERS', JSON.stringify(filterSet));
  }
}

function getFilters(){
  return JSON.parse(PropertiesService.getDocumentProperties().getProperty('FILTERS'));
}

function getEditFilterData(){
  return JSON.parse(CacheService.getDocumentCache().get('EDIT_FILTER'));
}

function showEditFilter(filterId){
  var cache = CacheService.getDocumentCache();
  var filterSet = JSON.parse(PropertiesService.getDocumentProperties().getProperty('FILTERS'));
  if(filterSet == null){return;}

  for(var i = 0; i < filterSet.filters.length; i++){
    if(filterSet.filters[i].id == filterId){
      cache.put('EDIT_FILTER', JSON.stringify(filterSet.filters[i]), 60);
      var html = HtmlService.createHtmlOutputFromFile('EditFilter')
        .setWidth(260)
        .setHeight(350);
      SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .showModalDialog(html, 'Edit filter');
    }
  }
}

function showFilterError(){
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    'Filter required',
    'This SAP object requires at least one filter to read data. Create one in the Filters tab.',
    ui.ButtonSet.OK);
}

function showAddFilterDialog(){
  var html = HtmlService.createHtmlOutputFromFile('AddFilter')
      .setWidth(260)
      .setHeight(350);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Add filter');
}
