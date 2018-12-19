/*
* CloudSimple 3.0
   by Naveen Rokkam, Paul J. Modderman, Gavin P. Quinn - Mindset Consulting
   v3.0   20181205 - Open Source Software
*/
//Constants. Merge this with the Constants.gs in the App Engine version of CloudShuttle
//Also note:
//Guess we'll have to do this through the old-school properties service stuff until
//we pull over the cache and properties mgmt script tools from the App Engine version
//Don't want to bother doing stuff twice.

/**
 * A global constant 'notice' text to include with each email
 * notification.
 */
var NOTICE = "The number of notifications this add-on produces are limited by the \
owner's available email quota; it will not send email notifications if the \
owner's daily email quota has been exceeded. Collaborators using this add-on on \
the same form will be able to adjust the notification settings, but will not be \
able to disable the notification triggers set by other collaborators.";

//Schedule repetition types
var repeatTypeHourly = 'HOURLY';
var repeatTypeDaily = 'DAILY';
var repeatTypeWeekly = 'WEEKLY';
var repeatTypeMonthly = 'MONTHLY';

var propReportSchedule = 'REPORT_SCHEDULE';
var propTriggerId = 'TRIGGER_ID_PROP';     // Store the id so we know what to delete
var propTriggerData = 'TRIGGER_DATA_PROP'; // Store this at the document level
var propLastRun = 'LAST_RUN_PROP';

/**
 * Function called by the trigger to execute the scheduled read.
 *
 */
function executeScheduledRead(){
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  // Check if the actions of the trigger require authorizations that have not
  // been supplied yet -- if so, warn the active user via email (if possible).
  // This check is required when using triggers with add-ons to maintain
  // functional triggers.
  if (authInfo.getAuthorizationStatus() ==
      ScriptApp.AuthorizationStatus.REQUIRED) {
    // Re-authorization is required. In this case, the user needs to be alerted
    // that they need to reauthorize; the normal trigger action is not
    // conducted, since it authorization needs to be provided first. Send at
    // most one 'Authorization Required' email a day, to avoid spamming users
    // of the add-on.
    sendReauthorizationRequest();
  } else {
    var selectFields = JSON.parse(PropertiesService.getUserProperties().getProperty('SELECT_FIELDS'));
    var filterIds = JSON.parse(PropertiesService.getUserProperties().getProperty('FILTER_IDS'));
    var existingSchedule = JSON.parse(PropertiesService.getUserProperties().getProperty(propReportSchedule));
    if(existingSchedule) {
      // Store the last run date / time string
      existingSchedule.lastRun = getFormattedDate(new Date(), Session.getScriptTimeZone());
      PropertiesService.getUserProperties().setProperty(propReportSchedule, JSON.stringify(existingSchedule));
      readService(selectFields, filterIds, false, existingSchedule.newSheet);
      sendScheduledMessage();
    } else {
      // Do nothing, we don't have an existing schedule.
    }
  }
}

/**
 * Called when the user needs to reauthorize. Sends the user of the
 * add-on an email explaining the need to reauthorize and provides
 * a link for the user to do so. Capped to send at most one email
 * a day to prevent spamming the users of the add-on.
 */
function sendReauthorizationRequest() {
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  var lastAuthEmailDate = settings.getProperty('lastAuthEmailDate');
  var today = new Date().toDateString();
  if (lastAuthEmailDate != today) {
    if (MailApp.getRemainingDailyQuota() > 0) {
      var template =
          HtmlService.createTemplateFromFile('AuthorizationEmail');
      template.url = authInfo.getAuthorizationUrl();
      template.notice = NOTICE;
      var message = template.evaluate();
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
          'Authorization Required',
          message.getContent(), {
            name: ADDON_TITLE,
            htmlBody: message.getContent()
          });
    }
    settings.setProperty('lastAuthEmailDate', today);
  }
}

/**
 * Delete all triggers for this doc. Should be used for debug purposes only.
 *
 */
function deleteAllTriggers() {
  // Delete the current trigger first
  deleteCurrentTrigger();
  // Remove all other triggers
  var allTriggers = ScriptApp.getProjectTriggers();
  // Loop over all triggers
  for (var i = 0; i < allTriggers.length; i++) {
    // Found the trigger and now delete it
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}

/**
 * @private
 *
 * Delete a trigger with the given unique ID.
 *
 * @param {object} triggerId Id for the trigger to be deleted
 * @return {bool} Whether or not a trigger was deleted
 *
 */
function deleteTriggerById_(triggerId) {
  var didDeleteTrigger = false;
  // Loop over all triggers.
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    // If the current trigger is the correct one, delete it.
    if (allTriggers[i].getUniqueId() == triggerId) {
      ScriptApp.deleteTrigger(allTriggers[i]);
      didDeleteTrigger = true;
      continue;
    }
  }
  return didDeleteTrigger;
}

function deleteCurrentTrigger() {
  var existingId = PropertiesService.getUserProperties().getProperty(propTriggerId);
  var result = false;
  if(existingId) {
    result = deleteTriggerById_(existingId);
  }

  // Remove all properties to ensure everything stays in sync. Just in case some
  // other app removed all triggers for this doc.
  PropertiesService.getUserProperties().deleteProperty(propTriggerId);
  PropertiesService.getUserProperties().deleteProperty(propReportSchedule);
  PropertiesService.getDocumentProperties().deleteProperty(propTriggerData);
  return result;
}

function createSchedule(schedule){
  // Only one trigger per user per type per document. We are using time based triggers,
  // if we have an existing time based trigger, remove it.
  deleteCurrentTrigger();

  // Save the schedule information the user properties
  PropertiesService.getUserProperties().setProperty(propReportSchedule, JSON.stringify(schedule));

  var triggerBuilder = ScriptApp.newTrigger('executeScheduledRead').timeBased();

  //These all need some better error handling
  switch(schedule.type) {
    case repeatTypeHourly:
      triggerBuilder.everyHours(schedule.repeatEveryHours);
      break;
    case repeatTypeDaily:
      triggerBuilder.everyDays(1)
             .atHour(schedule.hourSpan)
             .nearMinute(5);
      break;
    case repeatTypeWeekly:
      triggerBuilder.everyWeeks(1)
             .onWeekDay(getWeekDayFromString_(schedule.repeatDay))
             .atHour(schedule.hourSpan)
             .nearMinute(5);
      break;
    case repeatTypeMonthly:
      triggerBuilder.onMonthDay(schedule.repeatDayOfMonth)
             .atHour(schedule.hourSpan)
             .nearMinute(5);
      break;
  }

  var trigger = triggerBuilder.create();
  PropertiesService.getUserProperties().setProperty(propTriggerId, trigger.getUniqueId());
  var triggerProperties = {creator: Session.getActiveUser().getEmail()};
  PropertiesService.getDocumentProperties().setProperty(propTriggerData, JSON.stringify(triggerProperties));
}

/**
 * @private
 *
 * Turns a string into an enum for the onWeekDay function of the trigger builder.
 *
 * @param {string} weekdayString Day of week string
 * @return {int} Day of week enum
 *
 */
function getWeekDayFromString_(weekdayString) {
  var result = ScriptApp.WeekDay.FRIDAY;
  switch(weekdayString) {
    case "SUNDAY":
      result = ScriptApp.WeekDay.SUNDAY;
      break;
    case "MONDAY":
      result = ScriptApp.WeekDay.MONDAY;
      break;
    case "TUESDAY":
      result = ScriptApp.WeekDay.TUESDAY;
      break;
    case "WEDNESDAY":
      result = ScriptApp.WeekDay.WEDNESDAY;
      break;
    case "THURSDAY":
      result = ScriptApp.WeekDay.THURSDAY;
      break;
    case "FRIDAY":
      result = ScriptApp.WeekDay.FRIDAY;
      break;
    case "SATURDAY":
      result = ScriptApp.WeekDay.SATURDAY;
      break;

  }
  return result;
}

/**
 * Returns trigger related properties to the view so it can show the correct messaging
 * and buttons to the user.
 *
 * @return {string} JSON string of object that includes a message (string) and creator (bool).
 *
 */
function getTriggerProperties() {
  var data;
  var triggerPropsString = PropertiesService.getDocumentProperties().getProperty(propTriggerData);
  if(triggerPropsString) {
    var triggerProps = JSON.parse(triggerPropsString);
    if(triggerProps.creator === Session.getActiveUser().getEmail()) {
      var existingSchedule = JSON.parse(PropertiesService.getUserProperties().getProperty(propReportSchedule));
      data = {message: "You have an existing recurring report scheduled to run " + getScheduleString(), creator: true};
    } else {
      data = {message: triggerProps.creator +" already has a recurring report scheduled. Only one scheduled report can exist per document.", creator: false};
    }
  }
  return JSON.stringify(data);
}

function getScheduleString() {
  var result = "";
  var existingSchedule = JSON.parse(PropertiesService.getUserProperties().getProperty(propReportSchedule));
  if(existingSchedule) {
    switch(existingSchedule.type) {
      case "HOURLY":
        result = " every " + existingSchedule.repeatEveryHours + " hours. ";
        break;
      case "WEEKLY":
        result = " weekly.";
        break;
      case "DAILY":
        result = " daily.";
        break;
      case "MONTHLY":
        result = " monthly.";
        break;
    }
    if(existingSchedule.hasOwnProperty("lastRun")) result += " Last run " + existingSchedule.lastRun + ".";
  }
  return result;
}

function getLastRunDate() {
  var result = "";
  var existingSchedule = JSON.parse(PropertiesService.getUserProperties().getProperty(propReportSchedule));
  if(existingSchedule && existingSchedule.hasOwnProperty("lastRun")) {
    result = existingSchedule.lastRun;
  }else {
    result = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  }
  return result;
}

function getFormattedDate(date, timezone) {
  var result = "";
  var existingSchedule = JSON.parse(PropertiesService.getUserProperties().getProperty(propReportSchedule));
  if(existingSchedule && existingSchedule.hasOwnProperty("type")) {
    switch(existingSchedule.type) {
      case "HOURLY":
        result = Utilities.formatDate(date, timezone, "EEE MMM d, HH:mm");
        break;
      case "WEEKLY":
        result = Utilities.formatDate(date, timezone, "EEE MMM d, yyyy");
        break;
      case "DAILY":
        result = Utilities.formatDate(date, timezone, "EEE MMM d, HH:mm");
        break;
      case "MONTHLY":
        result = Utilities.formatDate(date, timezone, "MMM d, yyyy");
        break;
    }
  } else {
    result = Utilities.formatDate(date, timezone, 'yyyyMMddHHmmss');
  }
  return result;
}

function getScheduleData() {
  var existingSchedule = PropertiesService.getUserProperties().getProperty(propReportSchedule);
  return existingSchedule;
}

function hasExistingSchedule() {
  if(getScheduleData()) {
    return true;
  } else {
    return false;
  }
}

/**
 * Send out e-mail message after scheduled read is complete.
 *
 */
function sendScheduledMessage(){
  var existingSchedule = JSON.parse(PropertiesService.getUserProperties().getProperty(propReportSchedule));
  var ssUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var message = existingSchedule.message + '\n\n' + ssUrl;
  //distributionList messageSubject message

  var options = {};
  options.name = 'CloudShuttle Scheduled Report';
  options.noReply = true;

  MailApp.sendEmail(existingSchedule.distributionList, existingSchedule.messageSubject, message, options);
}
