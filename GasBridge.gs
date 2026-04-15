function getCalendarAppBundle_() {
  return globalThis.__calendarApp;
}

function doGet() {
  return getCalendarAppBundle_().doGet();
}

function start() {
  return getCalendarAppBundle_().start();
}

function getCurrentUserEmail() {
  return getCalendarAppBundle_().getCurrentUserEmail();
}

function getCurrentUserString() {
  return getCalendarAppBundle_().getCurrentUserEmail();
}

function getIdentitySummary() {
  return getCalendarAppBundle_().getIdentitySummary();
}

function getDebugSummary() {
  return getCalendarAppBundle_().getDebugSummary();
}

function getLoggerHealth() {
  return getCalendarAppBundle_().getLoggerHealth();
}

function requestCalendarForCurrentUser() {
  return getCalendarAppBundle_().requestCalendarForCurrentUser();
}

function requestCalendarByGreyBoxId(greyBoxId) {
  return getCalendarAppBundle_().requestCalendarByGreyBoxId(greyBoxId);
}

function createCalendarForRequester(title, fallbackEmail) {
  return getCalendarAppBundle_().createCalendarForRequester(title, fallbackEmail);
}

function listAvailableGreyBoxIds() {
  return getCalendarAppBundle_().listAvailableGreyBoxIds();
}

function listAllEventsForCurrentUser() {
  return getCalendarAppBundle_().listAllEventsForCurrentUser();
}

function listAllEventsForIdentifier(identifier) {
  return getCalendarAppBundle_().listAllEventsForIdentifier(identifier);
}

function syncTrackedTimeToSheet() {
  return getCalendarAppBundle_().syncTrackedTimeToSheet();
}

function syncAllTrackedTimeToSheet() {
  return getCalendarAppBundle_().syncAllTrackedTimeToSheet();
}
