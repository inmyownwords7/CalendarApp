/// <reference types="google-apps-script" />
/// <reference path="../types/logger-lib.d.ts" />

import {
  createCalendarForRequester as createCalendarForRequesterHandler,
  getDebugSummary as getDebugSummaryHandler,
  doGet as doGetHandler,
  getCurrentUserEmail as getCurrentUserEmailHandler,
  getLoggerHealth as getLoggerHealthHandler,
  getIdentitySummary as getIdentitySummaryHandler,
  listAllEventsForCalendarId as listAllEventsForCalendarIdHandler,
  listAllEventsForCalendarUrl as listAllEventsForCalendarUrlHandler,
  listAllEventsForEmail as listAllEventsForEmailHandler,
  listAllEventsForGreyBoxId as listAllEventsForGreyBoxIdHandler,
  listAllEventsForIdentifier as listAllEventsForIdentifierHandler,
  listAllEventsForCurrentUser as listAllEventsForCurrentUserHandler,
  listAvailableGreyBoxIds as listAvailableGreyBoxIdsHandler,
  requestCalendarForCurrentUser as requestCalendarForCurrentUserHandler,
  requestCalendarByGreyBoxId as requestCalendarByGreyBoxIdHandler,
  start as startHandler,
  syncAllTrackedTimeToSheet as syncAllTrackedTimeToSheetHandler,
  syncTrackedTimeToSheet as syncTrackedTimeToSheetHandler,
} from "../services/calendar";

/**
 * GAS entrypoint.
 * Export a single namespace on globalThis because top-level GAS bridge
 * function names can collide with plain global assignments.
 */
const calendarAppBridge = {
  hello: () => {
    Logger.log("Hello from GAS bundle!");
  },
  start: startHandler,
  doGet: doGetHandler,
  getCurrentUserEmail: getCurrentUserEmailHandler,
  getLoggerHealth: getLoggerHealthHandler,
  getIdentitySummary: getIdentitySummaryHandler,
  getDebugSummary: getDebugSummaryHandler,
  listAllEventsForCalendarId: listAllEventsForCalendarIdHandler,
  listAllEventsForCalendarUrl: listAllEventsForCalendarUrlHandler,
  listAllEventsForEmail: listAllEventsForEmailHandler,
  listAllEventsForGreyBoxId: listAllEventsForGreyBoxIdHandler,
  listAllEventsForIdentifier: listAllEventsForIdentifierHandler,
  listAllEventsForCurrentUser: listAllEventsForCurrentUserHandler,
  createCalendarForRequester: createCalendarForRequesterHandler,
  requestCalendarForCurrentUser: requestCalendarForCurrentUserHandler,
  requestCalendarByGreyBoxId: requestCalendarByGreyBoxIdHandler,
  listAvailableGreyBoxIds: listAvailableGreyBoxIdsHandler,
  syncAllTrackedTimeToSheet: syncAllTrackedTimeToSheetHandler,
  syncTrackedTimeToSheet: syncTrackedTimeToSheetHandler,
};

(globalThis as Record<string, unknown>).__calendarApp = calendarAppBridge;
Object.assign(globalThis as Record<string, unknown>, {
  _hello: calendarAppBridge.hello,
  _start: calendarAppBridge.start,
  _doGet: calendarAppBridge.doGet,
  _getCurrentUserEmail: calendarAppBridge.getCurrentUserEmail,
  _getLoggerHealth: calendarAppBridge.getLoggerHealth,
  _getIdentitySummary: calendarAppBridge.getIdentitySummary,
  _getDebugSummary: calendarAppBridge.getDebugSummary,
  _listAllEventsForCalendarId: calendarAppBridge.listAllEventsForCalendarId,
  _listAllEventsForCalendarUrl: calendarAppBridge.listAllEventsForCalendarUrl,
  _listAllEventsForEmail: calendarAppBridge.listAllEventsForEmail,
  _listAllEventsForGreyBoxId: calendarAppBridge.listAllEventsForGreyBoxId,
  _listAllEventsForIdentifier: calendarAppBridge.listAllEventsForIdentifier,
  _listAllEventsForCurrentUser: calendarAppBridge.listAllEventsForCurrentUser,
  _createCalendarForRequester: calendarAppBridge.createCalendarForRequester,
  _requestCalendarForCurrentUser: calendarAppBridge.requestCalendarForCurrentUser,
  _requestCalendarByGreyBoxId: calendarAppBridge.requestCalendarByGreyBoxId,
  _listAvailableGreyBoxIds: calendarAppBridge.listAvailableGreyBoxIds,
  _syncAllTrackedTimeToSheet: calendarAppBridge.syncAllTrackedTimeToSheet,
  _syncTrackedTimeToSheet: calendarAppBridge.syncTrackedTimeToSheet,
});
