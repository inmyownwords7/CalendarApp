/// <reference types="google-apps-script" />
/// <reference path="./types/types.d.ts" />

import {
  createCalendarForRequester,
  getDebugSummary,
  doGet,
  getCurrentUserEmail,
  getLoggerHealth,
  getIdentitySummary,
  listAllEventsForIdentifier,
  listAllEventsForCurrentUser,
  listAvailableGreyBoxIds,
  requestCalendarForCurrentUser,
  requestCalendarByGreyBoxId,
  start,
  syncAllTrackedTimeToSheet,
  syncTrackedTimeToSheet,
} from "./calendar";

/**
 * GAS entrypoint.
 * Export a single namespace on globalThis because top-level GAS bridge
 * function names can collide with plain global assignments.
 */
(globalThis as Record<string, unknown>).__calendarApp = {
  hello: () => {
    Logger.log("Hello from GAS bundle!");
  },
  start,
  doGet,
  getCurrentUserEmail,
  getLoggerHealth,
  getIdentitySummary,
  getDebugSummary,
  listAllEventsForIdentifier,
  listAllEventsForCurrentUser,
  createCalendarForRequester,
  requestCalendarForCurrentUser,
  requestCalendarByGreyBoxId,
  listAvailableGreyBoxIds,
  syncAllTrackedTimeToSheet,
  syncTrackedTimeToSheet,
};
