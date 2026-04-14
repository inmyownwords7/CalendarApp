/// <reference types="google-apps-script" />
/// <reference path="./types/types.d.ts" />

import {
  createCalendarForRequester,
  doGet,
  getCurrentUserEmail,
  listAvailableGreyBoxIds,
  requestCalendarForCurrentUser,
  requestCalendarByGreyBoxId,
  start,
  syncTrackedTimeToSheet,
} from "./calendar";

/**
 * GAS entrypoint.
 * Export global functions by assigning to globalThis because the build emits an IIFE bundle.
 */
(globalThis as Record<string, unknown>).hello = () => {
  Logger.log("Hello from GAS bundle!");
};

(globalThis as Record<string, unknown>).start = start;
(globalThis as Record<string, unknown>).doGet = doGet;
(globalThis as Record<string, unknown>).getCurrentUserEmail = getCurrentUserEmail;
(globalThis as Record<string, unknown>).createCalendarForRequester =
  createCalendarForRequester;
(globalThis as Record<string, unknown>).requestCalendarForCurrentUser =
  requestCalendarForCurrentUser;
(globalThis as Record<string, unknown>).requestCalendarByGreyBoxId =
  requestCalendarByGreyBoxId;
(globalThis as Record<string, unknown>).listAvailableGreyBoxIds =
  listAvailableGreyBoxIds;
(globalThis as Record<string, unknown>).syncTrackedTimeToSheet =
  syncTrackedTimeToSheet;
