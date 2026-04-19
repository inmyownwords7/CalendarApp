/// <reference types="google-apps-script" />
/// <reference path="../types/logger-lib.d.ts" />

type GreyBoxRecord = {
  greyBoxId: string;
  email: string;
  calendarTitle: string;
};

type CalendarRequestResponse = {
  ok: true;
  greyBoxId: string;
  email: string;
  calendarId: string;
  calendarUrl: string;
  title: string;
};

type EventTrackingContext = {
  greyBoxId: string;
  requesterEmail: string;
  calendarId: string;
  calendarTitle: string;
  calendarUrl: string;
};

type TimeEntryInput = {
  summary: string;
  startIso: string;
  endIso: string;
  description?: string;
};

type IdentitySummary = {
  activeEmail: string;
  effectiveEmail: string;
  isRequesterDeployerSame: boolean;
};

type DebugSummary = {
  activeEmail: string;
  effectiveEmail: string;
  isRequesterDeployerSame: boolean;
  mandateEmail: string;
  greyBoxId: string;
  calendarTitle: string;
  storedCalendarId: string;
};

type LoggerHealthSummary = {
  ok: boolean;
  spreadsheetId: string;
  operationalSheet: string;
  networkSheet: string;
  operationalSheetExists: boolean;
  networkSheetExists: boolean;
  loggerInitialized: boolean;
  failureStage?: string;
  errorMessage?: string;
};

export type {
  GreyBoxRecord,
  CalendarRequestResponse,
  EventTrackingContext,
  TimeEntryInput,
  IdentitySummary,
  DebugSummary,
  LoggerHealthSummary,
};
