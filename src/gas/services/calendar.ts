/// <reference types="google-apps-script" />
/// <reference path="../types/logger-lib.d.ts" />

import {
  getLoggerHealthInternal,
  logFunctionError,
  logFunctionStart,
  logFunctionSuccess,
  logInit,
  loggerError,
  loggerHttp,
  loggerInfo,
} from "./logger";
import {
  buildCalendarUrl,
  ensureCalendarForRecord,
  getCalendarService,
  getEventTrackingContextForCurrentUserInternal,
  getGreyBoxDirectory,
  getGreyBoxRecord,
  isGoogleCalendarId,
  getMandateMatchByEmail,
  getMandateMatchByIdentifier,
  getStoredCalendarId,
  normalizeEmail,
  shareCalendarWithUser,
} from "./records";
import {
  buildTrackedTimeRowsForRecord,
  listAllEventsForCalendar,
  parseRequiredDate,
  recordGreyBoxSheetRequest,
  upsertTrackedTimeRows,
  writeValidatedGreyBoxSheet,
} from "./tracking";
import type {
  CalendarRequestResponse,
  DebugSummary,
  EventTrackingContext,
  GreyBoxRecord,
  IdentitySummary,
  LoggerHealthSummary,
  TimeEntryInput,
} from "./types";

export type {
  CalendarRequestResponse,
  DebugSummary,
  EventTrackingContext,
  GreyBoxRecord,
  IdentitySummary,
  LoggerHealthSummary,
  TimeEntryInput,
} from "./types";

// Runtime entrypoints exposed through the GAS bridge.
function start(): void {
  logInit();

  loggerInfo({
    message: "Identity sync started",
    meta: { source: "identityBuild" },
  });

  loggerHttp({
    system: "SLACK",
    method: "GET",
    url: { raw: "https://slack.com/api/users.list" },
    status: 200,
    durationMs: 180,
  });
}

function doGet(): GoogleAppsScript.HTML.HtmlOutput {
  return HtmlService.createHtmlOutputFromFile("index").setTitle(
    "Calendar Request Manager"
  );
}

// Web app handlers used directly by `index.html`.
// Page load:
// - getIdentitySummary
// Buttons:
// - requestCalendarForCurrentUser
// - listAllEventsForEmail / listAllEventsForGreyBoxId /
//   listAllEventsForCalendarId / listAllEventsForCalendarUrl /
//   listAllEventsForIdentifier
// - getCurrentUserEmail
// - getLoggerHealth
// - getDebugSummary
function getCurrentUserEmail(): string {
  logInit();

  const email = Session.getActiveUser().getEmail() || "";

  const rec = loggerInfo({
    message: email
      ? "Fetched current user email"
      : "Current user email unavailable",
    meta: { email },
  });

  Logger.log("logId=" + rec.logId);
  return email;
}

function getLoggerHealth(): LoggerHealthSummary {
  return getLoggerHealthInternal();
}

function getIdentitySummary(): IdentitySummary {
  logInit();
  const startLog = logFunctionStart("getIdentitySummary");

  try {
    const activeEmail = normalizeEmail(Session.getActiveUser().getEmail() || "");
    const effectiveEmail = normalizeEmail(Session.getEffectiveUser().getEmail() || "");
    const isRequesterDeployerSame = Boolean(
      activeEmail && effectiveEmail && activeEmail === effectiveEmail
    );

    logFunctionSuccess(
      "getIdentitySummary",
      {
        activeEmail: activeEmail || "(blank)",
        effectiveEmail: effectiveEmail || "(blank)",
        isRequesterDeployerSame,
      },
      startLog.logId
    );

    return {
      activeEmail,
      effectiveEmail,
      isRequesterDeployerSame,
    };
  } catch (error) {
    logFunctionError("getIdentitySummary", error, {}, startLog.logId);
    throw error;
  }
}

function getDebugSummary(): DebugSummary {
  logInit();
  const startLog = logFunctionStart("getDebugSummary");

  try {
    const activeEmail = normalizeEmail(Session.getActiveUser().getEmail() || "");
    const effectiveEmail = normalizeEmail(Session.getEffectiveUser().getEmail() || "");

    if (!activeEmail) {
      throw new Error("Could not determine the signed-in user email.");
    }

    const mandateMatch = getMandateMatchByEmail(activeEmail);
    const storedCalendarId = getStoredCalendarId(mandateMatch.record.greyBoxId);

    const result: DebugSummary = {
      activeEmail,
      effectiveEmail,
      isRequesterDeployerSame: Boolean(
        activeEmail && effectiveEmail && activeEmail === effectiveEmail
      ),
      mandateEmail: mandateMatch.record.email,
      greyBoxId: mandateMatch.record.greyBoxId,
      calendarTitle: mandateMatch.record.calendarTitle,
      storedCalendarId,
    };

    logFunctionSuccess("getDebugSummary", result, startLog.logId);
    return result;
  } catch (error) {
    logFunctionError("getDebugSummary", error, {}, startLog.logId);
    throw error;
  }
}

function listAvailableGreyBoxIds(): string[] {
  return Object.keys(getGreyBoxDirectory()).sort();
}

function requestCalendarForCurrentUser(): CalendarRequestResponse {
  logInit();
  const startLog = logFunctionStart("requestCalendarForCurrentUser");
  const requesterEmail = normalizeEmail(Session.getActiveUser().getEmail());

  try {
    if (!requesterEmail) {
      throw new Error("Could not determine the signed-in user email.");
    }

    const mandateMatch = getMandateMatchByEmail(requesterEmail);
    const context = ensureCalendarForRecord(mandateMatch.record, requesterEmail);
    recordGreyBoxSheetRequest(context);

    loggerInfo({
      message: "Resolved mandate request for signed-in user",
      parentLogId: startLog.logId,
      meta: {
        action: "RESOLVE_MANDATE_REQUEST",
        requesterEmail: context.requesterEmail,
        greyBoxId: context.greyBoxId,
        calendarId: context.calendarId,
      },
    });

    const result: CalendarRequestResponse = {
      ok: true,
      greyBoxId: context.greyBoxId,
      email: context.requesterEmail,
      calendarId: context.calendarId,
      title: context.calendarTitle,
      calendarUrl: context.calendarUrl,
    };
    logFunctionSuccess(
      "requestCalendarForCurrentUser",
      {
        requesterEmail: context.requesterEmail,
        greyBoxId: context.greyBoxId,
        calendarId: context.calendarId,
      },
      startLog.logId
    );
    return result;
  } catch (error) {
    logFunctionError(
      "requestCalendarForCurrentUser",
      error,
      { requesterEmail: requesterEmail || "(blank)" },
      startLog.logId
    );
    throw error;
  }
}

function requestCalendarByGreyBoxId(
  greyBoxId: string
): CalendarRequestResponse {
  logInit();
  const startLog = logFunctionStart("requestCalendarByGreyBoxId", { greyBoxId });

  try {
    const record = getGreyBoxRecord(greyBoxId);

    loggerInfo({
      message: "Resolved grey-box request",
      parentLogId: startLog.logId,
      meta: {
        action: "RESOLVE_GREY_BOX_REQUEST",
        greyBoxId: record.greyBoxId,
        email: record.email,
      },
    });

    const result = createCalendarForEmail(record.calendarTitle, record.email, {
      requestType: "grey_box_id",
      greyBoxId: record.greyBoxId,
    });
    logFunctionSuccess(
      "requestCalendarByGreyBoxId",
      { greyBoxId: record.greyBoxId, calendarId: result.calendarId },
      startLog.logId
    );
    return result;
  } catch (error) {
    logFunctionError(
      "requestCalendarByGreyBoxId",
      error,
      { greyBoxId },
      startLog.logId
    );
    throw error;
  }
}

function createCalendarForRequester(
  title: string,
  fallbackEmail?: string | null
) {
  logInit();
  const startLog = logFunctionStart("createCalendarForRequester", {
    title,
    fallbackEmail: fallbackEmail || "",
  });

  try {
    const detected = Session.getActiveUser().getEmail();
    const email = normalizeEmail(detected || fallbackEmail || "");

    if (!email) {
      throw new Error("Could not determine your email. Please type it.");
    }

    const result = createCalendarForEmail(title, email, { requestType: "direct_email" });
    logFunctionSuccess(
      "createCalendarForRequester",
      { email, calendarId: result.calendarId },
      startLog.logId
    );
    return result;
  } catch (error) {
    logFunctionError(
      "createCalendarForRequester",
      error,
      { title, fallbackEmail: fallbackEmail || "" },
      startLog.logId
    );
    throw error;
  }
}

function getEventTrackingContextForCurrentUser(): EventTrackingContext {
  logInit();
  const startLog = logFunctionStart("getEventTrackingContextForCurrentUser");
  try {
    const context = getEventTrackingContextForCurrentUserInternal();
    logFunctionSuccess(
      "getEventTrackingContextForCurrentUser",
      {
        requesterEmail: context.requesterEmail,
        greyBoxId: context.greyBoxId,
        calendarId: context.calendarId,
      },
      startLog.logId
    );
    return context;
  } catch (error) {
    logFunctionError(
      "getEventTrackingContextForCurrentUser",
      error,
      {},
      startLog.logId
    );
    throw error;
  }
}

function listEvents(calendarId: string) {
  logInit();
  const startLog = logFunctionStart("listEvents", { calendarId });
  const calendar = getCalendarService();
  try {
    const response = calendar.Events.list(calendarId);
    logFunctionSuccess(
      "listEvents",
      { calendarId, eventCount: (response.items || []).length },
      startLog.logId
    );
    return response;
  } catch (error) {
    logFunctionError("listEvents", error, { calendarId }, startLog.logId);
    throw error;
  }
}

function listAllEventsForCurrentUser() {
  logInit();
  const context = getEventTrackingContextForCurrentUserInternal();
  const events = listAllEventsForCalendar(context.calendarId);
  writeValidatedGreyBoxSheet(context.greyBoxId, events);
  return events;
}

function listAllEventsForEmail(email: string) {
  logInit();
  const normalizedEmail = normalizeEmail(email);

  if (!normalizedEmail) {
    throw new Error("Email is required.");
  }

  const match = getMandateMatchByEmail(normalizedEmail);
  return listAllEventsForResolvedRecord(match.record);
}

function listAllEventsForGreyBoxId(greyBoxId: string) {
  logInit();
  const normalizedGreyBoxId = String(greyBoxId || "").trim();

  if (!normalizedGreyBoxId) {
    throw new Error("Grey Box ID is required.");
  }

  const match = getMandateMatchByIdentifier(normalizedGreyBoxId);
  return listAllEventsForResolvedRecord(match.record);
}

function listAllEventsForCalendarId(calendarId: string) {
  logInit();
  const normalizedCalendarId = String(calendarId || "").trim();

  if (!normalizedCalendarId) {
    throw new Error("Calendar ID is required.");
  }

  const events = listAllEventsForCalendar(normalizedCalendarId);
  let matchedRecord: GreyBoxRecord | null = null;

  try {
    matchedRecord = getMandateMatchByIdentifier(normalizedCalendarId).record;
  } catch (_error) {
    matchedRecord = null;
  }

  if (matchedRecord) {
    writeValidatedGreyBoxSheet(matchedRecord.greyBoxId, events);
  }

  return events;
}

function listAllEventsForCalendarUrl(calendarUrl: string) {
  logInit();
  const calendarId = extractCalendarIdFromUrl(calendarUrl);
  return listAllEventsForCalendarId(calendarId);
}

function listAllEventsForIdentifier(identifier: string) {
  logInit();
  const normalizedIdentifier = String(identifier || "").trim();

  if (!normalizedIdentifier) {
    throw new Error("Identifier is required.");
  }

  if (isGoogleCalendarId(normalizedIdentifier)) {
    const events = listAllEventsForCalendar(normalizedIdentifier);
    let matchedRecord: GreyBoxRecord | null = null;

    try {
      matchedRecord = getMandateMatchByIdentifier(normalizedIdentifier).record;
    } catch (_error) {
      matchedRecord = null;
    }

    if (matchedRecord) {
      writeValidatedGreyBoxSheet(matchedRecord.greyBoxId, events);
    }

    return events;
  }

  const match = getMandateMatchByIdentifier(normalizedIdentifier);
  const calendarId = getStoredCalendarId(match.record.greyBoxId);

  if (!calendarId) {
    throw new Error(
      `No stored calendar found for ${match.record.greyBoxId}. Request or create the calendar first.`
    );
  }

  const events = listAllEventsForCalendar(calendarId);
  writeValidatedGreyBoxSheet(match.record.greyBoxId, events);
  return events;
}

// Additional exported service functions that are not wired to the current UI
// buttons but remain part of the GAS service surface.
function listTrackedEventsForCurrentUser() {
  logInit();
  const startLog = logFunctionStart("listTrackedEventsForCurrentUser");
  const context = getEventTrackingContextForCurrentUserInternal();
  const calendar = getCalendarService();

  loggerInfo({
    message: "Listing tracked events for current user",
    meta: {
      action: "LIST_TRACKED_EVENTS",
      requesterEmail: context.requesterEmail,
      greyBoxId: context.greyBoxId,
      calendarId: context.calendarId,
    },
  });

  const response = calendar.Events.list(context.calendarId, {
    singleEvents: true,
    orderBy: "startTime",
    timeMin: new Date(0).toISOString(),
    maxResults: 250,
  });
  logFunctionSuccess(
    "listTrackedEventsForCurrentUser",
    {
      requesterEmail: context.requesterEmail,
      greyBoxId: context.greyBoxId,
      calendarId: context.calendarId,
      eventCount: (response.items || []).length,
    },
    startLog.logId
  );
  return response;
}

function insertEvent(
  calendarId: string,
  event: GoogleAppsScript.Calendar.Schema.Event
) {
  logInit();
  const calendar = getCalendarService();

  try {
    const ev = calendar.Events.insert(event, calendarId);
    loggerInfo({
      message: "Event inserted successfully",
      meta: {
        action: "INSERT_EVENT",
        calendarId,
        eventId: ev.id,
        summary: event.summary,
      },
    });
    return ev;
  } catch (error) {
    loggerError({
      message: "Failed to insert event",
      error,
      meta: {
        action: "INSERT_EVENT",
        calendarId,
        summary: event.summary,
      },
    });
    throw error;
  }
}

function recordTrackedTimeForCurrentUser(input: TimeEntryInput) {
  logInit();
  const startLog = logFunctionStart("recordTrackedTimeForCurrentUser", {
    summary: input.summary || "",
    startIso: input.startIso,
    endIso: input.endIso,
  });
  const context = getEventTrackingContextForCurrentUserInternal();
  const startAt = parseRequiredDate(input.startIso, "startIso");
  const endAt = parseRequiredDate(input.endIso, "endIso");

  if (endAt.getTime() <= startAt.getTime()) {
    const error = new Error("endIso must be after startIso.");
    logFunctionError(
      "recordTrackedTimeForCurrentUser",
      error,
      {
        requesterEmail: context.requesterEmail,
        greyBoxId: context.greyBoxId,
        startIso: input.startIso,
        endIso: input.endIso,
      },
      startLog.logId
    );
    throw error;
  }

  const event: GoogleAppsScript.Calendar.Schema.Event = {
    summary: (input.summary || "").trim() || "Tracked Time",
    description: (input.description || "").trim(),
    start: { dateTime: startAt.toISOString() },
    end: { dateTime: endAt.toISOString() },
  };

  loggerInfo({
    message: "Recording tracked time for current user",
    meta: {
      action: "RECORD_TRACKED_TIME",
      requesterEmail: context.requesterEmail,
      greyBoxId: context.greyBoxId,
      calendarId: context.calendarId,
      summary: event.summary,
      startIso: event.start?.dateTime,
      endIso: event.end?.dateTime,
    },
  });

  const insertedEvent = insertEvent(context.calendarId, event);
  logFunctionSuccess(
    "recordTrackedTimeForCurrentUser",
    {
      requesterEmail: context.requesterEmail,
      greyBoxId: context.greyBoxId,
      calendarId: context.calendarId,
      eventId: insertedEvent.id || "",
    },
    startLog.logId
  );
  return insertedEvent;
}

function syncTrackedTimeToSheet() {
  logInit();
  const startLog = logFunctionStart("syncTrackedTimeToSheet");
  const syncedAt = new Date();
  const context = getEventTrackingContextForCurrentUserInternal();
  const events = listAllEventsForCalendar(context.calendarId);
  writeValidatedGreyBoxSheet(context.greyBoxId, events);

  const rows = buildTrackedTimeRowsForRecord(
    {
      greyBoxId: context.greyBoxId,
      email: context.requesterEmail,
      calendarTitle: context.calendarTitle,
    },
    context.calendarId,
    events,
    syncedAt
  );

  return {
    ...upsertTrackedTimeRows(rows, startLog.logId, "SYNC_TRACKED_TIME"),
    greyBoxId: context.greyBoxId,
    email: context.requesterEmail,
    calendarId: context.calendarId,
    syncedAt: syncedAt.toISOString(),
  };
}

function syncAllTrackedTimeToSheet() {
  logInit();
  const startLog = logFunctionStart("syncAllTrackedTimeToSheet");
  const records = Object.values(getGreyBoxDirectory());
  const syncedAt = new Date();
  const rows: Array<Array<string | number>> = [];

  for (const record of records) {
    const calendarId = getStoredCalendarId(record.greyBoxId);

    if (!calendarId) {
      continue;
    }

    try {
      const events = listAllEventsForCalendar(calendarId);
      writeValidatedGreyBoxSheet(record.greyBoxId, events);
      rows.push(...buildTrackedTimeRowsForRecord(record, calendarId, events, syncedAt));
    } catch (error) {
      loggerError({
        message: "Failed to sync tracked calendar events",
        error,
        meta: {
          action: "SYNC_ALL_TRACKED_TIME",
          greyBoxId: record.greyBoxId,
          email: record.email,
          calendarId,
        },
      });
    }
  }

  const result = {
    ...upsertTrackedTimeRows(rows, startLog.logId, "SYNC_ALL_TRACKED_TIME"),
    syncedAt: syncedAt.toISOString(),
  };

  logFunctionSuccess("syncAllTrackedTimeToSheet", result, startLog.logId);
  return result;
}

function createCalendarForEmail(
  title: string,
  email: string,
  meta: Record<string, unknown> = {}
): CalendarRequestResponse {
  logInit();
  const startMs = Date.now();
  const startLog = logFunctionStart("createCalendarForEmail", { title, email, ...meta });
  const normalizedTitle = (title || "").trim();
  const normalizedEmail = normalizeEmail(email);

  try {
    if (!normalizedTitle) {
      throw new Error("Calendar title required.");
    }

    if (!normalizedEmail) {
      throw new Error("Target email required.");
    }

    const greyBoxId = String(meta.greyBoxId || "");
    const context = greyBoxId
      ? ensureCalendarForRecord(
          {
            greyBoxId,
            email: normalizedEmail,
            calendarTitle: normalizedTitle,
          },
          normalizedEmail
        )
      : (() => {
          const calendar = getCalendarService();
          const cal = calendar.Calendars.insert({
            summary: normalizedTitle,
            timeZone: Session.getScriptTimeZone(),
          });

          if (!cal.id) {
            throw new Error("Calendar creation failed: No calendar ID returned.");
          }

          shareCalendarWithUser(cal.id, normalizedEmail);

          return {
            greyBoxId: "",
            requesterEmail: normalizedEmail,
            calendarId: cal.id,
            calendarTitle: normalizedTitle,
            calendarUrl: buildCalendarUrl(cal.id),
          };
        })();

    const mainLog = loggerInfo({
      message: "Calendar created successfully",
      parentLogId: startLog.logId,
      meta: {
        action: "CREATE_CALENDAR",
        email: normalizedEmail,
        title: normalizedTitle,
        calendarId: context.calendarId,
        durationMs: Date.now() - startMs,
        ...meta,
      },
    });

    loggerHttp({
      parentLogId: mainLog.logId,
      system: "GOOGLE",
      method: "POST",
      endpoint: "calendar/v3/calendars",
      url: { raw: "https://www.googleapis.com/calendar/v3/calendars" },
      status: 200,
      durationMs: Date.now() - startMs,
      meta: {
        email: normalizedEmail,
        title: normalizedTitle,
        calendarId: context.calendarId,
        ...meta,
      },
    });

    logFunctionSuccess(
      "createCalendarForEmail",
      {
        email: normalizedEmail,
        title: normalizedTitle,
        greyBoxId: context.greyBoxId,
        calendarId: context.calendarId,
        durationMs: Date.now() - startMs,
      },
      startLog.logId
    );

    return {
      ok: true,
      greyBoxId: context.greyBoxId,
      email: normalizedEmail,
      calendarId: context.calendarId,
      title: normalizedTitle,
      calendarUrl: context.calendarUrl,
    };
  } catch (error) {
    logFunctionError(
      "createCalendarForEmail",
      error,
      {
        email: normalizedEmail || "(blank)",
        title: normalizedTitle || "(blank)",
        durationMs: Date.now() - startMs,
        ...meta,
      },
      startLog.logId
    );
    throw error;
  }
}

function listAllEventsForResolvedRecord(record: GreyBoxRecord) {
  const calendarId = getStoredCalendarId(record.greyBoxId);

  if (!calendarId) {
    throw new Error(
      `No stored calendar found for ${record.greyBoxId}. Request or create the calendar first.`
    );
  }

  const events = listAllEventsForCalendar(calendarId);
  writeValidatedGreyBoxSheet(record.greyBoxId, events);
  return events;
}

function extractCalendarIdFromUrl(calendarUrl: string): string {
  const rawUrl = String(calendarUrl || "").trim();

  if (!rawUrl) {
    throw new Error("Calendar URL is required.");
  }

  const cidMatch = rawUrl.match(/[?&]cid=([^&#]+)/);
  if (cidMatch && cidMatch[1]) {
    return decodeURIComponent(cidMatch[1]);
  }

  const settingsMatch = rawUrl.match(/\/settings\/calendar\/([^/?#]+)/);
  if (settingsMatch && settingsMatch[1]) {
    return decodeURIComponent(settingsMatch[1]);
  }

  throw new Error("Could not extract a calendar ID from the URL.");
}
export {
  start,
  doGet,
  // Web app handlers
  getCurrentUserEmail,
  getLoggerHealth,
  getIdentitySummary,
  getDebugSummary,
  requestCalendarForCurrentUser,
  listAllEventsForEmail,
  listAllEventsForGreyBoxId,
  listAllEventsForCalendarId,
  listAllEventsForCalendarUrl,
  listAllEventsForIdentifier,
  listAvailableGreyBoxIds,
  // Other exported service functions
  requestCalendarByGreyBoxId,
  createCalendarForRequester,
  getEventTrackingContextForCurrentUser,
  listEvents,
  listAllEventsForCurrentUser,
  listTrackedEventsForCurrentUser,
  insertEvent,
  recordTrackedTimeForCurrentUser,
  syncTrackedTimeToSheet,
  syncAllTrackedTimeToSheet,
  createCalendarForEmail,
  listAllEventsForResolvedRecord,
  extractCalendarIdFromUrl,
};
