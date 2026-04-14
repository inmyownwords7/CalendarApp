/// <reference types="google-apps-script" />
/// <reference path="./types/types.d.ts" />

export type GreyBoxRecord = {
  greyBoxId: string;
  email: string;
  calendarTitle: string;
};

type MandateMatch = {
  record: GreyBoxRecord;
};

export type CalendarRequestResponse = {
  ok: true;
  greyBoxId: string;
  email: string;
  calendarId: string;
  calendarUrl: string;
  title: string;
};

export type EventTrackingContext = {
  greyBoxId: string;
  requesterEmail: string;
  calendarId: string;
  calendarTitle: string;
  calendarUrl: string;
};

export type TimeEntryInput = {
  summary: string;
  startIso: string;
  endIso: string;
  description?: string;
};

export type IdentitySummary = {
  activeEmail: string;
  effectiveEmail: string;
  isRequesterDeployerSame: boolean;
};

const LOG_SPREADSHEET_ID = "1RFcvNfu07jUPpGQan59UcK6NKUqqiTW4mO1Gmtkvdas";
const MANDATE_SPREADSHEET_ID = "1pfvdPNPiJx4XnTUAVzstqENOTYl2GgLTYYea5KMWcpg";
const MANDATE_SHEET_NAME = "Mandate";
const CALENDAR_PROPERTY_PREFIX = "trackedCalendar:";
const TRACKED_TIME_SHEET_NAME = "Tracked_Time";
let isLoggerInitialized = false;

const REQUESTER_DIRECTORY: Record<string, GreyBoxRecord> = {
  "alex.manager@example.com": {
    greyBoxId: "GB1001",
    email: "alex.manager@example.com",
    calendarTitle: "Grey Box GB1001 Calendar",
  },
  "jamie.requester@example.com": {
    greyBoxId: "GB1002",
    email: "jamie.requester@example.com",
    calendarTitle: "Grey Box GB1002 Calendar",
  },
  "taylor.coordinator@example.com": {
    greyBoxId: "GB1003",
    email: "taylor.coordinator@example.com",
    calendarTitle: "Grey Box GB1003 Calendar",
  },
};

const GREY_BOX_DIRECTORY: Record<string, GreyBoxRecord> = Object.values(
  REQUESTER_DIRECTORY
).reduce<Record<string, GreyBoxRecord>>((acc, record) => {
  acc[record.greyBoxId] = record;
  return acc;
}, {});

function logInit(): void {
  if (isLoggerInitialized) {
    return;
  }

  LoggerLib.init({
    spreadsheetId: LOG_SPREADSHEET_ID,
    operationalSheet: "Operational_Log",
    networkSheet: "Network_Log",
    level: "DEBUG",
  });

  isLoggerInitialized = true;
}

function toErrorMeta(error: unknown): Record<string, unknown> {
  if (error instanceof Error) {
    return {
      name: error.name,
      message: error.message,
      stack: error.stack || "",
    };
  }

  return {
    message: String(error),
  };
}

function logFunctionStart(
  functionName: string,
  meta: Record<string, unknown> = {}
): LoggerLibOperationalLogRecord {
  logInit();
  return LoggerLib.info({
    message: `${functionName} started`,
    meta: {
      functionName,
      stage: "start",
      ...meta,
    },
  });
}

function logFunctionSuccess(
  functionName: string,
  meta: Record<string, unknown> = {},
  parentLogId?: string
): LoggerLibOperationalLogRecord {
  logInit();
  return LoggerLib.info({
    message: `${functionName} completed`,
    parentLogId,
    meta: {
      functionName,
      stage: "success",
      ...meta,
    },
  });
}

function logFunctionError(
  functionName: string,
  error: unknown,
  meta: Record<string, unknown> = {},
  parentLogId?: string
): LoggerLibOperationalLogRecord {
  logInit();
  return LoggerLib.error({
    message: `${functionName} failed`,
    error,
    parentLogId,
    meta: {
      functionName,
      stage: "error",
      ...meta,
      errorDetails: toErrorMeta(error),
    },
  });
}

function getCalendarService(): GoogleAppsScript.Calendar {
  if (!Calendar) {
    throw new Error("Advanced Calendar service is not enabled.");
  }
  return Calendar;
}

function normalizeGreyBoxId(greyBoxId: string): string {
  return (greyBoxId || "").trim().toUpperCase();
}

function normalizeEmail(email: string): string {
  return (email || "").trim().toLowerCase();
}

function buildTrackingCalendarTitle(greyBoxId: string): string {
  return `${normalizeGreyBoxId(greyBoxId)} - Tracking Calendar`;
}

function getMandateSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const startLog = logFunctionStart("getMandateSheet", {
    spreadsheetId: MANDATE_SPREADSHEET_ID,
    sheetName: MANDATE_SHEET_NAME,
  });

  try {
    const spreadsheet = SpreadsheetApp.openById(MANDATE_SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(MANDATE_SHEET_NAME);

    if (!sheet) {
      throw new Error(`Missing required sheet: ${MANDATE_SHEET_NAME}`);
    }

    logFunctionSuccess(
      "getMandateSheet",
      {
        spreadsheetId: MANDATE_SPREADSHEET_ID,
        sheetName: MANDATE_SHEET_NAME,
      },
      startLog.logId
    );
    return sheet;
  } catch (error) {
    logFunctionError(
      "getMandateSheet",
      error,
      {
        spreadsheetId: MANDATE_SPREADSHEET_ID,
        sheetName: MANDATE_SHEET_NAME,
      },
      startLog.logId
    );
    throw error;
  }
}

function getMandateSheetRecords(): GreyBoxRecord[] {
  const startLog = logFunctionStart("getMandateSheetRecords");

  try {
    const sheet = getMandateSheet();
    const values = sheet.getDataRange().getValues();

    if (values.length < 2) {
      logFunctionSuccess(
        "getMandateSheetRecords",
        { recordCount: 0, source: "sheet_empty" },
        startLog.logId
      );
      return [];
    }

    const headers = values[0].map((value) => String(value).trim());
    const greyBoxIdIndex = headers.indexOf("greyBoxId");
    const emailIndex = headers.indexOf("email");
    const calendarTitleIndex = headers.indexOf("calendarTitle");

    if (greyBoxIdIndex === -1 || emailIndex === -1) {
      throw new Error(
        `Mandate sheet must include 'greyBoxId' and 'email' columns in ${MANDATE_SHEET_NAME}.`
      );
    }

    const records = values
      .slice(1)
      .map((row) => {
        const greyBoxId = normalizeGreyBoxId(String(row[greyBoxIdIndex] || ""));
        const email = normalizeEmail(String(row[emailIndex] || ""));
        const calendarTitle =
          calendarTitleIndex === -1
            ? ""
            : String(row[calendarTitleIndex] || "").trim();

        if (!greyBoxId || !email) {
          return null;
        }

        return {
          greyBoxId,
          email,
          calendarTitle: calendarTitle || buildTrackingCalendarTitle(greyBoxId),
        };
      })
      .filter((record): record is GreyBoxRecord => record !== null);

    logFunctionSuccess(
      "getMandateSheetRecords",
      { recordCount: records.length },
      startLog.logId
    );
    return records;
  } catch (error) {
    logFunctionError("getMandateSheetRecords", error, {}, startLog.logId);
    throw error;
  }
}

function getGreyBoxDirectory(): Record<string, GreyBoxRecord> {
  const startLog = logFunctionStart("getGreyBoxDirectory");

  try {
    const mandateRecords = getMandateSheetRecords();

    if (mandateRecords.length > 0) {
      const directory = mandateRecords.reduce<Record<string, GreyBoxRecord>>(
        (acc, record) => {
          acc[record.greyBoxId] = record;
          return acc;
        },
        {}
      );

      logFunctionSuccess(
        "getGreyBoxDirectory",
        { source: "mandate_sheet", recordCount: mandateRecords.length },
        startLog.logId
      );
      return directory;
    }

    logFunctionSuccess(
      "getGreyBoxDirectory",
      { source: "fallback_directory", recordCount: Object.keys(GREY_BOX_DIRECTORY).length },
      startLog.logId
    );
    return GREY_BOX_DIRECTORY;
  } catch (error) {
    logFunctionError("getGreyBoxDirectory", error, {}, startLog.logId);
    throw error;
  }
}

function getGreyBoxRecord(greyBoxId: string): GreyBoxRecord {
  const startLog = logFunctionStart("getGreyBoxRecord", { greyBoxId });
  const normalized = normalizeGreyBoxId(greyBoxId);

  try {
    const record = getGreyBoxDirectory()[normalized];

    if (!record) {
      throw new Error(`Unknown grey-box ID: ${normalized || "(blank)"}`);
    }

    logFunctionSuccess(
      "getGreyBoxRecord",
      { greyBoxId: normalized, email: record.email },
      startLog.logId
    );
    return record;
  } catch (error) {
    logFunctionError(
      "getGreyBoxRecord",
      error,
      { greyBoxId: normalized || "(blank)" },
      startLog.logId
    );
    throw error;
  }
}

function getRequesterDirectory(): Record<string, GreyBoxRecord> {
  const startLog = logFunctionStart("getRequesterDirectory");

  try {
    const mandateRecords = getMandateSheetRecords();

    if (mandateRecords.length > 0) {
      const directory = mandateRecords.reduce<Record<string, GreyBoxRecord>>(
        (acc, record) => {
          acc[normalizeEmail(record.email)] = record;
          return acc;
        },
        {}
      );

      logFunctionSuccess(
        "getRequesterDirectory",
        { source: "mandate_sheet", recordCount: mandateRecords.length },
        startLog.logId
      );
      return directory;
    }

    logFunctionSuccess(
      "getRequesterDirectory",
      { source: "fallback_directory", recordCount: Object.keys(REQUESTER_DIRECTORY).length },
      startLog.logId
    );
    return REQUESTER_DIRECTORY;
  } catch (error) {
    logFunctionError("getRequesterDirectory", error, {}, startLog.logId);
    throw error;
  }
}

function getMandateMatchByEmail(email: string): MandateMatch {
  const startLog = logFunctionStart("getMandateMatchByEmail", { email });
  const normalized = normalizeEmail(email);
  try {
    const sheet = getMandateSheet();
    const values = sheet.getDataRange().getValues();

    if (values.length < 2) {
      throw new Error(`No mandate rows found in ${MANDATE_SHEET_NAME}.`);
    }

    const headers = values[0].map((value) => String(value).trim());
    const greyBoxIdIndex = headers.indexOf("greyBoxId");
    const emailIndex = headers.indexOf("email");
    const calendarTitleIndex = headers.indexOf("calendarTitle");

    if (greyBoxIdIndex === -1 || emailIndex === -1) {
      throw new Error(
        `Mandate sheet must include 'greyBoxId' and 'email' columns in ${MANDATE_SHEET_NAME}.`
      );
    }

    for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
      const row = values[rowIndex];
      const rowEmail = normalizeEmail(String(row[emailIndex] || ""));

      if (rowEmail !== normalized) {
        continue;
      }

      const greyBoxId = normalizeGreyBoxId(String(row[greyBoxIdIndex] || ""));
      if (!greyBoxId) {
        throw new Error(`Mandate row for ${normalized} is missing greyBoxId.`);
      }

      const calendarTitle =
        calendarTitleIndex === -1
          ? ""
          : String(row[calendarTitleIndex] || "").trim();

      logFunctionSuccess(
        "getMandateMatchByEmail",
        { email: normalized, source: "mandate_sheet", greyBoxId },
        startLog.logId
      );
      return {
        record: {
          greyBoxId,
          email: rowEmail,
          calendarTitle: calendarTitle || buildTrackingCalendarTitle(greyBoxId),
        },
      };
    }

    const fallbackRecord = getRequesterDirectory()[normalized];
    if (!fallbackRecord) {
      throw new Error(`No mandate found for signed-in user: ${normalized || "(blank)"}`);
    }

    logFunctionSuccess(
      "getMandateMatchByEmail",
      {
        email: normalized,
        source: "fallback_directory",
        greyBoxId: fallbackRecord.greyBoxId,
      },
      startLog.logId
    );
    return {
      record: {
        ...fallbackRecord,
        calendarTitle: buildTrackingCalendarTitle(fallbackRecord.greyBoxId),
      },
    };
  } catch (error) {
    logFunctionError(
      "getMandateMatchByEmail",
      error,
      { email: normalized || "(blank)" },
      startLog.logId
    );
    throw error;
  }
}

function getCalendarPropertyKey(greyBoxId: string): string {
  return `${CALENDAR_PROPERTY_PREFIX}${normalizeGreyBoxId(greyBoxId)}`;
}

function getStoredCalendarId(greyBoxId: string): string {
  const calendarId = (
    PropertiesService.getScriptProperties().getProperty(
      getCalendarPropertyKey(greyBoxId)
    ) || ""
  ).trim();

  logFunctionSuccess("getStoredCalendarId", {
    greyBoxId: normalizeGreyBoxId(greyBoxId),
    hasCalendarId: Boolean(calendarId),
    calendarId,
  });

  return calendarId;
}

function storeCalendarId(greyBoxId: string, calendarId: string): void {
  const startLog = logFunctionStart("storeCalendarId", {
    greyBoxId,
    calendarId,
  });
  PropertiesService.getScriptProperties().setProperty(
    getCalendarPropertyKey(greyBoxId),
    calendarId
  );
  logFunctionSuccess(
    "storeCalendarId",
    {
      greyBoxId: normalizeGreyBoxId(greyBoxId),
      calendarId,
    },
    startLog.logId
  );
}

function buildCalendarUrl(calendarId: string): string {
  return `https://calendar.google.com/calendar/u/0/r/settings/calendar/${encodeURIComponent(calendarId)}`;
}

function shareCalendarWithUser(calendarId: string, email: string): void {
  const startLog = logFunctionStart("shareCalendarWithUser", { calendarId, email });
  const calendar = getCalendarService();

  try {
    calendar.Acl.insert(
      { role: "writer", scope: { type: "user", value: email } },
      calendarId
    );
  } catch (error) {
    const message = String(error);

    if (
      message.includes("already exists") ||
      message.includes("duplicate") ||
      message.includes("ACL")
    ) {
      LoggerLib.info({
        message: "Calendar was already shared to user",
        meta: {
          action: "SHARE_CALENDAR",
          calendarId,
          email,
        },
      });
      logFunctionSuccess(
        "shareCalendarWithUser",
        { calendarId, email, result: "already_shared" },
        startLog.logId
      );
      return;
    }

    logFunctionError(
      "shareCalendarWithUser",
      error,
      { calendarId, email },
      startLog.logId
    );
    throw error;
  }

  logFunctionSuccess(
    "shareCalendarWithUser",
    { calendarId, email, result: "shared" },
    startLog.logId
  );
}

function ensureCalendarForRecord(
  record: GreyBoxRecord,
  requesterEmail: string
): EventTrackingContext {
  const startLog = logFunctionStart("ensureCalendarForRecord", {
    greyBoxId: record.greyBoxId,
    requesterEmail,
  });
  const calendar = getCalendarService();
  const storedCalendarId = getStoredCalendarId(record.greyBoxId);

  try {
    if (storedCalendarId) {
      shareCalendarWithUser(storedCalendarId, requesterEmail);

      const context = {
        greyBoxId: record.greyBoxId,
        requesterEmail,
        calendarId: storedCalendarId,
        calendarTitle: record.calendarTitle,
        calendarUrl: buildCalendarUrl(storedCalendarId),
      };

      logFunctionSuccess(
        "ensureCalendarForRecord",
        {
          greyBoxId: record.greyBoxId,
          requesterEmail,
          calendarId: storedCalendarId,
          source: "script_properties",
        },
        startLog.logId
      );
      return context;
    }

    const cal = calendar.Calendars.insert({
      summary: record.calendarTitle,
      timeZone: Session.getScriptTimeZone(),
    });

    if (!cal.id) {
      throw new Error("Calendar creation failed: No calendar ID returned.");
    }

    storeCalendarId(record.greyBoxId, cal.id);
    shareCalendarWithUser(cal.id, requesterEmail);

    const context = {
      greyBoxId: record.greyBoxId,
      requesterEmail,
      calendarId: cal.id,
      calendarTitle: record.calendarTitle,
      calendarUrl: buildCalendarUrl(cal.id),
    };

    logFunctionSuccess(
      "ensureCalendarForRecord",
      {
        greyBoxId: record.greyBoxId,
        requesterEmail,
        calendarId: cal.id,
        source: "calendar_created",
      },
      startLog.logId
    );
    return context;
  } catch (error) {
    logFunctionError(
      "ensureCalendarForRecord",
      error,
      {
        greyBoxId: record.greyBoxId,
        requesterEmail,
      },
      startLog.logId
    );
    throw error;
  }
}

function getEventTrackingContextForCurrentUserInternal(): EventTrackingContext {
  const startLog = logFunctionStart("getEventTrackingContextForCurrentUserInternal");
  const requesterEmail = normalizeEmail(Session.getActiveUser().getEmail());

  try {
    if (!requesterEmail) {
      throw new Error("Could not determine the signed-in user email.");
    }

    const mandateMatch = getMandateMatchByEmail(requesterEmail);
    const context = ensureCalendarForRecord(mandateMatch.record, requesterEmail);
    logFunctionSuccess(
      "getEventTrackingContextForCurrentUserInternal",
      {
        requesterEmail,
        greyBoxId: context.greyBoxId,
        calendarId: context.calendarId,
      },
      startLog.logId
    );
    return context;
  } catch (error) {
    logFunctionError(
      "getEventTrackingContextForCurrentUserInternal",
      error,
      { requesterEmail: requesterEmail || "(blank)" },
      startLog.logId
    );
    throw error;
  }
}

function parseRequiredDate(dateIso: string, fieldName: string): Date {
  const startLog = logFunctionStart("parseRequiredDate", { fieldName, dateIso });
  const value = new Date(dateIso);

  if (Number.isNaN(value.getTime())) {
    const error = new Error(`Invalid ${fieldName}. Expected ISO datetime.`);
    logFunctionError(
      "parseRequiredDate",
      error,
      { fieldName, dateIso },
      startLog.logId
    );
    throw error;
  }

  logFunctionSuccess(
    "parseRequiredDate",
    { fieldName, dateIso, parsedIso: value.toISOString() },
    startLog.logId
  );
  return value;
}

function listAllEventsForCalendar(
  calendarId: string
): GoogleAppsScript.Calendar.Schema.Event[] {
  const startLog = logFunctionStart("listAllEventsForCalendar", { calendarId });
  const calendar = getCalendarService();
  const items: GoogleAppsScript.Calendar.Schema.Event[] = [];
  let pageToken: string | undefined;

  try {
    do {
      const response = calendar.Events.list(calendarId, {
        singleEvents: true,
        orderBy: "startTime",
        maxResults: 250,
        pageToken,
      });

      items.push(...(response.items || []));
      pageToken = response.nextPageToken || undefined;
    } while (pageToken);

    logFunctionSuccess(
      "listAllEventsForCalendar",
      { calendarId, eventCount: items.length },
      startLog.logId
    );
    return items;
  } catch (error) {
    logFunctionError(
      "listAllEventsForCalendar",
      error,
      { calendarId },
      startLog.logId
    );
    throw error;
  }
}

function getEventBoundary(eventDate: {
  date?: string | null;
  dateTime?: string | null;
} | null | undefined): Date | null {
  if (!eventDate) {
    return null;
  }

  const rawValue = eventDate.dateTime || eventDate.date;
  if (!rawValue) {
    return null;
  }

  const parsed = new Date(rawValue);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function getOrCreateTrackedTimeSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const startLog = logFunctionStart("getOrCreateTrackedTimeSheet", {
    sheetName: TRACKED_TIME_SHEET_NAME,
  });
  const spreadsheet = SpreadsheetApp.openById(MANDATE_SPREADSHEET_ID);
  const existing = spreadsheet.getSheetByName(TRACKED_TIME_SHEET_NAME);

  if (existing) {
    logFunctionSuccess(
      "getOrCreateTrackedTimeSheet",
      { sheetName: TRACKED_TIME_SHEET_NAME, result: "existing" },
      startLog.logId
    );
    return existing;
  }

  const created = spreadsheet.insertSheet(TRACKED_TIME_SHEET_NAME);
  logFunctionSuccess(
    "getOrCreateTrackedTimeSheet",
    { sheetName: TRACKED_TIME_SHEET_NAME, result: "created" },
    startLog.logId
  );
  return created;
}

function getOrCreateSheet(sheetName: string): GoogleAppsScript.Spreadsheet.Sheet {
  const startLog = logFunctionStart("getOrCreateSheet", { sheetName });
  const spreadsheet = SpreadsheetApp.openById(MANDATE_SPREADSHEET_ID);
  const existing = spreadsheet.getSheetByName(sheetName);

  if (existing) {
    logFunctionSuccess(
      "getOrCreateSheet",
      { sheetName, result: "existing" },
      startLog.logId
    );
    return existing;
  }

  const created = spreadsheet.insertSheet(sheetName);
  logFunctionSuccess(
    "getOrCreateSheet",
    { sheetName, result: "created" },
    startLog.logId
  );
  return created;
}

function ensureHeaderRow(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  headers: string[]
): void {
  const width = headers.length;
  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, width).setValues([headers]);
    return;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, width).getValues()[0];
  const headersMatch = headers.every((header, index) => currentHeaders[index] === header);

  if (!headersMatch) {
    sheet.getRange(1, 1, 1, width).setValues([headers]);
  }
}

function getTrackedTimeHeaders(): string[] {
  return [
    "greyBoxId",
    "email",
    "calendarId",
    "eventId",
    "summary",
    "startIso",
    "endIso",
    "durationMinutes",
    "durationHours",
    "syncedAt",
  ];
}

function getGreyBoxSheetHeaders(): string[] {
  return ["EventTitle", "DurationHours"];
}

function getOrCreateGreyBoxSheet(greyBoxId: string): GoogleAppsScript.Spreadsheet.Sheet {
  const startLog = logFunctionStart("getOrCreateGreyBoxSheet", { greyBoxId });
  const sheet = getOrCreateSheet(normalizeGreyBoxId(greyBoxId));
  ensureHeaderRow(sheet, getGreyBoxSheetHeaders());
  logFunctionSuccess(
    "getOrCreateGreyBoxSheet",
    { greyBoxId: normalizeGreyBoxId(greyBoxId), sheetName: sheet.getName() },
    startLog.logId
  );
  return sheet;
}

function recordGreyBoxSheetRequest(context: EventTrackingContext): void {
  const startLog = logFunctionStart("recordGreyBoxSheetRequest", {
    greyBoxId: context.greyBoxId,
    requesterEmail: context.requesterEmail,
  });
  getOrCreateGreyBoxSheet(context.greyBoxId);
  logFunctionSuccess(
    "recordGreyBoxSheetRequest",
    {
      greyBoxId: context.greyBoxId,
      requesterEmail: context.requesterEmail,
    },
    startLog.logId
  );
}

function writeValidatedGreyBoxSheet(
  greyBoxId: string,
  events: GoogleAppsScript.Calendar.Schema.Event[]
): void {
  const startLog = logFunctionStart("writeValidatedGreyBoxSheet", {
    greyBoxId,
    inputEventCount: events.length,
  });
  const sheet = getOrCreateGreyBoxSheet(greyBoxId);
  const headers = getGreyBoxSheetHeaders();
  const validatedRows: Array<Array<string | number>> = [];
  let totalHours = 0;

  for (const event of events) {
    if (!event.id || event.status === "cancelled") {
      continue;
    }

    const startAt = getEventBoundary(event.start);
    const endAt = getEventBoundary(event.end);

    if (!startAt || !endAt) {
      continue;
    }

    const durationHours = Number(
      Math.max(0, (endAt.getTime() - startAt.getTime()) / 3600000).toFixed(2)
    );

    validatedRows.push([event.summary || "Untitled Event", durationHours]);
    totalHours += durationHours;
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (validatedRows.length > 0) {
    sheet
      .getRange(2, 1, validatedRows.length, headers.length)
      .setValues(validatedRows);
  }

  const totalRowIndex = validatedRows.length + 3;
  sheet
    .getRange(totalRowIndex, 1, 1, headers.length)
    .setValues([["TotalHours", Number(totalHours.toFixed(2))]]);
  logFunctionSuccess(
    "writeValidatedGreyBoxSheet",
    {
      greyBoxId: normalizeGreyBoxId(greyBoxId),
      validatedRowCount: validatedRows.length,
      totalHours: Number(totalHours.toFixed(2)),
    },
    startLog.logId
  );
}

export function listAvailableGreyBoxIds(): string[] {
  return Object.keys(getGreyBoxDirectory()).sort();
}

export function listEvents(calendarId: string) {
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

export function insertEvent(
  calendarId: string,
  event: GoogleAppsScript.Calendar.Schema.Event
) {
  logInit();
  const calendar = getCalendarService();

  try {
    const ev = calendar.Events.insert(event, calendarId);
    LoggerLib.info({
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
    LoggerLib.error({
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

export function start(): void {
  logInit();

  LoggerLib.info({
    message: "Identity sync started",
    meta: { source: "identityBuild" },
  });

  LoggerLib.http({
    system: "SLACK",
    method: "GET",
    url: { raw: "https://slack.com/api/users.list" },
    status: 200,
    durationMs: 180,
  });
}

export function getCurrentUserEmail(): string {
  logInit();

  const email = Session.getActiveUser().getEmail() || "";

  const rec = LoggerLib.info({
    message: email
      ? "Fetched current user email"
      : "Current user email unavailable",
    meta: { email },
  });

  Logger.log("logId=" + rec.logId);
  return email;
}

export function getIdentitySummary(): IdentitySummary {
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

export function doGet(): GoogleAppsScript.HTML.HtmlOutput {
  return HtmlService.createHtmlOutputFromFile("index").setTitle(
    "Calendar Request Manager"
  );
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

    const mainLog = LoggerLib.info({
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

    LoggerLib.http({
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

export function createCalendarForRequester(
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

export function requestCalendarByGreyBoxId(
  greyBoxId: string
): CalendarRequestResponse {
  logInit();
  const startLog = logFunctionStart("requestCalendarByGreyBoxId", { greyBoxId });

  try {
    const record = getGreyBoxRecord(greyBoxId);

    LoggerLib.info({
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

export function requestCalendarForCurrentUser(): CalendarRequestResponse {
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

    LoggerLib.info({
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

export function getEventTrackingContextForCurrentUser(): EventTrackingContext {
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

export function listTrackedEventsForCurrentUser() {
  logInit();
  const startLog = logFunctionStart("listTrackedEventsForCurrentUser");
  const context = getEventTrackingContextForCurrentUserInternal();
  const calendar = getCalendarService();

  LoggerLib.info({
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

export function recordTrackedTimeForCurrentUser(input: TimeEntryInput) {
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

  LoggerLib.info({
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

export function syncTrackedTimeToSheet() {
  logInit();
  const startLog = logFunctionStart("syncTrackedTimeToSheet");

  const records = Object.values(getGreyBoxDirectory());
  const rows: Array<Array<string | number>> = [];
  const syncedAt = new Date();

  for (const record of records) {
    const calendarId = getStoredCalendarId(record.greyBoxId);

    if (!calendarId) {
      continue;
    }

    try {
      const events = listAllEventsForCalendar(calendarId);
      writeValidatedGreyBoxSheet(record.greyBoxId, events);

      for (const event of events) {
        if (!event.id || event.status === "cancelled") {
          continue;
        }

        const startAt = getEventBoundary(event.start);
        const endAt = getEventBoundary(event.end);

        if (!startAt || !endAt) {
          continue;
        }

        const durationMinutes = Math.max(
          0,
          Math.round((endAt.getTime() - startAt.getTime()) / 60000)
        );

        rows.push([
          record.greyBoxId,
          record.email,
          calendarId,
          event.id,
          event.summary || "",
          startAt.toISOString(),
          endAt.toISOString(),
          durationMinutes,
          Number((durationMinutes / 60).toFixed(2)),
          syncedAt.toISOString(),
        ]);
      }
    } catch (error) {
      LoggerLib.error({
        message: "Failed to sync tracked calendar events",
        error,
        meta: {
          action: "SYNC_TRACKED_TIME",
          greyBoxId: record.greyBoxId,
          email: record.email,
          calendarId,
        },
      });
    }
  }

  const sheet = getOrCreateTrackedTimeSheet();
  const headers = getTrackedTimeHeaders();
  const width = headers.length;
  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, width).setValues([headers]);
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, width).getValues()[0];
    const headersMatch = headers.every((header, index) => currentHeaders[index] === header);

    if (!headersMatch) {
      sheet.getRange(1, 1, 1, width).setValues([headers]);
    }
  }

  const existingLastRow = sheet.getLastRow();
  const existingValues =
    existingLastRow > 1
      ? sheet.getRange(2, 1, existingLastRow - 1, width).getValues()
      : [];

  const rowIndexByEventKey = new Map<string, number>();

  existingValues.forEach((row, index) => {
    const calendarId = String(row[2] || "").trim();
    const eventId = String(row[3] || "").trim();

    if (!calendarId || !eventId) {
      return;
    }

    rowIndexByEventKey.set(`${calendarId}:${eventId}`, index + 2);
  });

  const rowsToAppend: Array<Array<string | number>> = [];
  let insertedCount = 0;
  let updatedCount = 0;

  for (const row of rows) {
    const calendarId = String(row[2] || "").trim();
    const eventId = String(row[3] || "").trim();
    const key = `${calendarId}:${eventId}`;
    const targetRow = rowIndexByEventKey.get(key);

    if (!targetRow) {
      rowsToAppend.push(row);
      insertedCount += 1;
      continue;
    }

    const currentRow = sheet.getRange(targetRow, 1, 1, width).getValues()[0];
    const hasChanged = row.some((value, index) => currentRow[index] !== value);

    if (!hasChanged) {
      continue;
    }

    sheet.getRange(targetRow, 1, 1, width).setValues([row]);
    updatedCount += 1;
  }

  if (rowsToAppend.length > 0) {
    const appendStartRow = sheet.getLastRow() + 1;
    sheet
      .getRange(appendStartRow, 1, rowsToAppend.length, width)
      .setValues(rowsToAppend);
  }

  LoggerLib.info({
    message: "Tracked time synced to sheet",
    parentLogId: startLog.logId,
    meta: {
      action: "SYNC_TRACKED_TIME",
      rowCount: rows.length,
      insertedCount,
      updatedCount,
      sheetName: TRACKED_TIME_SHEET_NAME,
      spreadsheetId: MANDATE_SPREADSHEET_ID,
    },
  });

  return {
    ok: true,
    rowCount: rows.length,
    insertedCount,
    updatedCount,
    sheetName: TRACKED_TIME_SHEET_NAME,
    spreadsheetId: MANDATE_SPREADSHEET_ID,
    syncedAt: syncedAt.toISOString(),
  };
}
