/// <reference types="google-apps-script" />
/// <reference path="../types/logger-lib.d.ts" />

import {
  logFunctionError,
  logFunctionStart,
  logFunctionSuccess,
  loggerInfo,
} from "./logger";
import {
  getCalendarService,
  isGoogleCalendarId,
  getTrackingSpreadsheetId,
  normalizeEmail,
  normalizeGreyBoxId,
  openTrackingSpreadsheet,
} from "./records";
import type { EventTrackingContext, GreyBoxRecord } from "./types";

const DEFAULT_EVENT_LIST_LIMIT = 20;
const TRACKED_TIME_SHEET_NAME = "Tracked_Time";

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
        maxResults: Math.min(250, DEFAULT_EVENT_LIST_LIMIT),
        pageToken,
      });

      items.push(...(response.items || []));
      pageToken =
        items.length >= DEFAULT_EVENT_LIST_LIMIT
          ? undefined
          : response.nextPageToken || undefined;
    } while (pageToken);

    const limitedItems = items.slice(0, DEFAULT_EVENT_LIST_LIMIT);

    logFunctionSuccess(
      "listAllEventsForCalendar",
      { calendarId, eventCount: limitedItems.length, limit: DEFAULT_EVENT_LIST_LIMIT },
      startLog.logId
    );
    return limitedItems;
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
    spreadsheetId: getTrackingSpreadsheetId(),
  });
  const spreadsheet = openTrackingSpreadsheet();
  const existing = spreadsheet.getSheetByName(TRACKED_TIME_SHEET_NAME);

  if (existing) {
    logFunctionSuccess(
      "getOrCreateTrackedTimeSheet",
      {
        sheetName: TRACKED_TIME_SHEET_NAME,
        result: "existing",
        spreadsheetId: spreadsheet.getId(),
      },
      startLog.logId
    );
    return existing;
  }

  const created = spreadsheet.insertSheet(TRACKED_TIME_SHEET_NAME);
  logFunctionSuccess(
    "getOrCreateTrackedTimeSheet",
    {
      sheetName: TRACKED_TIME_SHEET_NAME,
      result: "created",
      spreadsheetId: spreadsheet.getId(),
    },
    startLog.logId
  );
  return created;
}

function getOrCreateSheet(sheetName: string): GoogleAppsScript.Spreadsheet.Sheet {
  const spreadsheetId = getTrackingSpreadsheetId();
  const startLog = logFunctionStart("getOrCreateSheet", { sheetName, spreadsheetId });
  const spreadsheet = openTrackingSpreadsheet();
  const existing = spreadsheet.getSheetByName(sheetName);

  if (existing) {
    logFunctionSuccess(
      "getOrCreateSheet",
      { sheetName, result: "existing", spreadsheetId: spreadsheet.getId() },
      startLog.logId
    );
    return existing;
  }

  const created = spreadsheet.insertSheet(sheetName);
  logFunctionSuccess(
    "getOrCreateSheet",
    { sheetName, result: "created", spreadsheetId: spreadsheet.getId() },
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
  return [
    "eventName (title)",
    "eventId",
    "startTime",
    "endTime",
    "duration",
    "attendeeEmail",
    "responseStatus",
    "wasPresent",
    "attendanceMarkedBy",
    "attendanceMarkedAt",
  ];
}

function buildGreyBoxAttendanceKey(
  eventId: string,
  attendeeEmail: string,
  startTime: string,
  endTime: string
): string {
  return `${eventId}::${normalizeEmail(attendeeEmail)}::${startTime}::${endTime}`;
}

function ensureGreyBoxAttendanceFormatting(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const headers = getGreyBoxSheetHeaders();
  const wasPresentIndex = headers.indexOf("wasPresent");

  if (wasPresentIndex === -1) {
    return;
  }

  const column = wasPresentIndex + 1;
  const rowCount = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, column, rowCount, 1).insertCheckboxes();
}

function clearGreyBoxRowGroups(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const lastRow = sheet.getLastRow();

  for (let rowIndex = lastRow; rowIndex >= 2; rowIndex -= 1) {
    try {
      const group = sheet.getRowGroup(rowIndex, 1);
      if (group) {
        group.remove();
      }
    } catch (_error) {
      continue;
    }
  }
}

function applyGreyBoxEventIdGroups(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  validatedRows: Array<Array<string | number | boolean>>
): void {
  const firstDataRow = 2;
  const eventIdIndex = 1;
  let runStartIndex = 0;

  while (runStartIndex < validatedRows.length) {
    const eventId = String(validatedRows[runStartIndex][eventIdIndex] || "").trim();
    let runEndIndex = runStartIndex;

    while (
      runEndIndex + 1 < validatedRows.length &&
      String(validatedRows[runEndIndex + 1][eventIdIndex] || "").trim() === eventId
    ) {
      runEndIndex += 1;
    }

    const runLength = runEndIndex - runStartIndex + 1;
    if (eventId && runLength > 1) {
      const groupStartRow = firstDataRow + runStartIndex + 1;
      const groupRowCount = runLength - 1;

      sheet.getRange(groupStartRow, 1, groupRowCount, 1).shiftRowGroupDepth(1);
      const group = sheet.getRowGroup(groupStartRow, 1);
      if (group) {
        group.collapse();
      }
    }

    runStartIndex = runEndIndex + 1;
  }
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

function getEventHostEmail(
  event: GoogleAppsScript.Calendar.Schema.Event
): string {
  const organizerEmail = normalizeEmail(String(event.organizer?.email || ""));
  const creatorEmail = normalizeEmail(String(event.creator?.email || ""));

  if (organizerEmail && !isGoogleCalendarId(organizerEmail)) {
    return organizerEmail;
  }

  if (creatorEmail && !isGoogleCalendarId(creatorEmail)) {
    return creatorEmail;
  }

  return organizerEmail || creatorEmail;
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
  const validatedRows: Array<Array<string | number | boolean>> = [];
  let totalHours = 0;
  const existingValues =
    sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues()
      : [];
  const preservedAttendanceByKey = new Map<
    string,
    { wasPresent: string | boolean; attendanceMarkedBy: string; attendanceMarkedAt: string }
  >();

  existingValues.forEach((row) => {
    const eventId = String(row[1] || "").trim();
    const startTime = String(row[2] || "").trim();
    const endTime = String(row[3] || "").trim();
    const attendeeEmail = String(row[5] || "").trim();

    if (!eventId || !startTime || !endTime) {
      return;
    }

    preservedAttendanceByKey.set(
      buildGreyBoxAttendanceKey(eventId, attendeeEmail, startTime, endTime),
      {
        wasPresent: row[7] as string | boolean,
        attendanceMarkedBy: String(row[8] || "").trim(),
        attendanceMarkedAt: String(row[9] || "").trim(),
      }
    );
  });

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
    const durationLabel = `${durationHours} hour${durationHours === 1 ? "" : "s"}`;
    const startTime = Utilities.formatDate(startAt, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    const endTime = Utilities.formatDate(endAt, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    totalHours += durationHours;
    const organizerEmail = getEventHostEmail(event);

    const attendees = (event.attendees || []).filter((attendee) => {
      const attendeeEmail = normalizeEmail(String(attendee.email || ""));
      if (!attendeeEmail) {
        return false;
      }

      if (organizerEmail && attendeeEmail === organizerEmail) {
        return false;
      }

      return true;
    });

    if (organizerEmail) {
      const preserved = preservedAttendanceByKey.get(
        buildGreyBoxAttendanceKey(event.id, organizerEmail, startTime, endTime)
      );
      validatedRows.push([
        event.summary || "Untitled Event",
        event.id,
        startTime,
        endTime,
        durationLabel,
        organizerEmail,
        "host",
        preserved ? preserved.wasPresent : "",
        preserved ? preserved.attendanceMarkedBy : "",
        preserved ? preserved.attendanceMarkedAt : "",
      ]);
    }

    if (!organizerEmail && attendees.length === 0) {
      const preserved = preservedAttendanceByKey.get(
        buildGreyBoxAttendanceKey(event.id, "", startTime, endTime)
      );
      validatedRows.push([
        event.summary || "Untitled Event",
        event.id,
        startTime,
        endTime,
        durationLabel,
        "",
        "",
        preserved ? preserved.wasPresent : "",
        preserved ? preserved.attendanceMarkedBy : "",
        preserved ? preserved.attendanceMarkedAt : "",
      ]);
      continue;
    }

    for (const attendee of attendees) {
      const attendeeEmail = String(attendee.email || "").trim();
      const preserved = preservedAttendanceByKey.get(
        buildGreyBoxAttendanceKey(event.id, attendeeEmail, startTime, endTime)
      );
      validatedRows.push([
        event.summary || "Untitled Event",
        event.id,
        startTime,
        endTime,
        durationLabel,
        attendeeEmail,
        String(attendee.responseStatus || ""),
        preserved ? preserved.wasPresent : "",
        preserved ? preserved.attendanceMarkedBy : "",
        preserved ? preserved.attendanceMarkedAt : "",
      ]);
    }
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (validatedRows.length > 0) {
    sheet
      .getRange(2, 1, validatedRows.length, headers.length)
      .setValues(validatedRows);
  }

  clearGreyBoxRowGroups(sheet);
  if (validatedRows.length > 1) {
    applyGreyBoxEventIdGroups(sheet, validatedRows);
  }

  const totalRowIndex = validatedRows.length + 3;
  sheet
    .getRange(totalRowIndex, 1, 1, headers.length)
    .setValues([[
      "TotalHours",
      "",
      "",
      "",
      Number(totalHours.toFixed(2)),
      "",
      "",
      "",
      "",
      "",
    ]]);
  ensureGreyBoxAttendanceFormatting(sheet);
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

function buildTrackedTimeRowsForRecord(
  record: GreyBoxRecord,
  calendarId: string,
  events: GoogleAppsScript.Calendar.Schema.Event[],
  syncedAt: Date
): Array<Array<string | number>> {
  const rows: Array<Array<string | number>> = [];

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

  return rows;
}

function upsertTrackedTimeRows(
  rows: Array<Array<string | number>>,
  parentLogId?: string,
  action: string = "SYNC_TRACKED_TIME"
) {
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

  loggerInfo({
    message: "Tracked time synced to sheet",
    parentLogId,
    meta: {
      action,
      rowCount: rows.length,
      insertedCount,
      updatedCount,
      sheetName: TRACKED_TIME_SHEET_NAME,
      spreadsheetId: getTrackingSpreadsheetId(),
    },
  });

  return {
    ok: true,
    rowCount: rows.length,
    insertedCount,
    updatedCount,
    sheetName: TRACKED_TIME_SHEET_NAME,
    spreadsheetId: getTrackingSpreadsheetId(),
  };
}

export {
  parseRequiredDate,
  listAllEventsForCalendar,
  recordGreyBoxSheetRequest,
  writeValidatedGreyBoxSheet,
  buildTrackedTimeRowsForRecord,
  upsertTrackedTimeRows,
};
