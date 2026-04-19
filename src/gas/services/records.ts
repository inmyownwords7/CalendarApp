/// <reference types="google-apps-script" />
/// <reference path="../types/logger-lib.d.ts" />

import {
  logFunctionError,
  logFunctionStart,
  logFunctionSuccess,
  loggerInfo,
} from "./logger";
import type { EventTrackingContext, GreyBoxRecord } from "./types";

type MandateMatch = {
  record: GreyBoxRecord;
};

const MANDATE_SPREADSHEET_ID = "1pfvdPNPiJx4XnTUAVzstqENOTYl2GgLTYYea5KMWcpg";
const MANDATE_SHEET_CANDIDATES = ["Mandates", "Mandate"];
const CALENDAR_PROPERTY_PREFIX = "trackedCalendar:";
const CALENDAR_LOOKUP_PROPERTY_PREFIX = "trackedCalendarLookup:";
const TRACKING_SPREADSHEET_PROPERTY_KEY = "TRACKING_SPREADSHEET_ID";
const LOG_SPREADSHEET_ID = "1RFcvNfu07jUPpGQan59UcK6NKUqqiTW4mO1Gmtkvdas";

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

function isGoogleCalendarId(identifier: string): boolean {
  const normalized = String(identifier || "").trim().toLowerCase();
  return normalized.endsWith("@group.calendar.google.com");
}

function findHeaderIndex(headers: string[], candidates: string[]): number {
  for (const candidate of candidates) {
    const index = headers.indexOf(candidate);
    if (index !== -1) {
      return index;
    }
  }

  return -1;
}

function getTrackingSpreadsheetId(): string {
  const configuredId = (
    PropertiesService.getScriptProperties().getProperty(
      TRACKING_SPREADSHEET_PROPERTY_KEY
    ) || ""
  ).trim();

  return configuredId || LOG_SPREADSHEET_ID;
}

function openTrackingSpreadsheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
  return SpreadsheetApp.openById(getTrackingSpreadsheetId());
}

function buildTrackingCalendarTitle(greyBoxId: string): string {
  return `${normalizeGreyBoxId(greyBoxId)} - Tracking Calendar`;
}

function getMandateSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const startLog = logFunctionStart("getMandateSheet", {
    spreadsheetId: MANDATE_SPREADSHEET_ID,
    sheetNames: MANDATE_SHEET_CANDIDATES,
  });

  try {
    const spreadsheet = SpreadsheetApp.openById(MANDATE_SPREADSHEET_ID);
    const sheet = MANDATE_SHEET_CANDIDATES
      .map((name) => spreadsheet.getSheetByName(name))
      .find(
        (candidate): candidate is GoogleAppsScript.Spreadsheet.Sheet =>
          candidate !== null
      );

    if (!sheet) {
      throw new Error(
        `Missing required sheet. Tried: ${MANDATE_SHEET_CANDIDATES.join(", ")}`
      );
    }

    logFunctionSuccess(
      "getMandateSheet",
      {
        spreadsheetId: MANDATE_SPREADSHEET_ID,
        sheetName: sheet.getName(),
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
        sheetNames: MANDATE_SHEET_CANDIDATES,
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
    const greyBoxIdIndex = findHeaderIndex(headers, ["greyBoxId", "Greybox ID"]);
    const emailIndex = findHeaderIndex(headers, ["email", "Email (Org)"]);
    const calendarTitleIndex = findHeaderIndex(headers, [
      "calendarTitle",
      "Calendar Title",
      "Title",
    ]);

    if (greyBoxIdIndex === -1 || emailIndex === -1) {
      throw new Error(
        "Mandate sheet must include Greybox ID and Email (Org) columns."
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

function getUniqueGreyBoxIdPrefixMatch(identifier: string): string {
  const normalizedPrefix = normalizeGreyBoxId(identifier);

  if (!normalizedPrefix) {
    return "";
  }

  const matches = Object.keys(getGreyBoxDirectory()).filter((greyBoxId) =>
    greyBoxId.startsWith(normalizedPrefix)
  );

  return matches.length === 1 ? matches[0] : "";
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
      throw new Error("No mandate rows found in the mandate sheet.");
    }

    const headers = values[0].map((value) => String(value).trim());
    const greyBoxIdIndex = findHeaderIndex(headers, ["greyBoxId", "Greybox ID"]);
    const emailIndex = findHeaderIndex(headers, ["email", "Email (Org)"]);
    const calendarTitleIndex = findHeaderIndex(headers, [
      "calendarTitle",
      "Calendar Title",
      "Title",
    ]);

    if (greyBoxIdIndex === -1 || emailIndex === -1) {
      throw new Error(
        "Mandate sheet must include Greybox ID and Email (Org) columns."
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
      throw new Error(`No mandate found for email: ${normalized || "(blank)"}`);
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

function getMandateMatchByIdentifier(identifier: string): MandateMatch {
  const normalized = (identifier || "").trim();

  if (!normalized) {
    throw new Error("Identifier is required.");
  }

  if (isGoogleCalendarId(normalized)) {
    const matchedGreyBoxId = getGreyBoxIdForStoredCalendarId(normalized);
    if (matchedGreyBoxId) {
      const matchedRecord = getGreyBoxDirectory()[matchedGreyBoxId];
      if (matchedRecord) {
        return { record: matchedRecord };
      }
    }
  }

  if (normalized.includes("@")) {
    return getMandateMatchByEmail(normalized);
  }

  const normalizedGreyBoxId = normalizeGreyBoxId(normalized);
  const greyBoxRecord = getGreyBoxDirectory()[normalizedGreyBoxId];
  if (greyBoxRecord) {
    return { record: greyBoxRecord };
  }

  const prefixMatchedGreyBoxId = getUniqueGreyBoxIdPrefixMatch(normalized);
  if (prefixMatchedGreyBoxId) {
    return { record: getGreyBoxDirectory()[prefixMatchedGreyBoxId] };
  }

  const matchedGreyBoxId = getGreyBoxIdForStoredCalendarId(normalized);
  if (matchedGreyBoxId) {
    const matchedRecord = getGreyBoxDirectory()[matchedGreyBoxId];
    if (matchedRecord) {
      return { record: matchedRecord };
    }
  }

  const records = Object.values(getGreyBoxDirectory());
  for (const record of records) {
    const storedCalendarId = getStoredCalendarId(record.greyBoxId);
    if (storedCalendarId && storedCalendarId === normalized) {
      storeCalendarId(record.greyBoxId, storedCalendarId);
      return { record };
    }
  }

  throw new Error(`No mandate match found for identifier: ${normalized}`);
}

function getCalendarPropertyKey(greyBoxId: string): string {
  return `${CALENDAR_PROPERTY_PREFIX}${normalizeGreyBoxId(greyBoxId)}`;
}

function getCalendarLookupPropertyKey(calendarId: string): string {
  return `${CALENDAR_LOOKUP_PROPERTY_PREFIX}${String(calendarId || "").trim()}`;
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

function getGreyBoxIdForStoredCalendarId(calendarId: string): string {
  const normalizedCalendarId = String(calendarId || "").trim();

  if (!normalizedCalendarId) {
    return "";
  }

  const greyBoxId = normalizeGreyBoxId(
    PropertiesService.getScriptProperties().getProperty(
      getCalendarLookupPropertyKey(normalizedCalendarId)
    ) || ""
  );

  logFunctionSuccess("getGreyBoxIdForStoredCalendarId", {
    calendarId: normalizedCalendarId,
    hasGreyBoxId: Boolean(greyBoxId),
    greyBoxId,
  });

  return greyBoxId;
}

function storeCalendarId(greyBoxId: string, calendarId: string): void {
  const startLog = logFunctionStart("storeCalendarId", {
    greyBoxId,
    calendarId,
  });
  const normalizedGreyBoxId = normalizeGreyBoxId(greyBoxId);
  const normalizedCalendarId = String(calendarId || "").trim();
  const scriptProperties = PropertiesService.getScriptProperties();
  const previousCalendarId = (
    scriptProperties.getProperty(getCalendarPropertyKey(normalizedGreyBoxId)) || ""
  ).trim();

  if (previousCalendarId && previousCalendarId !== normalizedCalendarId) {
    scriptProperties.deleteProperty(getCalendarLookupPropertyKey(previousCalendarId));
  }

  scriptProperties.setProperty(
    getCalendarPropertyKey(normalizedGreyBoxId),
    normalizedCalendarId
  );
  scriptProperties.setProperty(
    getCalendarLookupPropertyKey(normalizedCalendarId),
    normalizedGreyBoxId
  );

  logFunctionSuccess(
    "storeCalendarId",
    {
      greyBoxId: normalizedGreyBoxId,
      calendarId: normalizedCalendarId,
      previousCalendarId,
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
      message.includes("ACL") ||
      message.includes("own access level")
    ) {
      loggerInfo({
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

export {
  getCalendarService,
  normalizeGreyBoxId,
  normalizeEmail,
  isGoogleCalendarId,
  getTrackingSpreadsheetId,
  openTrackingSpreadsheet,
  getGreyBoxDirectory,
  getGreyBoxRecord,
  getMandateMatchByEmail,
  getMandateMatchByIdentifier,
  getStoredCalendarId,
  getGreyBoxIdForStoredCalendarId,
  storeCalendarId,
  buildCalendarUrl,
  shareCalendarWithUser,
  ensureCalendarForRecord,
  getEventTrackingContextForCurrentUserInternal,
};
