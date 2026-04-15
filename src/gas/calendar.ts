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

export type DebugSummary = {
  activeEmail: string;
  effectiveEmail: string;
  isRequesterDeployerSame: boolean;
  mandateEmail: string;
  greyBoxId: string;
  calendarTitle: string;
  storedCalendarId: string;
};

export type LoggerHealthSummary = {
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

const LOG_SPREADSHEET_ID = "1RFcvNfu07jUPpGQan59UcK6NKUqqiTW4mO1Gmtkvdas";
const MANDATE_SPREADSHEET_ID = "1pfvdPNPiJx4XnTUAVzstqENOTYl2GgLTYYea5KMWcpg";
const MANDATE_SHEET_CANDIDATES = ["Mandates", "Mandate"];
const CALENDAR_PROPERTY_PREFIX = "trackedCalendar:";
const TRACKING_SPREADSHEET_PROPERTY_KEY = "TRACKING_SPREADSHEET_ID";
const TRACKED_TIME_SHEET_NAME = "Tracked_Time";
const DEFAULT_EVENT_LIST_LIMIT = 20;
const OPERATIONAL_LOG_SHEET_NAME = "Operational_Log";
const NETWORK_LOG_SHEET_NAME = "Network_Log";
const FALLBACK_OPERATIONAL_HEADERS = [[
  "Timestamp",
  "Log ID",
  "Level",
  "Message",
  "Meta JSON",
]];
const FALLBACK_NETWORK_HEADERS = [[
  "Timestamp",
  "Parent Log ID",
  "Message",
  "System",
  "Method",
  "Status",
  "Duration (ms)",
  "Meta JSON",
]];
let isLoggerInitialized = false;

function ensureSheetExists(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheetName: string
): GoogleAppsScript.Spreadsheet.Sheet {
  const existing = spreadsheet.getSheetByName(sheetName);
  if (existing) {
    return existing;
  }

  return spreadsheet.insertSheet(sheetName);
}

function getLoggerLibRuntime(): Partial<LoggerLibGlobal> {
  return ((globalThis as Record<string, unknown>).LoggerLib ||
    {}) as Partial<LoggerLibGlobal>;
}

function createFallbackOperationalRecord(
  level: LoggerLibLogLevel,
  message: string,
  meta: Record<string, unknown> = {},
  extra: Partial<LoggerLibOperationalLogRecord> = {}
): LoggerLibOperationalLogRecord {
  return {
    kind: "operational",
    ts: new Date(),
    logId: `fallback-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    level,
    message,
    meta,
    correlationId: extra.correlationId,
    parentLogId: extra.parentLogId,
  };
}

function stringifyMeta(value: unknown): string {
  try {
    return JSON.stringify(value ?? {});
  } catch (error) {
    return JSON.stringify({
      message: "Failed to serialize log payload",
      error: error instanceof Error ? error.message : String(error),
    });
  }
}

function ensureFallbackHeaders(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  headers: string[][]
): void {
  if (sheet.getLastRow() > 0) {
    return;
  }

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");
}

function appendFallbackRows(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rows: unknown[][]
): void {
  if (!rows.length) {
    return;
  }

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function appendFallbackOperationalRecord(record: LoggerLibOperationalLogRecord): void {
  try {
    const spreadsheet = SpreadsheetApp.openById(LOG_SPREADSHEET_ID);
    const sheet = ensureSheetExists(spreadsheet, OPERATIONAL_LOG_SHEET_NAME);
    ensureFallbackHeaders(sheet, FALLBACK_OPERATIONAL_HEADERS);
    appendFallbackRows(sheet, [[
      record.ts,
      record.logId,
      record.level,
      record.message,
      stringifyMeta(record.meta),
    ]]);
  } catch (error) {
    console.error("Fallback operational sheet write failed", toErrorMeta(error));
  }
}

function appendFallbackNetworkRecord(record: LoggerLibNetworkLogRecord): void {
  try {
    const spreadsheet = SpreadsheetApp.openById(LOG_SPREADSHEET_ID);
    const sheet = ensureSheetExists(spreadsheet, NETWORK_LOG_SHEET_NAME);
    ensureFallbackHeaders(sheet, FALLBACK_NETWORK_HEADERS);
    appendFallbackRows(sheet, [[
      record.ts,
      record.parentLogId || "",
      record.message,
      record.system,
      record.method,
      record.status,
      record.durationMs,
      stringifyMeta({
        ...record.meta,
        url: record.url.raw,
        ...(record.endpoint ? { endpoint: record.endpoint } : {}),
        ...(record.requestId ? { requestId: record.requestId } : {}),
        ...(typeof record.requestBytes === "number"
          ? { requestBytes: record.requestBytes }
          : {}),
        ...(typeof record.responseBytes === "number"
          ? { responseBytes: record.responseBytes }
          : {}),
        ...(record.error ? { error: record.error } : {}),
      }),
    ]]);
  } catch (error) {
    console.error("Fallback network sheet write failed", toErrorMeta(error));
  }
}

function loggerInfo(input: LoggerLibLogInput): LoggerLibOperationalLogRecord {
  const logger = getLoggerLibRuntime();
  if (typeof logger.info === "function") {
    try {
      return logger.info(input);
    } catch (error) {
      console.warn("LoggerLib.info failed; falling back to direct sheet logging.", toErrorMeta(error));
    }
  }

  console.log("[Logger fallback][INFO]", input.message, input.meta || {});
  const record = createFallbackOperationalRecord("INFO", input.message, input.meta || {}, {
    correlationId: input.correlationId,
    parentLogId: input.parentLogId,
  });
  appendFallbackOperationalRecord(record);
  return record;
}

function loggerError(
  input:
    | LoggerLibLogInput
    | { message: string; error: unknown; meta?: Record<string, unknown> }
): LoggerLibOperationalLogRecord {
  const logger = getLoggerLibRuntime();
  if (typeof logger.error === "function") {
    try {
      return logger.error(input);
    } catch (error) {
      console.warn("LoggerLib.error failed; falling back to direct sheet logging.", toErrorMeta(error));
    }
  }

  const meta = "error" in input
    ? { ...(input.meta || {}), errorDetails: toErrorMeta(input.error) }
    : input.meta || {};
  console.error("[Logger fallback][ERROR]", input.message, meta);
  const record = createFallbackOperationalRecord("ERROR", input.message, meta, {
    correlationId: "correlationId" in input ? input.correlationId : undefined,
    parentLogId: "parentLogId" in input ? input.parentLogId : undefined,
  });
  appendFallbackOperationalRecord(record);
  return record;
}

function loggerHttp(input: {
  system: "GOOGLE" | "SLACK" | "NOTION" | "OTHER";
  method: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
  url: LoggerLibUrlInfo;
  status: number;
  durationMs: number;
  message?: string;
  meta?: Record<string, unknown>;
  correlationId?: string;
  parentLogId?: string;
  requestId?: string;
  requestBytes?: number;
  responseBytes?: number;
  endpoint?: string;
  error?: {
    name?: string;
    message: string;
    code?: string | number;
  };
}): LoggerLibNetworkLogRecord {
  const logger = getLoggerLibRuntime();
  if (typeof logger.http === "function") {
    try {
      return logger.http(input);
    } catch (error) {
      console.warn("LoggerLib.http failed; falling back to direct sheet logging.", toErrorMeta(error));
    }
  }

  const message = input.message || `${input.method} ${input.url.raw}`;
  console.log("[Logger fallback][HTTP]", message, input);
  const record: LoggerLibNetworkLogRecord = {
    kind: "network",
    ts: new Date(),
    logId: `fallback-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    level: "INFO",
    message,
    meta: input.meta || {},
    correlationId: input.correlationId,
    parentLogId: input.parentLogId,
    system: input.system,
    method: input.method,
    url: input.url,
    status: input.status,
    durationMs: input.durationMs,
    requestId: input.requestId,
    requestBytes: input.requestBytes,
    responseBytes: input.responseBytes,
    endpoint: input.endpoint,
    error: input.error,
  };
  appendFallbackNetworkRecord(record);
  return record;
}

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

  try {
    const spreadsheet = SpreadsheetApp.openById(LOG_SPREADSHEET_ID);

    ensureSheetExists(spreadsheet, OPERATIONAL_LOG_SHEET_NAME);
    ensureSheetExists(spreadsheet, NETWORK_LOG_SHEET_NAME);

    const logger = getLoggerLibRuntime();
    if (typeof logger.init === "function") {
      logger.init({
        spreadsheetId: LOG_SPREADSHEET_ID,
        operationalSheet: OPERATIONAL_LOG_SHEET_NAME,
        networkSheet: NETWORK_LOG_SHEET_NAME,
        level: "DEBUG",
      });
    } else {
      console.warn("LoggerLib.init is unavailable; continuing without explicit initialization.");
    }

    isLoggerInitialized = true;
  } catch (error) {
    isLoggerInitialized = false;
    console.error("Logger initialization failed", toErrorMeta(error));
    throw error;
  }
}

function getLoggerHealthInternal(): LoggerHealthSummary {
  let spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null = null;

  try {
    spreadsheet = SpreadsheetApp.openById(LOG_SPREADSHEET_ID);
  } catch (error) {
    return {
      ok: false,
      spreadsheetId: LOG_SPREADSHEET_ID,
      operationalSheet: OPERATIONAL_LOG_SHEET_NAME,
      networkSheet: NETWORK_LOG_SHEET_NAME,
      operationalSheetExists: false,
      networkSheetExists: false,
      loggerInitialized: isLoggerInitialized,
      failureStage: "open_spreadsheet",
      errorMessage: error instanceof Error ? error.message : String(error),
    };
  }

  try {
    ensureSheetExists(spreadsheet, OPERATIONAL_LOG_SHEET_NAME);
  } catch (error) {
    return {
      ok: false,
      spreadsheetId: LOG_SPREADSHEET_ID,
      operationalSheet: OPERATIONAL_LOG_SHEET_NAME,
      networkSheet: NETWORK_LOG_SHEET_NAME,
      operationalSheetExists: Boolean(spreadsheet.getSheetByName(OPERATIONAL_LOG_SHEET_NAME)),
      networkSheetExists: Boolean(spreadsheet.getSheetByName(NETWORK_LOG_SHEET_NAME)),
      loggerInitialized: isLoggerInitialized,
      failureStage: "ensure_operational_sheet",
      errorMessage: error instanceof Error ? error.message : String(error),
    };
  }

  try {
    ensureSheetExists(spreadsheet, NETWORK_LOG_SHEET_NAME);
  } catch (error) {
    return {
      ok: false,
      spreadsheetId: LOG_SPREADSHEET_ID,
      operationalSheet: OPERATIONAL_LOG_SHEET_NAME,
      networkSheet: NETWORK_LOG_SHEET_NAME,
      operationalSheetExists: Boolean(spreadsheet.getSheetByName(OPERATIONAL_LOG_SHEET_NAME)),
      networkSheetExists: Boolean(spreadsheet.getSheetByName(NETWORK_LOG_SHEET_NAME)),
      loggerInitialized: isLoggerInitialized,
      failureStage: "ensure_network_sheet",
      errorMessage: error instanceof Error ? error.message : String(error),
    };
  }

  try {
    const logger = getLoggerLibRuntime();
    if (typeof logger.init === "function") {
      logger.init({
        spreadsheetId: LOG_SPREADSHEET_ID,
        operationalSheet: OPERATIONAL_LOG_SHEET_NAME,
        networkSheet: NETWORK_LOG_SHEET_NAME,
        level: "DEBUG",
      });
    }
    isLoggerInitialized = true;
  } catch (error) {
    return {
      ok: false,
      spreadsheetId: LOG_SPREADSHEET_ID,
      operationalSheet: OPERATIONAL_LOG_SHEET_NAME,
      networkSheet: NETWORK_LOG_SHEET_NAME,
      operationalSheetExists: Boolean(spreadsheet.getSheetByName(OPERATIONAL_LOG_SHEET_NAME)),
      networkSheetExists: Boolean(spreadsheet.getSheetByName(NETWORK_LOG_SHEET_NAME)),
      loggerInitialized: isLoggerInitialized,
      failureStage: "logger_lib_init",
      errorMessage: error instanceof Error ? error.message : String(error),
    };
  }

  return {
    ok: true,
    spreadsheetId: LOG_SPREADSHEET_ID,
    operationalSheet: OPERATIONAL_LOG_SHEET_NAME,
    networkSheet: NETWORK_LOG_SHEET_NAME,
    operationalSheetExists: Boolean(spreadsheet.getSheetByName(OPERATIONAL_LOG_SHEET_NAME)),
    networkSheetExists: Boolean(spreadsheet.getSheetByName(NETWORK_LOG_SHEET_NAME)),
    loggerInitialized: isLoggerInitialized,
  };
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
  return loggerInfo({
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
  return loggerInfo({
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
  return loggerError({
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
        `Mandate sheet must include Greybox ID and Email (Org) columns.`
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
        `Mandate sheet must include Greybox ID and Email (Org) columns.`
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

function getMandateMatchByIdentifier(identifier: string): MandateMatch {
  const normalized = (identifier || "").trim();

  if (!normalized) {
    throw new Error("Identifier is required.");
  }

  if (normalized.includes("@")) {
    return getMandateMatchByEmail(normalized);
  }

  const normalizedGreyBoxId = normalizeGreyBoxId(normalized);
  const greyBoxRecord = getGreyBoxDirectory()[normalizedGreyBoxId];
  if (greyBoxRecord) {
    return { record: greyBoxRecord };
  }

  const records = Object.values(getGreyBoxDirectory());
  for (const record of records) {
    const storedCalendarId = getStoredCalendarId(record.greyBoxId);
    if (storedCalendarId && storedCalendarId === normalized) {
      return { record };
    }
  }

  throw new Error(
    `No mandate match found for identifier: ${normalized}`
  );
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

function formatEventTimeRange(startAt: Date, endAt: Date): string {
  const timeZone = Session.getScriptTimeZone();
  const startText = Utilities.formatDate(startAt, timeZone, "yyyy-MM-dd HH:mm");
  const endText = Utilities.formatDate(endAt, timeZone, "yyyy-MM-dd HH:mm");
  return `${startText} to ${endText}`;
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

    const attendees = (event.attendees || []).filter(
      (attendee) => Boolean((attendee.email || "").trim())
    );

    if (attendees.length === 0) {
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

export function listAllEventsForCurrentUser() {
  logInit();
  const context = getEventTrackingContextForCurrentUserInternal();
  const events = listAllEventsForCalendar(context.calendarId);
  writeValidatedGreyBoxSheet(context.greyBoxId, events);
  return events;
}

export function listAllEventsForIdentifier(identifier: string) {
  logInit();
  const match = getMandateMatchByIdentifier(identifier);
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

export function insertEvent(
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

export function start(): void {
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

export function getCurrentUserEmail(): string {
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

export function getLoggerHealth(): LoggerHealthSummary {
  return getLoggerHealthInternal();
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

export function getDebugSummary(): DebugSummary {
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

export function syncTrackedTimeToSheet() {
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

export function syncAllTrackedTimeToSheet() {
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
