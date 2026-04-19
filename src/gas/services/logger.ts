/// <reference types="google-apps-script" />
/// <reference path="../types/logger-lib.d.ts" />

import type { LoggerHealthSummary } from "./types";

const LOG_SPREADSHEET_ID = "1RFcvNfu07jUPpGQan59UcK6NKUqqiTW4mO1Gmtkvdas";
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

export {
  loggerInfo,
  loggerError,
  loggerHttp,
  logInit,
  getLoggerHealthInternal,
  toErrorMeta,
  logFunctionStart,
  logFunctionSuccess,
  logFunctionError,
};
