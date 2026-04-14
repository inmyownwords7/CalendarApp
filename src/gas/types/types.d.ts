type LoggerLibLogInput = {
  message: string;
  meta?: Record<string, unknown>;
  correlationId?: string;
  parentLogId?: string;
};

type LoggerLibUrlInfo = {
  raw: string;
  parts?: {
    base?: string;
    path: readonly string[];
    query?: Record<string, string>;
  };
};

type LoggerLibLogLevel = "DEBUG" | "INFO" | "WARN" | "ERROR";

interface LoggerLibBaseLogRecord {
  ts: Date;
  logId: string;
  correlationId?: string;
  parentLogId?: string;
  level: LoggerLibLogLevel;
  message: string;
  meta: Record<string, unknown>;
}

interface LoggerLibOperationalLogRecord extends LoggerLibBaseLogRecord {
  kind: "operational";
}

interface LoggerLibNetworkLogRecord extends LoggerLibBaseLogRecord {
  kind: "network";
  system: "GOOGLE" | "SLACK" | "NOTION" | "OTHER";
  method: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
  url: LoggerLibUrlInfo;
  status: number;
  durationMs: number;
  requestId?: string;
  requestBytes?: number;
  responseBytes?: number;
  endpoint?: string;
  error?: {
    name?: string;
    message: string;
    code?: string | number;
  };
}

interface LoggerLibGlobal {
  init(config: {
    spreadsheetId: string;
    operationalSheet?: string;
    networkSheet?: string;
    operationalHeaders?: string[];
    networkHeaders?: string[];
    level?: LoggerLibLogLevel;
  }): void;

  log(input: LoggerLibLogInput): LoggerLibOperationalLogRecord;
  debug(input: LoggerLibLogInput): LoggerLibOperationalLogRecord;
  info(input: LoggerLibLogInput): LoggerLibOperationalLogRecord;
  warn(input: LoggerLibLogInput): LoggerLibOperationalLogRecord;
  error(
    input:
      | LoggerLibLogInput
      | { message: string; error: unknown; meta?: Record<string, unknown> }
  ): LoggerLibOperationalLogRecord;

  http(input: {
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
  }): LoggerLibNetworkLogRecord;
}

declare const LoggerLib: LoggerLibGlobal;
