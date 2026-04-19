import { build } from "esbuild";

const GAS_BUILD_OUTFILE = "dist/Code.gs";

const GAS_EXPOSED_FUNCTIONS = Object.freeze([
  "hello",
  "start",
  "doGet",
  "getCurrentUserEmail",
  "getLoggerHealth",
  "getIdentitySummary",
  "getDebugSummary",
  "listAllEventsForCalendarId",
  "listAllEventsForCalendarUrl",
  "listAllEventsForEmail",
  "listAllEventsForGreyBoxId",
  "listAllEventsForIdentifier",
  "listAllEventsForCurrentUser",
  "createCalendarForRequester",
  "requestCalendarForCurrentUser",
  "requestCalendarByGreyBoxId",
  "listAvailableGreyBoxIds",
  "syncAllTrackedTimeToSheet",
  "syncTrackedTimeToSheet"
]);

function buildGasFooter() {
  return GAS_EXPOSED_FUNCTIONS.map(
    (name) =>
      `function ${name}() { return globalThis._${name}.apply(globalThis, arguments); }`
  ).join("\n");
}

async function buildGas() {
  await build({
    entryPoints: ["src/gas/entrypoints/index.ts"],
    bundle: true,
    platform: "browser",
    target: "es2019",
    format: "iife",
    outfile: GAS_BUILD_OUTFILE,
    sourcemap: false,
    legalComments: "none",
    footer: {
      js: buildGasFooter()
    }
  });
}

export { GAS_BUILD_OUTFILE, GAS_EXPOSED_FUNCTIONS, buildGas, buildGasFooter };
