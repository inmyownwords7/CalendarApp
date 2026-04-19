# CalendarApp

Google Apps Script web app for provisioning tracked calendars, listing calendar events, and writing event/attendance data into Google Sheets.

## What This App Does

This app is built around a mandate spreadsheet that maps a person to a Grey Box ID and a tracking calendar.

Current supported flow:

1. A user requests a tracking calendar.
2. The app creates or reuses a stored Google Calendar for that user.
3. The app can list events for a target person by:
   - email
   - Grey Box ID
   - calendar ID
4. The app writes event rows into a per-user sheet.
5. The app writes synced event rows into a `Tracked_Time` sheet.

It also includes debug helpers for identity and logging.

## Main Concepts

### Mandate Sheet

The app reads a spreadsheet that contains the user-to-Grey Box mapping.

Current accepted sheet names:

- `Mandates`
- `Mandate`

Current accepted required headers:

- `Greybox ID`
- `Email (Org)`

Optional title aliases:

- `calendarTitle`
- `Calendar Title`
- `Title`

### Tracking Calendar

Each person can have one stored tracked calendar. The calendar ID is stored in Apps Script script properties using a key derived from the Grey Box ID.

### Tracking Sheets

The app writes tracking data to a spreadsheet chosen in this order:

1. Script property `TRACKING_SPREADSHEET_ID`
2. `LOG_SPREADSHEET_ID`

That means if `TRACKING_SPREADSHEET_ID` is not set, the tracking sheets are created in the same spreadsheet as the log sheets.

### Log Sheets

Operational and network logs are written to:

- `Operational_Log`
- `Network_Log`

If the external logger library fails at runtime, the app falls back to writing logs directly to those sheets.

## Current Sheet Outputs

### Per-user Event Sheet

Each target person gets a sheet named after their Grey Box ID.

Current columns:

- `eventName (title)`
- `eventId`
- `startTime`
- `endTime`
- `duration`
- `attendeeEmail`
- `responseStatus`
- `wasPresent`
- `attendanceMarkedBy`
- `attendanceMarkedAt`

Notes:

- `responseStatus` comes from Google Calendar attendee data.
- `wasPresent`, `attendanceMarkedBy`, and `attendanceMarkedAt` are intended to be maintained manually.
- Those manual fields are preserved across sync runs.
- A total row is written at the bottom with total hours.

### Tracked_Time Sheet

The `Tracked_Time` sheet is the normalized sync sheet.

Current columns:

- `greyBoxId`
- `email`
- `calendarId`
- `eventId`
- `summary`
- `startIso`
- `endIso`
- `durationMinutes`
- `durationHours`
- `syncedAt`

## Event Listing Behavior

### Recurring Events

Recurring events are expanded using Google Calendar `singleEvents: true`.

Current event listing limit:

- `20` events maximum per list/sync call

This cap exists to keep recurring event series from flooding the sheet. It can be made configurable later.

### Attendance / RSVP

The app currently supports pulling attendee RSVP information from Calendar:

- attendee email
- response status such as `accepted`, `declined`, `tentative`, `needsAction`

The intended model is:

- Google Calendar provides RSVP state
- a human later marks actual attendance using `wasPresent`

## Entry Functions

These functions are exposed through [src/gas/entrypoints/index.ts](./src/gas/entrypoints/index.ts) on the `globalThis.__calendarApp` namespace for web app and Apps Script usage.

### Calendar Provisioning

- `requestCalendarForCurrentUser()`
  Creates or reuses the signed-in user's tracking calendar.

- `requestCalendarByGreyBoxId(greyBoxId)`
  Creates or reuses the target Grey Box user's tracking calendar.

### Event Listing

- `listAllEventsForCurrentUser()`
  Lists the current user's tracked calendar events and repopulates that user's sheet.

- `listAllEventsForEmail(email)`
  Lists events for the person matched by mandate email.

- `listAllEventsForGreyBoxId(greyBoxId)`
  Lists events for the matched Grey Box ID.

- `listAllEventsForCalendarId(calendarId)`
  Lists events directly from the supplied calendar ID.

- `listAllEventsForCalendarUrl(calendarUrl)`
  Extracts a calendar ID from a Google Calendar URL and lists events for it.

- `listAllEventsForIdentifier(identifier)`
  Auto-detects the identifier type and lists events for:
  - email
  - Grey Box ID
  - stored calendar ID
  - Google Calendar ID

This remains the generic fallback function, but the web UI now prefers explicit mode-specific entrypoints.

### Sync

- `syncTrackedTimeToSheet()`
  Syncs only the signed-in user's tracked calendar into `Tracked_Time` and the per-user sheet.

- `syncAllTrackedTimeToSheet()`
  Admin-style full sync across every stored calendar currently known to the app.

### Debug / Diagnostics

- `getCurrentUserEmail()`
- `getIdentitySummary()`
- `getDebugSummary()`
- `getLoggerHealth()`
- `listAvailableGreyBoxIds()`

## Web App UI

The web UI is defined in [index.html](./index.html).

### Page-load Calls

- `getIdentitySummary()`
  Loads the current requester/effective-user identity state.

- `listAvailableGreyBoxIds()`
  Loads Grey Box IDs for datalist suggestions and Grey Box prefix autocomplete.

### Buttons

- `Request calendar`
  Calls `requestCalendarForCurrentUser()`

- `List events`
  Uses the selected `Lookup by` mode and calls one of:
  - `listAllEventsForGreyBoxId()`
  - `listAllEventsForEmail()`
  - `listAllEventsForCalendarId()`
  - `listAllEventsForCalendarUrl()`
  - `listAllEventsForIdentifier()` when `Auto` is selected

- `Check current user`
  Calls `getCurrentUserEmail()`

- `Check logger`
  Calls `getLoggerHealth()`

- `Debug identity`
  Calls `getDebugSummary()`

### UI To GAS Mapping

| UI control | GAS function | Source |
| --- | --- | --- |
| Page load: identity bootstrap | `getIdentitySummary()` | `src/gas/services/calendar.ts` |
| Page load: Grey Box suggestions | `listAvailableGreyBoxIds()` | `src/gas/services/calendar.ts` |
| `Request calendar` | `requestCalendarForCurrentUser()` | `src/gas/services/calendar.ts` |
| `List events` with `Grey Box ID` | `listAllEventsForGreyBoxId()` | `src/gas/services/calendar.ts` |
| `List events` with `Email` | `listAllEventsForEmail()` | `src/gas/services/calendar.ts` |
| `List events` with `Calendar ID` | `listAllEventsForCalendarId()` | `src/gas/services/calendar.ts` |
| `List events` with `Calendar URL` | `listAllEventsForCalendarUrl()` | `src/gas/services/calendar.ts` |
| `List events` with `Auto` | `listAllEventsForIdentifier()` | `src/gas/services/calendar.ts` |
| `Check current user` | `getCurrentUserEmail()` | `src/gas/services/calendar.ts` |
| `Check logger` | `getLoggerHealth()` | `src/gas/services/calendar.ts` |
| `Debug identity` | `getDebugSummary()` | `src/gas/services/calendar.ts` |

### `List events` Input Modes

The `Lookup by` selector changes which backend function is called:

- `Grey Box ID`
- `Email`
- `Calendar ID`
- `Calendar URL`
- `Auto`

Grey Box suggestions and prefix autocomplete are active for `Grey Box ID` and `Auto`.

## Build And Deploy

### Local Build

```bash
npm run build:gas
```

This builds the GAS bundle to:

- `dist/Code.gs`

### Push To Apps Script

```bash
npm run clasp:push
```

### Combined Build + Push

```bash
npm run build
```

This runs:

1. `npm run build:gas`
2. `clasp push`

### Important Deployment Note

`clasp push` updates project files, but it does not automatically publish a new web app version.

After pushing, you still need to update the web app deployment in Apps Script:

1. Open Apps Script
2. `Deploy`
3. `Manage deployments`
4. Edit the target web app deployment
5. Point it to the latest version
6. Use that deployment URL

If the browser still shows old UI code, you are almost always testing the wrong deployment URL or an older deployment version.

## Project Structure

- [src/gas/services/calendar.ts](./src/gas/services/calendar.ts)
  Main business logic for identity, calendars, event listing, sync, logging, and sheet writes.

- [src/gas/entrypoints/index.ts](./src/gas/entrypoints/index.ts)
  Registers bundle handlers on `globalThis` for the esbuild footer wrappers and `globalThis.__calendarApp`.

- [index.html](./index.html)
  Web app UI.

- [scripts/build-gas.mjs](./scripts/build-gas.mjs)
  esbuild script for generating `dist/Code.gs`.

- [scripts/gas-build.mjs](./scripts/gas-build.mjs)
  Shared GAS build config and the list of Apps Script functions exposed in the build footer.

- [appsscript.json](./appsscript.json)
  Apps Script manifest.

## Apps Script Runtime Requirements

### Advanced Services

The app uses the Advanced Calendar service:

- `Calendar` v3

This must also be enabled in the linked Google Cloud project.

### Library

The app references a logger library in `appsscript.json` as `LoggerLib`.

If the library is misconfigured or missing methods at runtime, the app now falls back to direct sheet logging so the main flows still work.

## Known Design Choices

- The app is currently configured as a web app that executes as the deploying user.
- Identity resolution uses `Session.getActiveUser()` and `Session.getEffectiveUser()`.
- Sharing a calendar to its owner is treated as a no-op.
- Event listing is intentionally capped at 20 entries for now.
- Per-user event sheets are attendance-oriented, not raw event dumps.

## Future Work

Planned or likely next improvements:

- dropdown to choose event list size
- UI support for event listing by identifier without using the Apps Script editor
- master event sheet across all `eventId` values
- explicit attendance review workflow
- caching mandate-sheet reads to reduce latency
- clearer admin-only vs self-service actions
