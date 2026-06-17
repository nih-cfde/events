// -----------------------------------------------------------------------------
// constants
// -----------------------------------------------------------------------------

const APPROVER_EMAILS = ["cfde.icc@gmail.com"];
const FROM_EMAIL = "cfde.icc@gmail.com";
const HELP_CONTACT = "Swathi Thaker at snthaker@uab.edu.";
const FORM_URL =
  "https://docs.google.com/forms/d/1g1rq941ju15Zi2YMv70DDL33giW_xZ7XBrxSuMz8hi0";
const FORM_ID = "1g1rq941ju15Zi2YMv70DDL33giW_xZ7XBrxSuMz8hi0";
const SHEET_URL =
  "https://docs.google.com/spreadsheets/d/1uYt3DBp-PFNTGpBE8r1fssrFnECZ8G13yK1pstT-sgg";
const CALENDAR_URL = "https://cfdeconnect.org/calendar";
const CALENDAR_ID =
  "3e81b0f8035b4e83b9394300144fcd290c034f9a5c9d1d00a3bb2a71d143785d@group.calendar.google.com";

// map of column "keys" (var names used in this script) to "names" (actual header cell values in sheet)
// (so this script can be fixed more easily if form/sheet names change)
const COLUMNS = {
  // form
  title: "Title",
  organizer: "Organizer",
  involved: "Involved",
  purpose: "Purpose",
  length: "Length",
  start: "Start",
  end: "End",
  link: "Link",
  tags: "Tags",
  format: "Format",
  location: "Location",
  description: "Description",

  // other
  timestamp: "Timestamp",
  submitter: "Email Address",
  source: "Source",
  status: "Status",
  comments: "Comments",
  id: "ID",
};

// -----------------------------------------------------------------------------
// globals
// -----------------------------------------------------------------------------

// sheet object
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

// get row/col counts
const lastRow = sheet.getLastRow();
const lastColumn = sheet.getLastColumn();

console.log({ lastRow, lastColumn });

// columns as object
const columns = Object.fromEntries(
  sheet
    // get first row (headers)
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map((name, index) => ({
      // column key
      key: Object.keys(COLUMNS).find((key) => COLUMNS[key] === name) || name,
      // actual column name in sheet
      name,
      // column number in sheet (1-indexed)
      index: index + 1,
    }))
    // convert to object
    .map((column) => [column.key, column]),
);

console.log({ columns });

// rows as arrays of objects
const rows =
  lastRow <= 1
    ? []
    : sheet
        // 2nd row to last row, 1st column to last column
        .getRange(2, 1, lastRow - 1, lastColumn)
        // get values as 2D array
        .getValues()
        // map each row array to an object with keys from headers
        .map((row) =>
          Object.fromEntries(
            Object.values(columns).map(({ key, index }) => [
              // make object keys match column keys
              key,
              // fallback if cell value null/undefined
              row[index - 1] ?? "",
            ]),
          ),
        )
        // add more helpful properties to each row object
        .map((row, rowIndex) => ({
          ...row,
          // parse dates as date objects
          timestamp: new Date(row.timestamp),
          start: new Date(row.start),
          end: new Date(row.end),
          // unique key for row, based on stable properties
          key: [row.title, row.start, row.submitter].join("|"),
          // row number in sheet (1-indexed)
          index: rowIndex + 2,
        }));

// add update method for each row for convenient setting of values
for (const row of rows)
  row.update = (key, value) => {
    // find column by key
    const column = Object.values(columns).find((column) => column.key === key);
    if (!column) {
      console.debug("no matching column", { row: row.index, key, value });
      return;
    }
    // update value in sheet
    sheet.getRange(row.index, column.index).setValue(value);
    // update value in object (for later use in same execution)
    row[key] = value;
  };

console.log({ rows: rows.length });

// calendar object
const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

// script properties
const script = PropertiesService.getScriptProperties();

// form object
const form = FormApp.openById(FORM_ID);

// form responses
let responses;

// -----------------------------------------------------------------------------
// util
// -----------------------------------------------------------------------------

// send email
function sendEmail(to, from, subject, body) {
  console.log("sendEmail");
  to = [to]
    .flat()
    .filter(Boolean)
    .map((address) => address.trim())
    .join(",");
  from = from.trim();
  subject = subject.trim();
  body = body.trim();
  console.log({ to, from, subject, body });
  MailApp.sendEmail({ to, from, subject, body });
}

// format certain details in row as multi-line string
function formatDetails(
  row,
  keys,
  valueSeparator = ": ",
  entrySeparator = "\n",
) {
  return (
    keys
      // label/value
      .map((key) => [COLUMNS[key], row[key]])
      // show that field exists but just empty
      .map(([key, value]) =>
        [
          key,
          value && value instanceof Date
            ? // format as date
              value.toDateString()
            : // format as string
              (value ?? "-"),
        ].join(valueSeparator),
      )
      .join(entrySeparator)
  );
}

// use script properties as simple persistent key-value store for particular row
function makeRowStore(prefix) {
  const getKey = (row) => `${prefix}:${row.key}`;
  const set = (row, value) => script.setProperty(getKey(row), value);
  const get = (row) => script.getProperty(getKey(row));
  const clear = (row) => script.deleteProperty(getKey(row));
  return { get, set, clear };
}

// track approval email status for each row
const approvalEmail = makeRowStore("approval_email");
// track upcoming event reminder email status for each row
const upcomingEmail = makeRowStore("upcoming_email");
// track recent event reminder email status for each row
const recentEmail = makeRowStore("recent_email");

// look up calendar entry
function getEntry(id) {
  try {
    return calendar.getEventById(id);
  } catch {}
}

// get link to edit form response
function getEditLink(row) {
  console.log("getEditLink");
  console.log({ row: row.index });

  // get responses
  if (!responses) responses = form.getResponses();

  // timestamp of row
  const rowTime = row.timestamp.getTime();
  for (const response of responses) {
    // timestamp of form response
    const responseTime = response.getTimestamp().getTime();
    // match by timestamp (loosely, ms sometimes omitted)
    if (Math.abs(responseTime - rowTime) < 2000) {
      console.log(`matched row ${rowTime} to response ${responseTime}`);
      return response.getEditResponseUrl();
    }
  }

  return "";
}

// today, +/- n days
function now(days) {
  return new Date(new Date().getTime() + days * 24 * 60 * 60 * 1000);
}

// -----------------------------------------------------------------------------
// tasks
// -----------------------------------------------------------------------------

// initialize empty statuses
function initStatuses() {
  console.log("initStatuses");

  for (const row of rows) {
    console.log(`row ${row.index}`);

    if (!row.status) {
      row.update("status", "Pending");
      console.log("set to pending");
    }
  }
}

// send approval request emails for pending events
function sendApprovals() {
  console.log("sendApprovals");

  for (const row of rows) {
    console.log(`row ${row.index}`);

    // only pending events
    if (row.status !== "Pending") {
      console.log("not pending, ignoring");
      continue;
    }

    // don't resend
    if (approvalEmail.get(row) === "sent") {
      console.log("already sent, ignoring");
      continue;
    }

    // email subject
    const subject = `New Event Submission: ${row.title || "(no title)"}`;

    // email body
    const body = `
A new event awaits your review:

${formatDetails(row, [
  "title",
  "start",
  "end",
  "purpose",
  "location",
  "link",
  "length",
  "involved",
  "submitter",
  "source",
])}

Please review row # ${row.index} here:
${SHEET_URL}

Set "Approval Status" to "Approved" or "Denied," and optionally add "Approval Comments".
`;

    sendEmail(APPROVER_EMAILS, FROM_EMAIL, subject, body);

    // mark as sent
    approvalEmail.set(row, "sent");
    console.log("marked as sent");
  }
}

// send reminder emails for upcoming events
function sendUpcoming() {
  console.log("sendUpcoming");

  // some days in future
  const windowStart = now(14);
  // some more days in future
  const windowEnd = now(15);

  console.log({ windowStart, windowEnd });

  for (const row of rows) {
    console.log(`row ${row.index}`);

    // only events that start within window
    if (row.start < windowStart || row.start > windowEnd) {
      console.log("outside of window, ignoring");
      continue;
    }

    // only approved events
    if (row.status !== "Approved") {
      console.log("not approved, ignoring");
      continue;
    }

    // don't resend
    if (upcomingEmail.get(row) === "sent") {
      console.log("already sent, ignoring");
      continue;
    }

    // only events with necessary fields
    if (!row.submitter) {
      console.log("missing fields, ignoring");
      continue;
    }

    // email subject
    const subject = `Prepare for your upcoming event "${row.title}"`;

    // email body
    const body = `
Hello event organizer,

As part of the CFDE Evaluation Core's event reporting efforts, we are asking you to gather some info during your upcoming event. Please refer back to the "POST-EVENT" part of the Google Form where you originally registered your event:

${getEditLink(row) || FORM_URL}

Please prepare to survey your attendees and record notes so that you can answer these questions in detail. Once your event has concluded, we will be reminding you to fill out that section and update your response.

Thank you for helping us demonstrate the impact and value of CFDE events!

${formatDetails(row, ["title", "start", "end"])}
`;

    sendEmail(row.submitter, FROM_EMAIL, subject, body);

    // mark as sent
    upcomingEmail.set(row, "sent");
    console.log("marked as sent");
  }
}

// send reminder emails for recent events
function sendRecent() {
  console.log("sendRecent");

  // some days in past
  const windowEnd = now(-3);
  // some more days in past
  const windowStart = now(-2);

  console.log({ windowStart, windowEnd });

  for (const row of rows) {
    console.log(`row ${row.index}`);

    // only events that end within window
    if (row.end < windowStart || row.end > windowEnd) {
      console.log("outside of window, ignoring");
      continue;
    }

    // only approved events
    if (row.status !== "Approved") {
      console.log("not approved, ignoring");
      continue;
    }

    // don't resend
    if (recentEmail.get(row) === "sent") {
      console.log("already sent, ignoring");
      continue;
    }

    // only events with necessary fields
    if (!row.submitter) {
      console.log("missing fields, ignoring");
      continue;
    }

    // email subject
    const subject = `Answer questions about your recent event "${row.title}"`;

    // email body
    const body = `
Hello event organizer,

As part of the CFDE Evaluation Core's event reporting efforts, we are asking you to answer some questions about your recent event. Please refer back to the "POST-EVENT" part of the Google Form where you originally registered your event:

${getEditLink(row) || FORM_URL}

Please fill out that section and update your original response.

Thank you for helping us demonstrate the impact and value of CFDE events!

${formatDetails(row, ["title", "start", "end"])}
`;

    sendEmail(row.submitter, FROM_EMAIL, subject, body);

    // mark as sent
    recentEmail.set(row, "sent");
    console.log("marked as sent");
  }
}

// validate event
function validateEvent(row, status) {
  console.log("validateEvent");
  console.log({ row: row.index, status });

  // do some validation that can't easily be achieved in form
  const errors = [];

  // validate link
  if (row.format.match(/virtual|online/i) && !row.link)
    errors.push(`Missing ${columns.link.name}`);
  // validate location
  if (row.format.match(/in[- ]?person/i) && !row.location)
    errors.push(`Missing ${columns.location.name}`);

  if (!errors.length) {
    console.log("no errors, ignoring");
    return;
  }

  // reset status so it can be fixed and resubmitted
  approvalEmail.clear(row);
  row.update("status", "Pending");
  console.log("set to pending");

  // email subject
  const subject = `Approval blocked (row ${row.index})`;

  // email body
  const body = `
Cannot approve row ${row.index} in ${SHEET_URL}.

Errors:
${errors.map((error) => `- ${error}`).join("\n")}
`;

  sendEmail(APPROVER_EMAILS, FROM_EMAIL, subject, body);

  return true;
}

// handle setting status to pending
function eventPending(row, status) {
  console.log("eventPending");
  console.log({ row: row.index, status });

  if (status !== "Pending") {
    console.log("not pending, ignoring");
    return;
  }
  approvalEmail.clear(row);
  sendApprovals();

  return true;
}

// handle denied event
function eventDenied(row, status) {
  console.log("eventDenied");

  if (status !== "Denied") {
    console.log("not denied, ignoring");
    return;
  }

  if (row.id) {
    console.log("deleting calendar entry");
    // delete calender entry
    getEntry(row.id)?.deleteEvent();
    // clear id from sheet
    row.update("id", "");
  }

  if (!row.submitter) {
    console.log("no submitter, ignoring");
    return;
  }

  // email subject
  const subject = `Your event "${row.title}" was denied`;

  // email body
  const body = `
Your CFDE event submission was not approved. Comments:

${row.comments}

If you have any questions, please contact ${HELP_CONTACT}.
`;

  sendEmail(row.submitter, FROM_EMAIL, subject, body);

  return true;
}

// handle approved event
function eventApproved(row, status) {
  console.log("eventApproved");
  console.log({ row: row.index, status });

  if (status !== "Approved") {
    console.log("not approved, ignoring");
    return;
  }

  // make calendar entry description to (PUBLIC, do not include private info)
  const description = formatDetails(row, [
    "description",
    "tags",
    "organizer",
    "involved",
    "purpose",
    "length",
    "link",
  ]);

  // get location for calendar entry
  const location =
    row.format.match(/virtual|online/i) && row.link ? row.link : row.location;

  // get existing calendar entry
  let entry = getEntry(row.id);

  // if exists
  if (entry) {
    console.log("updating existing calendar entry");
    // update entry
    entry
      .setTitle(row.title)
      .setTime(row.start, row.end)
      .setDescription(description)
      .setLocation(location);
  } else {
    console.log("creating new calendar entry");
    // create new entry
    entry = calendar.createEvent(row.title, row.start, row.end, {
      description,
      location,
    });
  }

  // update sheet with calendar entry id
  row.update("id", entry.getId());

  if (!row.submitter) {
    console.log("no submitter, ignoring");
    return;
  }

  // email subject
  const subject = `Your event "${row.title}" was approved`;

  // email body
  const body = `
Your event is now live on our shared calendar: ${CALENDAR_URL}

If you have any questions, please contact ${HELP_CONTACT}.
`;

  sendEmail(row.submitter, FROM_EMAIL, subject, body);

  return true;
}

// handle removed events
function eventsRemoved() {
  console.log("eventsRemoved");

  const idsKey = "cfde_event_ids";

  // list of all current calendar entry ids in sheet
  const _new =
    lastRow <= 1
      ? []
      : sheet
          .getRange(2, columns.id.index, lastRow - 1, 1)
          .getValues()
          .flat()
          .filter(Boolean);
  console.log({ new: _new });

  // previous persisted list of calendar entry ids
  const old = JSON.parse(script.getProperty(idsKey) || "[]");
  console.log({ old });

  // which ids were removed from sheet
  const removed = old.filter((id) => !_new.includes(id));
  console.log({ removed });

  // delete removed entries from calendar
  for (const id of removed) {
    console.log(`removing calendar entry ${id}`);
    getEntry(id)?.deleteEvent();
  }

  // persist ids for next time
  script.setProperty(idsKey, JSON.stringify(_new));
}

// -----------------------------------------------------------------------------
// triggers
// -----------------------------------------------------------------------------

// run when form submitted
function onFormSubmit() {
  console.log("onFormSubmit");

  // initialize empty statuses
  initStatuses();
  // send approval request emails
  sendApprovals();
}

// run when spreadsheet changed
function onChange(event) {
  console.log("onChange");

  const type = event.changeType;
  console.log({ type });

  if (type === "REMOVE_ROW") {
    console.log("removed row");
    eventsRemoved();
    return;
  }

  if (type === "EDIT") {
    console.log("edited cell");

    // range of cells just edited
    const range = event.source.getActiveRange();

    // make sure range is valid
    if (!range || range.getSheet().getName() !== sheet.getName()) return;

    // get row of edited cell
    const rowIndex = range.getRow();
    // get column of edited cell
    const columnIndex = range.getColumn();
    // get new cell value
    const value = range.getValue();

    console.log({ row: rowIndex, column: columnIndex, value });

    // get full row object
    const row = rows.find((row) => row.index === rowIndex);

    console.log({ row });

    // if status cell edited
    if (rowIndex > 1 && columnIndex === columns.status.index && row) {
      console.log("status edited");
      if (validateEvent(row, value)) return;
      if (eventPending(row, value)) return;
      if (eventDenied(row, value)) return;
      if (eventApproved(row, value)) return;
    }

    return;
  }
}

// run daily
function onDaily() {
  console.log("onDaily");

  // (re)send approval request emails until admin has handled it
  sendApprovals();
  // send reminders dependent on date window
  sendUpcoming();
  sendRecent();
}
