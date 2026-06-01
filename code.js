// -----------------------------------------------------------------------------
// constants
// -----------------------------------------------------------------------------

const APPROVER_EMAILS = ["vince.rubinetti@gmail.com"];
const FROM_EMAIL = "vince.rubinetti@gmail.com";
const HELP_CONTACT = "Swathi Thaker at snthaker@uab.edu.";
const FORM_URL =
  "https://docs.google.com/forms/d/1g1rq941ju15Zi2YMv70DDL33giW_xZ7XBrxSuMz8hi0";
const SHEET_URL =
  "https://docs.google.com/spreadsheets/d/1uYt3DBp-PFNTGpBE8r1fssrFnECZ8G13yK1pstT-sgg/edit?usp=sharing";
const CALENDAR_URL = "https://cfdeconnect.org/calendar";
const CALENDAR_ID =
  "c670108b40b3d31f74822cb74014d4b1a592c58e85be478eafeb0e5321b8a4b3@group.calendar.google.com";

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
const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

// columns as object
const columns = Object.fromEntries(
  sheet
    // get first row (headers)
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map((name, index) => ({
      // column key
      key: Object.keys(COLUMNS).find((key) => COLUMNS[key] === name),
      // actual column name in sheet
      name,
      // column number in sheet (1-indexed)
      index: index + 1,
    }))
    // convert to object
    .map((column) => [column.key, column]),
);

// rows as arrays of objects
const rows = sheet
  // 2nd row to last row, 1st column to last column
  .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
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
    start: new Date(row.start),
    end: new Date(row.end),
    // unique key for row, based on stable properties
    key: [row.title, row.start, row.submitter].join("|"),
    // row number in sheet (1-indexed)
    index: rowIndex + 2,
  }))
  .map((row) => ({
    ...row,
    update: (key, value) => {
      // find column by key
      const column = Object.values(columns).find(
        (column) => column.key === key,
      );
      if (!column) return;
      // update value in sheet
      sheet.getRange(row.index, column.index).setValue(value);
      // update value in object (for later use in same execution)
      row[key] = value;
    },
  }));

// calendar object
const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

// script properties
const script = PropertiesService.getScriptProperties();

// -----------------------------------------------------------------------------
// util
// -----------------------------------------------------------------------------

// send email
function sendEmail(to, from, subject, body) {
  MailApp.sendEmail({
    to: [to]
      .flat()
      .filter(Boolean)
      .map((address) => address.trim())
      .join(","),
    from: from.trim(),
    subject: subject.trim(),
    body: body.trim(),
  });
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
const approval = makeRowStore("approval_email");
// track reminder email status for each row
const reminder = makeRowStore("reminder_email");

// look up calendar entry
function getEntry(id) {
  try {
    return calendar.getEventById(id);
  } catch {}
}

// -----------------------------------------------------------------------------
// tasks
// -----------------------------------------------------------------------------

// initialize empty statuses
function initStatuses() {
  for (const row of rows) if (!row.status) row.update("status", "Pending");
}

// send approval request emails for pending events
function sendApprovals() {
  for (const row of rows) {
    // only pending events
    if (row.status !== "Pending") continue;

    // don't resend
    if (approval.get(row) === "sent") continue;

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

Set "Approval Status" to "Approved" or "Denied," and optionally add "Approval Comments.
`;

    sendEmail(APPROVER_EMAILS, FROM_EMAIL, subject, body);

    // mark as sent
    approval.set(row, "sent");
  }
}

// send reminder emails for upcoming approved events
function sendReminders() {
  // current time
  const now = new Date();
  // a bit in the future
  const windowStart = new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000);
  // a bit more in the future
  const windowEnd = new Date(now.getTime() + 15 * 24 * 60 * 60 * 1000);

  for (const row of rows) {
    // only events that start within window
    if (row.start < windowStart || row.start > windowEnd) continue;

    // don't resend
    if (reminder.get(row) === "sent") continue;

    // only approved events
    if (row.status !== "Approved") continue;

    // only events with particular fields
    if (!row.submitter || !row.title || !row.start) continue;

    // email subject
    const subject = `Reminder: Upcoming Event "${row.title}"`;

    // email body
    const body = `
Hello event organizer,

As part of the CFDE Evaluation Core's event reporting efforts, we are asking you to gather some info during your upcoming event. Please refer back to the "POST-EVENT" part of the Google Form where you originally registered your event:

${FORM_URL}

Please prepare to survey your attendees and record notes so that you can answer these questions in detail. Once your event has concluded, we will be reminding you to fill out that section and update your response.

Thank you for helping us demonstrate the impact and value of CFDE events!

${formatDetails(row, ["title", "start", "end"])}
`;

    sendEmail(row.submitter, FROM_EMAIL, subject, body);

    // mark as sent
    reminder.set(row, "sent");
  }
}

// validate event
function validateEvent(row, status) {
  // do some validation that can't easily be achieved in form
  const errors = [];

  // validate link
  if (row.format.match(/virtual|online/i) && !row.link)
    errors.push(`Missing ${columns.link}`);
  // validate location
  if (row.format.match(/in[- ]?person/i) && !row.location)
    errors.push(`Missing ${columns.location}`);

  if (!errors.length) return;

  // report errors
  // reset status so it can be fixed and resubmitted
  row.update("status", "Pending");
  approval.clear(row);

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
  if (status !== "Pending") return;
  approval.clear(row);
  sendApprovals();
  return true;
}

// handle denied event
function eventDenied(row, status) {
  if (status !== "Denied") return;

  if (row.id) {
    // delete calender entry
    getEntry(row.id)?.deleteEvent();
    // clear id from sheet
    row.update("id", "");
  }

  if (!row.submitter) return;

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
  if (status !== "Approved") return;

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
  if (entry)
    // update entry
    entry
      .setTitle(row.title)
      .setTime(row.start, row.end)
      .setDescription(description)
      .setLocation(location);
  else
    // create new entry
    entry = calendar.createEvent(row.title, row.start, row.end, {
      description,
      location,
    });

  // update sheet with calendar entry id
  row.update("id", entry.getId());

  // clear reminder dedupe so updated approved events can get a fresh reminder
  reminder.clear(row);

  if (!row.submitter) return;

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
  const idsKey = "cfde_event_ids";

  // list of all current calendar entry ids in sheet
  const ids = sheet
    .getRange(2, columns.id.index, sheet.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(Boolean);

  // previous persisted list of calendar entry ids
  const previousIds = JSON.parse(script.getProperty(idsKey) || "[]");

  // which ids were removed from sheet
  const removed = previousIds.filter((id) => !ids.includes(id));

  // delete removed entries from calendar
  for (const id of removed) getEntry(id)?.deleteEvent();

  // persist ids for next time
  script.setProperty(idsKey, JSON.stringify(ids));
}

// -----------------------------------------------------------------------------
// triggers
// -----------------------------------------------------------------------------

// run when form submitted
function onFormSubmit() {
  // initialize empty statuses
  initStatuses();
  // send approval request emails
  sendApprovals();
}

// run when spreadsheet changed
function onChange(event) {
  const type = event.changeType;

  if (type === "REMOVE_ROW") {
    eventsRemoved();
    return;
  }

  if (type === "EDIT") {
    // get active range
    const range = sheet.getActiveRange();

    // get row of edited cell
    const rowIndex = range.getRow();
    // get column of edited cell
    const columnIndex = range.getColumn();

    // get new status value
    const status = range.getValue();

    // get full row object
    const row = rows.find((row) => row.index === rowIndex);

    // if status cell edited
    if (rowIndex > 1 && columnIndex === columns.status.index && row) {
      if (validateEvent(row, status)) return;
      if (eventPending(row, status)) return;
      if (eventDenied(row, status)) return;
      if (eventApproved(row, status)) return;
    }

    return;
  }
}

// run daily
function onDaily() {
  // (re)send approval request emails until admin has handled it
  sendApprovals();
  // send reminders dependent on date window
  sendReminders();
}
