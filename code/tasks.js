// -----------------------------------------------------------------------------
// action-based
// -----------------------------------------------------------------------------

// initialize new event
function initEvent(row) {
  // send email to admin
  const subject = newSubject(row);
  const body = newBody(row);
  sendEmail(ADMIN_EMAIL, ADMIN_EMAIL, subject, body);

  // set status to pending
  row.update("status", "Pending");

  // auto-validate event
  const errors = validateEvent(row);
  if (errors.length) {
    // add errors as comment
    if (!row.comments) row.update("comments", errors.join("\n"));
    row.update("status", "Denied");
  }
}

// do some validation that can't easily be achieved in form
function validateEvent(row) {
  const errors = [];

  // validate link
  if (row.format.match(/virtual|online/i) && !row.link)
    errors.push(`Missing "${columns.link.name}"`);
  // validate location
  if (row.format.match(/in[- ]?person/i) && !row.location)
    errors.push(`Missing "${columns.location.name}"`);

  return errors;
}

// remove calendar entry
function removeEntry(row) {
  // delete calendar entry
  getEntry(row.id)?.deleteEvent();
  // clear id from sheet
  row.update("id", "");
}

// add/update calendar entry
function updateEntry(row) {
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
}

// deny event
function denyEvent(row) {
  removeEntry(row);

  if (!row.submitter) return console.log("no submitter, ignoring");

  // send email to user
  const subject = deniedSubject(row);
  const body = deniedBody(row);
  sendEmail(row.submitter, ADMIN_EMAIL, subject, body);
}

// approve event
function approveEvent(row) {
  updateEntry(row);

  if (!row.submitter) return console.log("no submitter, ignoring");

  // send email to user
  const subject = approvedSubject(row);
  const body = approvedBody(row);
  sendEmail(row.submitter, ADMIN_EMAIL, subject, body);
}

// handle event(s) removed by admin
function handleRemoved() {
  const idsKey = "cfde_event_ids";

  // list of all current calendar entry ids in sheet
  const current = rows.map((row) => row.id).filter(Boolean);
  console.log({ current });

  // previous persisted list of calendar entry ids
  const old = JSON.parse(script.getProperty(idsKey) || "[]");
  console.log({ old });

  // which ids were removed from sheet
  const removed = old.filter((id) => !current.includes(id));
  console.log({ removed });

  // delete removed entries from calendar
  for (const id of removed) {
    console.log("removing calendar entry", { id });
    getEntry(id)?.deleteEvent();
  }

  // persist ids for next time
  script.setProperty(idsKey, JSON.stringify(current));
}

// -----------------------------------------------------------------------------
// time-based
// -----------------------------------------------------------------------------

// send reminder emails for upcoming events
function sendUpcoming() {
  // few days in future
  let rows = rowsInWindow(14, 14 + 1);

  // only approved events
  rows = rows.filter((row) => row.status === "Approved");

  // only events with necessary fields
  rows = rows.filter((row) => row.submitter);

  for (const row of rows) {
    // send email to user
    const subject = upcomingSubject(row);
    const body = upcomingBody(row);
    sendEmail(row.submitter, ADMIN_EMAIL, subject, body);
  }
}

// send reminder emails for recent events
function sendRecent() {
  // few days in past
  let rows = rowsInWindow(-7, -7 + 1);

  // only approved events
  rows = rows.filter((row) => row.status === "Approved");

  // only events with necessary fields
  rows = rows.filter((row) => row.submitter);

  for (const row of rows) {
    // send email to user
    const subject = recentSubject(row);
    const body = recentBody(row);
    sendEmail(row.submitter, ADMIN_EMAIL, subject, body);
  }
}
