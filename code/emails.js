// -----------------------------------------------------------------------------
// new event, to admin
// -----------------------------------------------------------------------------

const newSubject = (row) => `New event submission: "${row.title}"`;

const newBody = (row) => `
A new event has been submitted:
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

Please review row # ${row.index} here and set "Status" and optionally "Comments":
${SHEET_URL}
`;

// -----------------------------------------------------------------------------
// common
// -----------------------------------------------------------------------------

const closer = `
Thank you for helping us demonstrate the impact and value of CFDE events! If you have any questions, please reply to this email.

- CFDE Integration and Coordination Center
`;

// -----------------------------------------------------------------------------
// event approved, to submitter
// -----------------------------------------------------------------------------

const approvedSubject = (row) => `Your event "${row.title}" was approved`;

const approvedBody = (row) => `
Hello event organizer,

Your event has been approved and is now live on our shared calendar!

Calendar:
${CALENDAR_URL}

Edit your form response:
${row.edit || FORM_URL}

Event details:
${formatDetails(row, ["title", "start", "end"])}

${closer}
`;

// -----------------------------------------------------------------------------
// event denied, to submitter
// -----------------------------------------------------------------------------

const deniedSubject = (row) => `Your event "${row.title}" was not approved`;

const deniedBody = (row) => `
Hello event organizer,

Your event was not approved:
${row.comments}

Edit your form response:
${row.edit || FORM_URL}

Event details:
${formatDetails(row, ["title", "start", "end"])}

${closer}
`;

// -----------------------------------------------------------------------------
// upcoming event reminder, to submitter
// -----------------------------------------------------------------------------

const upcomingSubject = (row) =>
  `Prepare for your upcoming event "${row.title}"`;

const upcomingBody = (row) => `
Hello event organizer,

As part of the CFDE Evaluation Core's event reporting efforts, we are asking you to gather some info during your upcoming event. Please refer back to the "POST-EVENT" part of the Google Form where you originally registered your event:

Edit your form response:
${row.edit || FORM_URL}

Event details:
${formatDetails(row, ["title", "start", "end"])}

Please prepare to survey your attendees and record notes so that you can answer these questions in detail. Once your event has concluded, we will remind you to fill out that section and update your response.

${closer}
`;

// -----------------------------------------------------------------------------
// recent event reminder, to submitter
// -----------------------------------------------------------------------------

const recentSubject = (row) =>
  `Answer questions about your recent event "${row.title}"`;

const recentBody = (row) => `
Hello event organizer,

As part of the CFDE Evaluation Core's event reporting efforts, we are asking you to answer some questions about your recent event. Please refer back to the "POST-EVENT" part of the Google Form where you originally registered your event:

Edit your form response:
${row.edit || FORM_URL}

Event details:
${formatDetails(row, ["title", "start", "end"])}

Please fill out that section and update your original response.

${closer}
`;
