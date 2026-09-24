// -----------------------------------------------------------------------------
// constants
// -----------------------------------------------------------------------------

const ADMIN_EMAIL = "cfde.icc@gmail.com";

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
  description: "Description",
  organizer: "Organizer",
  involved: "Involved",
  length: "Length",
  start: "Start",
  end: "End",
  link: "Link",
  format: "Format",
  location: "Location",
  purpose: "Purpose",
  tags: "Tags",

  // other
  status: "Status",
  selected: "Selected",
  comments: "Comments",
  edit: "Edit",
  id: "ID",
  timestamp: "Timestamp",
  submitter: "Email Address",
  source: "Source",
};

// -----------------------------------------------------------------------------
// globals
// -----------------------------------------------------------------------------

// sheet object
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

// get row/col counts
const lastRow = sheet.getLastRow();
const lastColumn = sheet.getLastColumn();

// sheet name
const sheetName = sheet.getName();

console.log({ lastRow, lastColumn, sheetName });

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
    console.log("updating row", { row: row.index, key, value });
    // find column by key
    const column = Object.values(columns).find((column) => column.key === key);
    if (!column)
      throw Error(
        stringify("no matching column", { row: row.index, key, value }),
      );
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
