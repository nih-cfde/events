// run when form submitted/edited by user
function onFormSubmit({ response }) {
  // get row corresponding to form response
  const row = getResponseRow(response);
  console.debug({ row: row?.index });
  if (!row) throw Error("no matching row");

  // init form response edit link
  const edit = response.getEditResponseUrl();
  if (edit && !row.edit) row.update("edit", edit);

  // completely new event
  if (row.status === "") initEvent(row);
}

// run when spreadsheet changed
function onChange({ changeType: type, source }) {
  // event(s) removed by admin
  if (type === "REMOVE_ROW") handleRemoved();

  // cell edited by admin or user
  if (type === "EDIT") {
    // get range of edited cells
    const range = source.getActiveRange();
    if (!range) throw Error("no range");

    // make sure on right sheet
    if (range.getSheet().getName() !== sheetName) throw Error("wrong sheet");

    // get row of edited cell
    const rowIndex = range.getRow();
    // get column of edited cell
    const columnIndex = range.getColumn();
    // get new cell value
    const value = range.getValue();
    console.debug({ rowIndex, columnIndex, value });

    // get full row object
    const row = rows.find((row) => row.index === rowIndex);
    if (!row) throw Error("no matching row");
    console.debug({ row: row.index });

    // admin changed status column
    if (columnIndex === columns.status.index) {
      if (row.status === "Denied") denyEvent(row);
      if (row.status === "Approved") approveEvent(row);
    }
    // admin or user changed other column
    else {
      // update calendar entry
      if (row.status === "Approved") updateEntry(row);
    }
  }
}

// run daily
function onDaily() {
  // send reminders dependent on date window
  sendUpcoming();
  sendRecent();
}

// install triggers and authorize permissions
function install() {
  // clear existing triggers so re-running doesn't create duplicates
  for (const trigger of ScriptApp.getProjectTriggers())
    ScriptApp.deleteTrigger(trigger);

  // run when form submitted
  ScriptApp.newTrigger("onFormSubmit").forForm(form).onFormSubmit().create();
  // run when spreadsheet changed
  ScriptApp.newTrigger("onChange")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange()
    .create();
  // run daily
  ScriptApp.newTrigger("onDaily").timeBased().everyDays(1).create();
}
