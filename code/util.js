// send email
function sendEmail(to, from, subject, body, dryRun = false) {
  to = [to]
    .flat()
    .filter(Boolean)
    .map((address) => address.trim())
    .join(",");
  from = from.trim();
  subject = subject.trim();
  body = body.trim();
  console.log({ to, from, subject, body });
  if (dryRun) console.log("dry run, not sending");
  else MailApp.sendEmail({ to, from, subject, body });
}

// format certain details in row as multi-line string
function formatDetails(
  row,
  keys,
  valueSeparator = " — ",
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

// look up calendar entry
function getEntry(id) {
  try {
    return calendar.getEventById(id);
  } catch {}
}

// get row corresponding to form response
function getResponseRow(response) {
  console.log({ response: response.getId() });

  // timestamp of form response
  const responseTime = response.getTimestamp().getTime();
  for (const row of rows) {
    // timestamp of row
    const rowTime = row.timestamp.getTime();
    // match by timestamp (loosely, ms omitted)
    if (Math.abs(rowTime - responseTime) < 2000) {
      console.log("matched response to row", {
        row: row.index,
        responseTime,
        rowTime,
      });
      return row;
    }
  }
}

// today, +/- n days
function now(days) {
  return new Date(new Date().getTime() + days * 24 * 60 * 60 * 1000);
}

// get rows in date window
function rowsInWindow(start, end) {
  // relative to today
  start = now(start);
  end = now(end);
  const inWindow = rows.filter((row) => row.start >= start && row.start <= end);
  console.log({ start, end, inWindow: inWindow.length });
  return inWindow;
}

// stringify args for logging
function stringify(...args) {
  return args
    .map((arg) => (typeof arg === "object" ? JSON.stringify(arg) : arg))
    .join(" ");
}

// automatically add debug logging to all functions
for (const name of Object.keys(globalThis)) {
  const original = globalThis[name];
  if (typeof original === "function")
    globalThis[name] = function (...args) {
      console.log(`${name}()`);
      console.log("arguments", args);
      const result = original.apply(this, args);
      console.log("result", result);
      return result;
    };
}
