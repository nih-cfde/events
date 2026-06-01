# Setup

Run `initStatuses()` manually once to authorize permissions.

Show `appsscript.json` in editor via settings and set it to:

```json
{
  "timeZone": "America/New_York",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/script.send_mail"
  ]
}
```

Copy `code.js` to editor.

Add triggers:

| Function       | Event Source     | Event Type     |
| -------------- | ---------------- | -------------- |
| `onFormSubmit` | From spreadsheet | On form submit |
| `onChange`     | From spreadsheet | On change      |
| `onDaily`      | Time-driven      | Day timer      |
