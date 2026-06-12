# Setup

Show `appsscript.json` in editor and copy local `appsscript.json` content.

Copy local `code.js` content to editor.

Add triggers:

| Function       | Event Source     | Event Type     |
| -------------- | ---------------- | -------------- |
| `onFormSubmit` | From spreadsheet | On form submit |
| `onChange`     | From spreadsheet | On change      |
| `onDaily`      | Time-driven      | Day timer      |

Run `initStatuses()` manually once to authorize permissions.
