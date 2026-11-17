# Gmail Attachments → Drive + Sheets + Daily Digest (Google Apps Script)

For anyone who receives vendor quotes, invoices or order confirmations by email
and spends time saving attachments, renaming them and keeping a list.

**Every 15 minutes** the script finds new emails with attachments, saves each
PDF / Excel / Word / CSV file to Drive in a folder per month, renames it to
`<date> <sender> - <original name>`, and logs it in a Google Sheet with sender,
subject, the PO number it finds in the email, size and a Drive link. Threads get
an `auto-filed` label and message IDs are remembered, so nothing is filed twice.

**Every evening** it emails a digest of the day: how many files from whom, the
PO numbers mentioned, and a link to each file.

## Setup (5 minutes)

1. Go to [script.google.com](https://script.google.com) → **New project**, paste `Code.gs`.
2. Adjust `CONFIG` at the top if needed (search query, allowed file types, digest hour).
3. Select `setupTriggers` → **Run** → approve the Gmail, Drive and Sheets permissions.
   It installs both schedules and does the first run straight away.
4. A sheet called **Attachment Log** appears in your Drive; files land in **Email Attachments/**.

## Tests

The parsing and formatting logic (PO number extraction, sender names, safe file
names, digest text) is plain JavaScript with no Google calls, and is covered by
unit tests that run locally:

```bash
node --test tests/
```

## Design choices

- **Idempotent**: a label on the thread plus a message-ID column in the sheet,
  so re-runs never file the same message twice. A new reply in a thread that is
  already labelled is skipped; remove the label to have the thread re-checked.
- **Batch writes**: all new rows are written to the sheet in one call, which
  keeps each run well inside Apps Script's time limit.
- **No secrets in code**: the sheet ID is stored in Script Properties after the
  first run.
