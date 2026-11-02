# Outlook Email Puller

A local Python tool that syncs Microsoft Graph emails whose follow-up flag is
`notFlagged`, stores their details and attachments, runs a processing hook, and
updates the Graph flag so processed emails are not picked up again. PDF creation
is opt-in: choose stored emails explicitly from the command line.

## System design

`pull_emails.py` coordinates authentication, sync, persistence, processing, and
flag updates:

1. MSAL checks the encrypted local token cache and opens a browser only when
   interactive authentication or new consent is needed.
2. `graph_client.py` paginates through messages flagged `notFlagged`. If Graph
   rejects filtering by the nested flag status, it falls back to paging and
   filtering locally. It requests immutable Graph message IDs so IDs remain
   stable when messages move between folders.
3. `email_store.py` upserts each message and its attachment metadata in SQLite.
   File and item attachments are downloaded to disk; repeated syncs reuse a
   downloaded file if it is still present. Reference attachments are recorded
   but cannot be downloaded by this tool.
4. `email_processing.py` receives the stored email record and returns a
   processing status, which sync applies to Graph. Sync does not create a PDF.
5. `--convert-to-pdf` selects one or more stored messages for PDF creation.
   `email_to_pdf.py` renders the Outlook-style header and body, resolves embedded
   `cid:` images, converts supported attachments (including attached `.eml`
   messages), and appends them in attachment order. PDFs go to `data/pdfs/`.

### Local database

The default database is `data/emails.sqlite3`; attachment files are under
`data/attachments/` and generated message PDFs are under `data/pdfs/`. These
data files are excluded from Git.

- **`emails`** — Graph ID (primary key), internet message ID, subject, sender,
  to/cc/bcc/reply-to recipients, body and content type, received/sent dates,
  read/importance/conversation metadata, flag and processing status, generated
  PDF path, timestamps, errors, and the raw Graph message JSON.
- **`attachments`** — email Graph ID and attachment ID (composite primary key),
  name, content type, size, inline/content-ID metadata, attachment type and
  original order, local path, SHA-256, download status, and errors. Each
  attachment references its email with a foreign key.

Emails are uniquely keyed by immutable Graph ID. Re-syncing a message upserts
that row rather than creating a duplicate; attachment rows are also upserted by
Graph attachment ID.

## Setup and usage

1. Install dependencies:

   ```powershell
   pip install -r requirements.txt
   ```

   Microsoft Office documents (`.docx`, `.xlsx`, `.pptx`, `.rtf`, and compatible
   formats) require LibreOffice installed with `soffice` on `PATH`. PDFs are
   appended as-is; text, HTML, `.eml`, and common image formats are converted
   locally. Unsupported attachment types (such as Outlook `.msg` files) or
   failed conversions leave the stored email and Graph flag unchanged; no
   partial PDF is saved.

2. Copy `config.example.json` to `config.json` and set `client_id` to the
   Application (client) ID from your Entra app registration. The example uses
   the `consumers` authority for personal Microsoft accounts. Configure the app
   to allow personal accounts; use your tenant authority for a work/school-only
   registration.
3. Ensure the app has the delegated Microsoft Graph `Mail.ReadWrite` permission.
   The script requests this scope; the first run may prompt for sign-in and
   consent. The token cache is stored in `token_cache.bin` and protected with
   Windows DPAPI for the current Windows user.
4. Run a sync (this is the default command):

   ```powershell
   python pull_emails.py
   ```

Other commands:

```powershell
python pull_emails.py --tui
python pull_emails.py --list
python pull_emails.py --list-stored
python pull_emails.py --convert-to-pdf "<GRAPH_MESSAGE_ID>"
python pull_emails.py --convert-to-pdf "<GRAPH_MESSAGE_ID_1>" "<GRAPH_MESSAGE_ID_2>"
python pull_emails.py --message-id "<GRAPH_MESSAGE_ID>" --flag-status flagged
python pull_emails.py --message-id "<GRAPH_MESSAGE_ID>" --flag-status notFlagged
python pull_emails.py --message-id "<GRAPH_MESSAGE_ID>" --flag-status complete
```

Manually setting `notFlagged` queues the message for the next sync. `--list`
prints recent Exchange message IDs for use with `--message-id`. After syncing,
use `--list-stored` to find local message IDs, then pass one or more of them to
`--convert-to-pdf`. Listing and conversion use the local database and do not
require Microsoft Graph authentication.

### Terminal UI

Run `python pull_emails.py --tui` to browse synced messages. Select a row with
the arrow keys to view sender, recipients, dates, status, attachment names, a
body preview, and PDF availability. Use:

- **P** — generate or regenerate the selected email PDF.
- **D** — confirm and delete the local database entry, downloaded attachments,
  and generated PDF. This does not delete the message from Exchange.
- **S** — authenticate if needed, pull currently `notFlagged` messages from
  Exchange, download their attachments, and refresh the list when sync finishes.
- **R** — refresh the list.
- **Q** — quit.

Browsing, PDF conversion, and deletion use the local database without Graph
authentication. Pressing **S** uses the same Graph sign-in as
`python pull_emails.py`; a browser may open if interactive sign-in is needed.

## Development notes for future changes

- Keep Graph calls in `graph_client.py`, database operations/schema in
  `email_store.py`, sync orchestration in `pull_emails.py`, and downstream
  decision logic in `email_processing.py`.
- Preserve idempotency: use immutable Graph IDs and attachment IDs as keys,
  atomically write downloaded files, and avoid duplicate rows on retries.
- If the SQLite schema changes, add a migration path for existing
  `data/emails.sqlite3` files. Do not commit `config.json`, `token_cache.bin`,
  the `data/` directory, or real mailbox content.
- Test with mocked Graph responses and temporary SQLite/filesystem paths; avoid
  requiring live credentials for unit tests. Run tests with:

  ```powershell
  python -m unittest discover -s tests -v
  ```

### Implementing `email_processing.py`

`process_email(email)` runs during sync after the original message and its
attachments have been downloaded. PDF generation is a separate, opt-in command,
so `pdf_path` is normally `None` when the hook runs. The `email` argument is a
dictionary from `EmailStore.get_email()` with decoded values:

- Message metadata such as `graph_id`, `subject`, `sender`, `to_recipients`,
  `cc_recipients`, `bcc_recipients`, `sent_datetime`, and `received_datetime`.
- `body_content` and `body_content_type` (`text` or `html`), plus `raw_message`
  containing the original Graph message.
- `attachments`, ordered as returned by Graph. Each row includes its name,
  content type, inline/content-ID metadata, `local_path`, `download_status`,
  checksum, and attachment index.
- `pdf_path`, which is `None` until a PDF has been generated with
  `--convert-to-pdf`; afterward it contains the local PDF path.

The hook must return exactly `"complete"` or `"flagged"`:

- Return `"complete"` after successful downstream work to mark the message
  complete in Graph.
- Return `"flagged"` when the message should remain flagged for follow-up.
- Do not return `"notFlagged"`: that is the sync queue state and would make the
  message eligible for the next sync again.

Example starting point for processing metadata independently of PDF creation:

```python
def process_email(email):
    # Replace this with downstream processing of the message or attachments.
    # Use email["graph_id"] as an idempotency key for external actions.
    return "complete"
```

Let processing errors raise so the sync records the failure and leaves the
message eligible for retry. Make external side effects idempotent, because a
failed run may be attempted again. Use `--convert-to-pdf` separately when a
combined PDF is needed; PDF rendering belongs in `email_to_pdf.py`.
