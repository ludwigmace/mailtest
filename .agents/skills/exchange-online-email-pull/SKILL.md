---
name: exchange-online-email-pull
description: Use when implementing, extending, or debugging Exchange Online email retrieval and mailbox synchronization through Microsoft Graph. Covers delegated authentication (MSAL), paging, delta queries, JSON batching, attachment downloads (file/item types), local storage, processing flags, or opt-in PDF export.
---

# Exchange Online Email Pulling

Use this skill to build, maintain, or troubleshoot a local email puller for a user's Exchange Online mailbox through Microsoft Graph. Adapt the patterns below to the host project's language, framework, configuration, and storage conventions; do not assume the project has the same CLI or database schema as this repository.

## Start by inspecting the project

Before editing, locate its authentication/configuration code, Graph client, message and attachment persistence, processing hook, command-line entry point, and tests. Check its README and local agent instructions. Keep responsibilities separate where practical:

- Authentication and configuration
- Graph HTTP client, pagination, and batching
- Message/attachment persistence
- Sync orchestration and processing status
- Optional export or user-interface features

Do not overwrite local configuration, token caches, mailbox data, or unrelated user changes.

## Authentication and permissions

- For a signed-in user's mailbox, use Microsoft Authentication Library (MSAL) with a **public client** and delegated Graph permissions. Use `/me` endpoints for the authenticated user; do not substitute app-only credentials for a user-authorized workflow.
- Use the narrowest delegated scope that supports the requested operations: `Mail.Read` for read-only pulling; `Mail.ReadWrite` when the workflow also updates follow-up flags or other message properties. Explain the permission and consent requirement to the user.
- Keep tenant/authority, client ID, scopes, and other settings configurable. Commit a template with placeholders, never real credentials. Keep access tokens and token caches out of source control. On Windows, protect persistent MSAL cache data with Windows DPAPI when available.
- Handle silent-token failure by using the project's interactive sign-in flow. Do not print access tokens, authorization headers, message bodies, or attachment contents in logs.

For a Python implementation, the core MSAL pattern is:

```python
import msal

app = msal.PublicClientApplication(
    client_id,
    authority=authority,
    token_cache=persistent_cache,
)
accounts = app.get_accounts()
result = app.acquire_token_silent(scopes, account=accounts[0]) if accounts else None
if not result or "access_token" not in result:
    result = app.acquire_token_interactive(scopes=scopes)
if not result or "access_token" not in result:
    detail = (result or {}).get("error_description", "Token acquisition failed")
    raise RuntimeError(detail)
access_token = result["access_token"]
```

Use an Entra app registration configured for the intended account types and public-client interactive sign-in. Never put a real client secret in a desktop or local mailbox puller.

## Microsoft Graph retrieval

- Use Microsoft Graph `v1.0` and request only the message fields the workflow needs, such as `id`, `internetMessageId`, `subject`, sender/recipient fields, `body`, timestamps, `flag`, and `hasAttachments`.
- The usual collection endpoint is `GET https://microsoft.com`. Send query parameters such as `$top`, `$select`, and (for an unflagged queue) `$filter` as HTTP parameters rather than concatenating unescaped user input into a URL.
- If messages may move between folders, send `Prefer: IdType="ImmutableId"` and use the returned immutable message ID as the durable key in the local store.
- Send the acquired token as `Authorization: Bearer <access-token>`; never log that header. Raise or handle HTTP errors before parsing a response as JSON.
- Follow every `@odata.nextLink` exactly as returned; do not construct later page URLs from the initial query. Handle Graph throttling and transient failures using `Retry-After` where provided and bounded retries with backoff.
- For high-volume mailboxes or full incremental synchronization, prefer Microsoft Graph Delta Queries (`/messages/delta`) over scanning entire folders.
- For a workflow that pulls the unflagged queue, query `flag/flagStatus eq 'notFlagged'`. Some Graph responses/tenants may reject a nested-property filter with HTTP 400. In that specific case, page through the messages and filter locally; do not turn permission, authentication, or server errors into an unbounded fallback scan.
- Consider utilizing JSON Batching (`POST /v1.0/$batch`) when fetching attachment metadata for multiple messages simultaneously to minimize HTTP overhead and avoid throttling.
- Retrieve attachments per message from its `/attachments` endpoint. Do not rely only on `hasAttachments`: inline attachments may be omitted from that summary field. Preserve attachment order and metadata such as `name`, `contentType`, `size`, `isInline`, and `contentId`.
- The attachment listing route is `GET /v1.0/me/messages/{message-id}/attachments`. 
- **Important Endpoint Quirk:** Downloading raw bytes using the `/$value` endpoint (`GET /v1.0/me/messages/{message-id}/attachments/{attachment-id}/$value`) only works if the attachment's `@odata.type` is `#microsoft.graph.fileAttachment`. If the type is `#microsoft.graph.itemAttachment` (an attached email or calendar item), do not append `/$value`; retrieve the nested attachment object structure directly via JSON. URL-encode IDs as path segments and retain the immutable ID preference header.
- Stream attachment bytes from the Graph endpoint to a temporary file, close the response, then atomically replace the destination. Record a checksum and download error/status so interrupted downloads can be retried. 

## Local storage and processing

- Upsert messages using the immutable Graph ID; avoid duplicate rows when a message is synced again. Use a composite `(message_id, attachment_id)` key for attachments. Persist normalized fields needed for processing plus raw Graph JSON when it is useful for later work.
- Treat sync as retryable stages: message upsert, attachment download, downstream processing, then Graph flag update. Mark a message complete only after required downloads and processing succeed. Keep failures eligible for retry and preserve actionable error details.
- Make downstream effects idempotent. Use the immutable message ID as an idempotency key when posting to another service, writing an archive, or triggering a job; a failed run may execute again.
- In an unflagged-queue workflow, valid terminal outcomes are typically `flagged` or `complete`. Avoid returning `notFlagged` after successful processing because that can select the message again on the next sync.
- Do not add PDF work to the sync path when the user requests opt-in PDF generation. Store the email and downloadable attachments during sync, then let the user select stored message IDs from a separate CLI/TUI action. Keep PDF generation failures separate from Graph sync and flag state.

## Optional PDF export

When requested by the host project, render a print-style email header/body, resolve HTML `cid:` references to downloaded inline images, convert supported attachments, and append them in original attachment order. Append existing PDFs without re-rendering. Conversion tools differ by format; document any system dependencies (for example, LibreOffice for Office documents) and fail clearly for unsupported formats rather than silently creating an incomplete PDF.

## Security and operational behavior

- Treat email HTML and attachments as untrusted input. Do not execute macros or active content; avoid fetching remote resources from message HTML. Restrict local file reads during HTML rendering to known attachment directories.
- **Path Traversal Mitigation:** Sanitize attachment names strictly. Use message and attachment IDs in storage paths to avoid collisions. Ensure that the final target path resolved by the file system remains strictly inside the designated download directory (e.g., explicitly verify and block path manipulation attempts like `../../`).
- Keep mailbox data, downloaded attachments, PDFs, configuration, and token caches out of Git. Do not include message data in test fixtures unless it is synthetic.
- Avoid destructive remote operations unless the user explicitly asks for them. A local delete action should state clearly whether it removes only the local database/files or also affects Exchange.
- Keep the tool useful without live credentials in tests. Mock Graph responses, MSAL, and file downloads; test pagination, fallback filtering, idempotent upserts, attachment retries, status updates, and user-selected actions.

## Common Graph Edge Cases & Troubleshooting

- **HTTP 429 (Too Many Requests) / HTTP 503 (Service Unavailable):** Inspect the response headers for `Retry-After`. Implement exponential backoff with jitter based on that value.
- **HTTP 400 (Bad Request on Filters):** If filtering by a nested property like `flag/flagStatus` fails due to tenant-specific schema restrictions, fall back gracefully to client-side filtering after pulling the active page, and log an operational warning.
- **HTTP 401 (Unauthorized):** Usually indicates the access token expired mid-sync. Catch this error gracefully, invoke `acquire_token_silent` to refresh the token, and retry the failed request exactly once before failing the sync stage.
