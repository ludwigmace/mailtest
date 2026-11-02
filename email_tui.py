"""Terminal interface for browsing synced emails and managing local PDFs."""

import html
import re
from pathlib import Path

from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import Horizontal, Vertical
from textual.screen import ModalScreen
from textual import work
from textual.widgets import Button, DataTable, Footer, Header, Static

from email_store import EmailStore
from pull_emails import (
    DATABASE_PATH,
    GraphClient,
    acquire_access_token,
    convert_email_to_pdf,
    delete_email_locally,
    load_config,
    sync_emails,
)


class ConfirmDeleteScreen(ModalScreen[bool]):
    CSS = """
    ConfirmDeleteScreen { align: center middle; }
    #confirm-dialog {
        width: 64;
        height: auto;
        padding: 1 2;
        border: round $error;
        background: $surface;
    }
    #confirm-actions { height: auto; margin-top: 1; align-horizontal: right; }
    #confirm-actions Button { margin-left: 1; }
    """

    def __init__(self, subject):
        super().__init__()
        self.subject = subject

    def compose(self) -> ComposeResult:
        with Vertical(id="confirm-dialog"):
            yield Static(
                f"Delete the local copy of:\n{self.subject}\n\n"
                "This removes its database record, downloaded attachments, and PDF. "
                "It does not delete the Exchange message.",
                markup=False,
            )
            with Horizontal(id="confirm-actions"):
                yield Button("Cancel", id="cancel-delete")
                yield Button("Delete local copy", id="confirm-delete", variant="error")

    def on_button_pressed(self, event: Button.Pressed) -> None:
        self.dismiss(event.button.id == "confirm-delete")


class EmailTui(App):
    TITLE = "Outlook Email Browser"
    SUB_TITLE = "Local synced messages and PDFs"
    BINDINGS = [
        Binding("p", "generate_pdf", "Generate / regenerate PDF"),
        Binding("d", "delete_local", "Delete local copy"),
        Binding("s", "sync_server", "Sync from Exchange"),
        Binding("r", "refresh", "Refresh list"),
        Binding("q", "quit", "Quit"),
    ]
    CSS = """
    #help-line { height: 1; color: $text-muted; padding-left: 1; }
    #email-table { height: 1fr; }
    #email-details {
        height: 11;
        border: round $primary;
        padding: 0 1;
        overflow-y: auto;
    }
    #status-line { height: 2; padding-left: 1; }
    ConfirmDeleteScreen { layer: modal; }
    """

    def __init__(self, store):
        super().__init__()
        self.store = store
        self.emails = {}
        self.selected_id = None
        self._syncing = False

    def compose(self) -> ComposeResult:
        yield Header()
        yield Static(
            "Select a row. Press P to generate/regenerate its PDF, D to delete "
            "the local copy, S to sync from Exchange, R to refresh.",
            id="help-line",
            markup=False,
        )
        yield DataTable(id="email-table", cursor_type="row", zebra_stripes=True)
        yield Static("Select an email to view details.", id="email-details", markup=False)
        yield Static("", id="status-line", markup=False)
        yield Footer()

    def on_mount(self) -> None:
        table = self.query_one(DataTable)
        table.add_columns("Received", "From", "Subject", "Flag", "State", "Files", "PDF")
        self.refresh_emails()
        table.focus()

    def refresh_emails(self, select_id=None) -> None:
        table = self.query_one(DataTable)
        old_selection = select_id or self.selected_id
        table.clear(columns=False)
        self.emails.clear()

        for email in self.store.list_recent_emails(limit=5000):
            graph_id = email["graph_id"]
            self.emails[graph_id] = email
            sender = (email.get("sender") or {}).get("emailAddress") or {}
            sender_text = sender.get("name") or sender.get("address") or "Unknown"
            subject = email.get("subject") or "(No Subject)"
            received = (email.get("received_datetime") or "")[:16].replace("T", " ")
            pdf_path = email.get("pdf_path")
            pdf_status = "Ready" if pdf_path and Path(pdf_path).is_file() else "—"
            table.add_row(
                received,
                sender_text,
                subject,
                email.get("flag_status") or "—",
                email.get("processing_status") or "—",
                str(email.get("attachment_count", 0)),
                pdf_status,
                key=graph_id,
            )

        if not self.emails:
            self.selected_id = None
            self.query_one("#email-details", Static).update(
                "No synced emails. Run python pull_emails.py to sync messages."
            )
            return

        self.selected_id = old_selection if old_selection in self.emails else next(iter(self.emails))
        row_index = list(self.emails).index(self.selected_id)
        table.move_cursor(row=row_index, column=0)
        self.update_details(self.selected_id)

    def on_data_table_row_highlighted(
        self, event: DataTable.RowHighlighted
    ) -> None:
        graph_id = getattr(event.row_key, "value", event.row_key)
        self.selected_id = str(graph_id)
        self.update_details(self.selected_id)

    def update_details(self, graph_id) -> None:
        email = self.store.get_email(graph_id)
        if email is None:
            return
        sender = _format_person(email.get("sender"))
        to = _format_recipients(email.get("to_recipients")) or "—"
        cc = _format_recipients(email.get("cc_recipients")) or "—"
        attachments = email.get("attachments", [])
        attachment_lines = [
            f"  • {item.get('name') or '(unnamed)'} — {item.get('download_status')}"
            for item in attachments
        ] or ["  (none)"]
        pdf_path = email.get("pdf_path")
        pdf_text = str(pdf_path) if pdf_path and Path(pdf_path).is_file() else "Not generated"
        body = email.get("body_content") or ""
        if (email.get("body_content_type") or "").lower() == "html":
            body = re.sub(r"(?is)<(script|style)\b[^>]*>.*?</\1\s*>", " ", body)
            body = re.sub(r"<[^>]+>", " ", body)
        body = " ".join(html.unescape(body).split())
        if len(body) > 240:
            body = body[:237] + "..."

        details = (
            f"Subject: {email.get('subject') or '(No Subject)'}\n"
            f"From: {sender}\n"
            f"To: {to}\n"
            f"Cc: {cc}\n"
            f"Received: {email.get('received_datetime') or '—'}\n"
            f"Flag / processing: {email.get('flag_status') or '—'} / "
            f"{email.get('processing_status') or '—'}\n"
            f"PDF: {pdf_text}\n"
            f"Attachments ({len(attachments)}):\n"
            + "\n".join(attachment_lines)
            + f"\nPreview: {body or '(empty)'}"
        )
        self.query_one("#email-details", Static).update(details)

    def _set_status(self, message):
        self.query_one("#status-line", Static).update(message)

    def action_refresh(self) -> None:
        if self._syncing:
            self._set_status("Wait for the server sync to finish before refreshing.")
            return
        self.refresh_emails()
        self._set_status(f"Loaded {len(self.emails)} stored emails.")

    def action_sync_server(self) -> None:
        if self._syncing:
            self._set_status("A server sync is already running.")
            return
        self._syncing = True
        self._set_status("Syncing with Exchange. A sign-in browser may open.")
        self._sync_in_background()

    @work(thread=True, exclusive=True)
    def _sync_in_background(self) -> None:
        log_messages = []
        try:
            client_id, authority, scopes = load_config()
            access_token = acquire_access_token(
                client_id, authority, scopes, log=log_messages.append
            )
            graph = GraphClient(access_token)
            worker_store = EmailStore(DATABASE_PATH)
            try:
                summary = sync_emails(graph, worker_store, log=log_messages.append)
            finally:
                worker_store.close()
        except Exception as error:
            self.call_from_thread(
                self._sync_finished, str(error), None, log_messages
            )
            return
        self.call_from_thread(self._sync_finished, None, summary, log_messages)

    def _sync_finished(self, error, summary, log_messages) -> None:
        self._syncing = False
        self.refresh_emails()
        if error:
            self._set_status(f"Server sync failed: {error}")
            return
        succeeded, failed = summary
        status = f"Sync finished: {succeeded} processed, {failed} failed."
        if failed:
            failures = [
                message for message in log_messages if message.startswith("Failed to process")
            ]
            if failures:
                status += "\n" + failures[0]
        self._set_status(status)

    def action_generate_pdf(self) -> None:
        if self._syncing:
            self._set_status("Wait for the server sync to finish before converting.")
            return
        if not self.selected_id:
            self._set_status("Select an email first.")
            return
        try:
            output_path = convert_email_to_pdf(self.store, self.selected_id)
        except Exception as error:
            self._set_status(f"PDF conversion failed: {error}")
            return
        self.refresh_emails(select_id=self.selected_id)
        self._set_status(f"PDF created: {output_path}")

    def action_delete_local(self) -> None:
        if self._syncing:
            self._set_status("Wait for the server sync to finish before deleting.")
            return
        if not self.selected_id:
            self._set_status("Select an email first.")
            return
        subject = self.emails[self.selected_id].get("subject") or "(No Subject)"
        self.push_screen(
            ConfirmDeleteScreen(subject),
            self._delete_confirmed,
        )

    def action_quit(self) -> None:
        if self._syncing:
            self._set_status("Wait for the server sync to finish before quitting.")
            return
        self.exit()

    def _delete_confirmed(self, confirmed) -> None:
        if not confirmed or not self.selected_id:
            return
        graph_id = self.selected_id
        try:
            cleanup_errors = delete_email_locally(self.store, graph_id)
        except Exception as error:
            self._set_status(f"Delete failed: {error}")
            return
        self.selected_id = None
        self.refresh_emails()
        if cleanup_errors:
            self._set_status(
                "Email removed from the database; some local files could not be "
                "removed: "
                + "; ".join(cleanup_errors)
            )
        else:
            self._set_status(f"Deleted local copy of {graph_id}.")


def _format_person(person):
    address = (person or {}).get("emailAddress") or {}
    name = address.get("name")
    email_address = address.get("address")
    if name and email_address:
        return f"{name} <{email_address}>"
    return name or email_address or "—"


def _format_recipients(people):
    return "; ".join(_format_person(person) for person in people or [])
