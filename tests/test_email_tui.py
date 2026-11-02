import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from email_store import EmailStore
import email_tui
from email_tui import EmailTui
import pull_emails
from textual.widgets import DataTable


def make_message():
    return {
        "id": "tui-message-id",
        "subject": "TUI example",
        "from": {
            "emailAddress": {"name": "Sender", "address": "sender@example.com"}
        },
        "toRecipients": [],
        "ccRecipients": [],
        "bccRecipients": [],
        "replyTo": [],
        "body": {"contentType": "text", "content": "Example body"},
        "receivedDateTime": "2026-10-17T10:00:00Z",
        "sentDateTime": "2026-10-17T09:59:00Z",
        "flag": {"flagStatus": "complete"},
    }


class FakeGraph:
    def __init__(self, message):
        self.message = message
        self.flags = []

    def list_not_flagged_messages(self):
        yield self.message

    def list_attachments(self, _message_id):
        return iter(())

    def set_flag(self, message_id, flag_status):
        self.flags.append((message_id, flag_status))


class EmailTuiTests(unittest.IsolatedAsyncioTestCase):
    async def test_pdf_action_and_confirmed_local_delete(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            attachment_root = root / "attachments"
            pdf_root = root / "pdfs"
            store = EmailStore(root / "emails.sqlite3")
            store.upsert_email(make_message())
            attachment_directory = attachment_root / "tui-message"
            attachment_directory.mkdir(parents=True)
            attachment_path = attachment_directory / "note.txt"
            attachment_path.write_text("Attached note", encoding="utf-8")
            store.upsert_attachment(
                "tui-message-id",
                {
                    "id": "tui-attachment-id",
                    "name": "note.txt",
                    "contentType": "text/plain",
                    "size": attachment_path.stat().st_size,
                },
                "fileAttachment",
                attachment_index=0,
            )
            store.set_attachment_downloaded(
                "tui-message-id", "tui-attachment-id", attachment_path, "sha256"
            )

            with (
                patch.object(pull_emails, "ATTACHMENT_ROOT", attachment_root),
                patch.object(pull_emails, "PDF_ROOT", pdf_root),
            ):
                app = EmailTui(store)
                async with app.run_test() as pilot:
                    await pilot.pause()
                    self.assertEqual(app.query_one(DataTable).row_count, 1)
                    self.assertEqual(app.selected_id, "tui-message-id")

                    await pilot.press("p")
                    await pilot.pause()
                    generated_email = store.get_email("tui-message-id")
                    pdf_path = Path(generated_email["pdf_path"])
                    self.assertTrue(pdf_path.is_file())

                    app.action_delete_local()
                    await pilot.pause()
                    await pilot.click("#confirm-delete")
                    await pilot.pause()
                    self.assertIsNone(store.get_email("tui-message-id"))
                    self.assertFalse(pdf_path.exists())
                    self.assertFalse(attachment_path.exists())

            store.close()

    async def test_server_sync_action_pulls_and_refreshes_stored_emails(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            database_path = Path(temporary_directory) / "emails.sqlite3"
            store = EmailStore(database_path)
            message = make_message()
            message["id"] = "server-pulled-id"
            message["flag"] = {"flagStatus": "notFlagged"}
            graph = FakeGraph(message)

            with (
                patch.object(email_tui, "DATABASE_PATH", database_path),
                patch.object(
                    email_tui,
                    "load_config",
                    return_value=("client-id", "authority", ["Mail.ReadWrite"]),
                ),
                patch.object(
                    email_tui, "acquire_access_token", return_value="access-token"
                ),
                patch.object(email_tui, "GraphClient", return_value=graph),
            ):
                app = EmailTui(store)
                async with app.run_test() as pilot:
                    await pilot.pause()
                    await pilot.press("s")
                    for _ in range(40):
                        await pilot.pause(0.05)
                        if not app._syncing:
                            break

                    self.assertFalse(app._syncing)
                    self.assertEqual(app.query_one(DataTable).row_count, 1)
                    self.assertIsNotNone(store.get_email("server-pulled-id"))
                    self.assertEqual(
                        graph.flags, [("server-pulled-id", "complete")]
                    )

            store.close()

    async def test_pdf_action_uses_the_highlighted_row(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            store = EmailStore(root / "emails.sqlite3")
            self.addCleanup(store.close)
            first = make_message()
            first["id"] = "older-email"
            first["receivedDateTime"] = "2026-10-16T10:00:00Z"
            second = make_message()
            second["id"] = "newer-email"
            second["receivedDateTime"] = "2026-10-17T10:00:00Z"
            store.upsert_email(first)
            store.upsert_email(second)

            with patch.object(pull_emails, "PDF_ROOT", root / "pdfs"):
                app = EmailTui(store)
                async with app.run_test() as pilot:
                    await pilot.pause()
                    self.assertEqual(app.selected_id, "newer-email")

                    await pilot.press("down")
                    await pilot.pause()
                    self.assertEqual(app.selected_id, "older-email")

                    await pilot.press("p")
                    await pilot.pause()
                    self.assertIsNotNone(store.get_email("older-email")["pdf_path"])
                    self.assertIsNone(store.get_email("newer-email")["pdf_path"])

            store.close()


if __name__ == "__main__":
    unittest.main()
