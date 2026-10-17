import hashlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pull_emails
from email_store import EmailStore


class FakeResponse:
    def __init__(self, content):
        self.content = content
        self.closed = False

    def iter_content(self, chunk_size):
        for offset in range(0, len(self.content), chunk_size):
            yield self.content[offset : offset + chunk_size]

    def close(self):
        self.closed = True


class FakeGraph:
    def __init__(self, message):
        self.message = message
        self.flags = []
        self.download_count = 0

    def list_not_flagged_messages(self):
        yield self.message

    def list_attachments(self, message_id):
        yield {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "id": "attachment-1",
            "name": "report.txt",
            "contentType": "text/plain",
            "size": 11,
            "isInline": False,
        }

    def download_attachment(self, message_id, attachment_id):
        self.download_count += 1
        return FakeResponse(b"hello world")

    def set_flag(self, message_id, flag_status):
        self.flags.append((message_id, flag_status))


def make_message(subject="Quarterly report"):
    return {
        "id": "immutable-message-id",
        "internetMessageId": "<message@example.com>",
        "subject": subject,
        "from": {"emailAddress": {"name": "Sender", "address": "sender@example.com"}},
        "toRecipients": [
            {"emailAddress": {"name": "Recipient", "address": "to@example.com"}}
        ],
        "ccRecipients": [],
        "bccRecipients": [],
        "replyTo": [],
        "body": {"contentType": "text", "content": "Message body"},
        "receivedDateTime": "2026-10-17T10:00:00Z",
        "sentDateTime": "2026-10-17T09:59:00Z",
        "isRead": False,
        "importance": "normal",
        "conversationId": "conversation-1",
        "hasAttachments": True,
        "flag": {"flagStatus": "notFlagged"},
    }


class EmailSyncTests(unittest.TestCase):
    def test_upsert_download_process_and_resync_updates_existing_email(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            store = EmailStore(root / "emails.sqlite3")
            graph = FakeGraph(make_message())

            with patch.object(pull_emails, "ATTACHMENT_ROOT", root / "attachments"):
                pull_emails.sync_emails(graph, store)

                email = store.get_email("immutable-message-id")
                self.assertEqual(email["processing_status"], "processed")
                self.assertEqual(email["flag_status"], "complete")
                self.assertEqual(email["body_content"], "Message body")
                self.assertEqual(email["to_recipients"][0]["emailAddress"]["address"], "to@example.com")
                self.assertEqual(len(email["attachments"]), 1)

                attachment = email["attachments"][0]
                attachment_path = Path(attachment["local_path"])
                self.assertEqual(attachment["download_status"], "downloaded")
                self.assertEqual(attachment_path.read_bytes(), b"hello world")
                self.assertEqual(
                    attachment["sha256"], hashlib.sha256(b"hello world").hexdigest()
                )
                self.assertEqual(graph.flags, [("immutable-message-id", "complete")])

                graph.message = make_message(subject="Updated subject")
                pull_emails.sync_emails(graph, store)

            email_count = store.connection.execute(
                "SELECT COUNT(*) FROM emails"
            ).fetchone()[0]
            attachment_count = store.connection.execute(
                "SELECT COUNT(*) FROM attachments"
            ).fetchone()[0]
            updated_email = store.get_email("immutable-message-id")

            self.assertEqual(email_count, 1)
            self.assertEqual(attachment_count, 1)
            self.assertEqual(updated_email["subject"], "Updated subject")
            self.assertEqual(graph.download_count, 1)
            self.assertEqual(len(graph.flags), 2)
            store.close()


if __name__ == "__main__":
    unittest.main()
