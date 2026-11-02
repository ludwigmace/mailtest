"""SQLite persistence for pulled messages and their attachments."""

import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path


def _utc_now():
    return datetime.now(timezone.utc).isoformat()


def _json(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


class EmailStore:
    def __init__(self, database_path):
        database_path = Path(database_path)
        if str(database_path) != ":memory:":
            database_path.parent.mkdir(parents=True, exist_ok=True)
        self.connection = sqlite3.connect(str(database_path))
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")
        self.connection.execute("PRAGMA journal_mode = WAL")
        self._initialize_schema()

    def _initialize_schema(self):
        self.connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS emails (
                graph_id TEXT PRIMARY KEY,
                internet_message_id TEXT,
                subject TEXT,
                sender_json TEXT NOT NULL,
                to_recipients_json TEXT NOT NULL,
                cc_recipients_json TEXT NOT NULL,
                bcc_recipients_json TEXT NOT NULL,
                reply_to_json TEXT NOT NULL,
                body_content_type TEXT,
                body_content TEXT,
                received_datetime TEXT,
                sent_datetime TEXT,
                is_read INTEGER,
                importance TEXT,
                conversation_id TEXT,
                has_attachments INTEGER NOT NULL DEFAULT 0,
                flag_status TEXT,
                processing_status TEXT NOT NULL DEFAULT 'pending',
                downloaded_at TEXT NOT NULL,
                processed_at TEXT,
                pdf_path TEXT,
                last_error TEXT,
                raw_message_json TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS attachments (
                graph_id TEXT NOT NULL,
                attachment_id TEXT NOT NULL,
                name TEXT NOT NULL,
                content_type TEXT,
                size INTEGER,
                is_inline INTEGER NOT NULL DEFAULT 0,
                content_id TEXT,
                attachment_type TEXT,
                attachment_index INTEGER NOT NULL DEFAULT 0,
                local_path TEXT,
                sha256 TEXT,
                download_status TEXT NOT NULL DEFAULT 'pending',
                download_error TEXT,
                downloaded_at TEXT,
                PRIMARY KEY (graph_id, attachment_id),
                FOREIGN KEY (graph_id) REFERENCES emails(graph_id) ON DELETE CASCADE
            );

            CREATE INDEX IF NOT EXISTS idx_emails_flag_status
                ON emails(flag_status);
            CREATE INDEX IF NOT EXISTS idx_attachments_graph_id
                ON attachments(graph_id);
            """
        )
        email_columns = {
            row["name"]
            for row in self.connection.execute("PRAGMA table_info(emails)")
        }
        if "pdf_path" not in email_columns:
            self.connection.execute("ALTER TABLE emails ADD COLUMN pdf_path TEXT")
        attachment_columns = {
            row["name"]
            for row in self.connection.execute("PRAGMA table_info(attachments)")
        }
        if "attachment_index" not in attachment_columns:
            self.connection.execute(
                "ALTER TABLE attachments ADD COLUMN attachment_index "
                "INTEGER NOT NULL DEFAULT 0"
            )
        self.connection.commit()

    def upsert_email(self, message):
        graph_id = message.get("id")
        if not graph_id:
            raise ValueError("Graph message is missing its id")

        body = message.get("body") or {}
        flag = message.get("flag") or {}
        now = _utc_now()
        values = (
            graph_id,
            message.get("internetMessageId"),
            message.get("subject"),
            _json(message.get("from") or {}),
            _json(message.get("toRecipients") or []),
            _json(message.get("ccRecipients") or []),
            _json(message.get("bccRecipients") or []),
            _json(message.get("replyTo") or []),
            body.get("contentType"),
            body.get("content"),
            message.get("receivedDateTime"),
            message.get("sentDateTime"),
            None if message.get("isRead") is None else int(message["isRead"]),
            message.get("importance"),
            message.get("conversationId"),
            int(bool(message.get("hasAttachments"))),
            flag.get("flagStatus"),
            now,
            _json(message),
            now,
            now,
        )
        with self.connection:
            self.connection.execute(
                """
                INSERT INTO emails (
                    graph_id, internet_message_id, subject, sender_json,
                    to_recipients_json, cc_recipients_json, bcc_recipients_json,
                    reply_to_json, body_content_type, body_content,
                    received_datetime, sent_datetime, is_read, importance,
                    conversation_id, has_attachments, flag_status, downloaded_at,
                    raw_message_json, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(graph_id) DO UPDATE SET
                    internet_message_id = excluded.internet_message_id,
                    subject = excluded.subject,
                    sender_json = excluded.sender_json,
                    to_recipients_json = excluded.to_recipients_json,
                    cc_recipients_json = excluded.cc_recipients_json,
                    bcc_recipients_json = excluded.bcc_recipients_json,
                    reply_to_json = excluded.reply_to_json,
                    body_content_type = excluded.body_content_type,
                    body_content = excluded.body_content,
                    received_datetime = excluded.received_datetime,
                    sent_datetime = excluded.sent_datetime,
                    is_read = excluded.is_read,
                    importance = excluded.importance,
                    conversation_id = excluded.conversation_id,
                    has_attachments = excluded.has_attachments,
                    flag_status = excluded.flag_status,
                    processing_status = 'pending',
                    downloaded_at = excluded.downloaded_at,
                    processed_at = NULL,
                    pdf_path = CASE
                        WHEN emails.subject IS NOT excluded.subject
                          OR emails.sender_json IS NOT excluded.sender_json
                          OR emails.to_recipients_json IS NOT excluded.to_recipients_json
                          OR emails.cc_recipients_json IS NOT excluded.cc_recipients_json
                          OR emails.bcc_recipients_json IS NOT excluded.bcc_recipients_json
                          OR emails.reply_to_json IS NOT excluded.reply_to_json
                          OR emails.body_content_type IS NOT excluded.body_content_type
                          OR emails.body_content IS NOT excluded.body_content
                          OR emails.received_datetime IS NOT excluded.received_datetime
                          OR emails.sent_datetime IS NOT excluded.sent_datetime
                        THEN NULL
                        ELSE emails.pdf_path
                    END,
                    last_error = NULL,
                    raw_message_json = excluded.raw_message_json,
                    updated_at = excluded.updated_at
                """,
                values,
            )

    def get_attachment(self, graph_id, attachment_id):
        row = self.connection.execute(
            "SELECT * FROM attachments WHERE graph_id = ? AND attachment_id = ?",
            (graph_id, attachment_id),
        ).fetchone()
        return dict(row) if row else None

    def upsert_attachment(
        self, graph_id, attachment, attachment_type, attachment_index=0
    ):
        attachment_id = attachment.get("id")
        if not attachment_id:
            raise ValueError("Graph attachment is missing its id")
        metadata = (
            attachment.get("name") or "attachment",
            attachment.get("contentType"),
            attachment.get("size"),
            int(bool(attachment.get("isInline"))),
            attachment.get("contentId"),
            attachment_type,
            attachment_index,
        )
        with self.connection:
            existing = self.connection.execute(
                """SELECT name, content_type, size, is_inline, content_id,
                          attachment_type, attachment_index
                   FROM attachments WHERE graph_id = ? AND attachment_id = ?""",
                (graph_id, attachment_id),
            ).fetchone()
            if not existing or tuple(existing) != metadata:
                self.connection.execute(
                    "UPDATE emails SET pdf_path = NULL WHERE graph_id = ?",
                    (graph_id,),
                )
            self.connection.execute(
                """
                INSERT INTO attachments (
                    graph_id, attachment_id, name, content_type, size,
                    is_inline, content_id, attachment_type, attachment_index
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(graph_id, attachment_id) DO UPDATE SET
                    name = excluded.name,
                    content_type = excluded.content_type,
                    size = excluded.size,
                    is_inline = excluded.is_inline,
                    content_id = excluded.content_id,
                    attachment_type = excluded.attachment_type,
                    attachment_index = excluded.attachment_index
                """,
                (
                    graph_id,
                    attachment_id,
                    *metadata,
                ),
            )

    def set_attachment_downloaded(self, graph_id, attachment_id, path, sha256):
        with self.connection:
            self.connection.execute(
                """
                UPDATE attachments
                SET local_path = ?, sha256 = ?, download_status = 'downloaded',
                    download_error = NULL, downloaded_at = ?
                WHERE graph_id = ? AND attachment_id = ?
                """,
                (str(path), sha256, _utc_now(), graph_id, attachment_id),
            )

    def set_attachment_status(self, graph_id, attachment_id, status, error=None):
        with self.connection:
            self.connection.execute(
                """
                UPDATE attachments
                SET download_status = ?, download_error = ?
                WHERE graph_id = ? AND attachment_id = ?
                """,
                (status, error, graph_id, attachment_id),
            )

    def get_email(self, graph_id):
        row = self.connection.execute(
            "SELECT * FROM emails WHERE graph_id = ?", (graph_id,)
        ).fetchone()
        if not row:
            return None

        email = dict(row)
        for column, key in (
            ("sender_json", "sender"),
            ("to_recipients_json", "to_recipients"),
            ("cc_recipients_json", "cc_recipients"),
            ("bcc_recipients_json", "bcc_recipients"),
            ("reply_to_json", "reply_to"),
            ("raw_message_json", "raw_message"),
        ):
            email[key] = json.loads(email.pop(column))

        email["attachments"] = [
            dict(attachment)
            for attachment in self.connection.execute(
                "SELECT * FROM attachments WHERE graph_id = ? "
                "ORDER BY attachment_index, name",
                (graph_id,),
            )
        ]
        return email

    def list_recent_emails(self, limit=20):
        rows = self.connection.execute(
            """
            SELECT graph_id, subject, sender_json, received_datetime,
                   processing_status, flag_status, pdf_path,
                   (SELECT COUNT(*) FROM attachments
                    WHERE attachments.graph_id = emails.graph_id) AS attachment_count
            FROM emails
            ORDER BY received_datetime DESC, downloaded_at DESC
            LIMIT ?
            """,
            (limit,),
        )
        emails = []
        for row in rows:
            email = dict(row)
            email["sender"] = json.loads(email.pop("sender_json"))
            emails.append(email)
        return emails

    def delete_email(self, graph_id):
        email = self.get_email(graph_id)
        if email is None:
            return None
        with self.connection:
            self.connection.execute(
                "DELETE FROM emails WHERE graph_id = ?", (graph_id,)
            )
        return email

    def set_processing_status(self, graph_id, status, flag_status=None, error=None):
        processed_at = _utc_now() if status == "processed" else None
        with self.connection:
            self.connection.execute(
                """
                UPDATE emails
                SET processing_status = ?,
                    flag_status = COALESCE(?, flag_status),
                    processed_at = ?, last_error = ?, updated_at = ?
                WHERE graph_id = ?
                """,
                (status, flag_status, processed_at, error, _utc_now(), graph_id),
            )

    def set_pdf_path(self, graph_id, pdf_path):
        with self.connection:
            self.connection.execute(
                "UPDATE emails SET pdf_path = ?, updated_at = ? WHERE graph_id = ?",
                (str(pdf_path), _utc_now(), graph_id),
            )

    def close(self):
        self.connection.close()
