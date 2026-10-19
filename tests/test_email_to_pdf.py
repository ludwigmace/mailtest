import tempfile
import unittest
from email.message import EmailMessage
from email.policy import SMTP
from pathlib import Path

from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

from email_to_pdf import _format_datetime, create_email_pdf


class EmailToPdfTests(unittest.TestCase):
    def test_renders_cid_image_and_appends_converted_and_existing_pdfs(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            image_path = root / "inline.png"
            Image.new("RGB", (80, 40), "red").save(image_path)

            text_path = root / "notes.txt"
            text_path.write_text("Appended text attachment", encoding="utf-8")

            existing_pdf = root / "appendix.pdf"
            appendix = canvas.Canvas(str(existing_pdf))
            appendix.drawString(72, 720, "Existing PDF attachment")
            appendix.save()

            email = {
                "subject": "Rendered message",
                "sender": {
                    "emailAddress": {
                        "name": "Example Sender",
                        "address": "sender@example.com",
                    }
                },
                "to_recipients": [
                    {
                        "emailAddress": {
                            "name": "Example Recipient",
                            "address": "recipient@example.com",
                        }
                    }
                ],
                "cc_recipients": [],
                "bcc_recipients": [],
                "sent_datetime": "2026-10-17T10:00:00Z",
                "body_content_type": "html",
                "body_content": (
                    "<p>Message body text</p>"
                    '<img src="cid:brand@example.com" alt="embedded logo">'
                ),
                "attachments": [
                    {
                        "name": "inline.png",
                        "local_path": str(image_path),
                        "content_id": "<brand@example.com>",
                        "download_status": "downloaded",
                        "is_inline": True,
                    },
                    {
                        "name": "notes.txt",
                        "local_path": str(text_path),
                        "content_type": "text/plain",
                        "download_status": "downloaded",
                        "is_inline": False,
                    },
                    {
                        "name": "appendix.pdf",
                        "local_path": str(existing_pdf),
                        "content_type": "application/pdf",
                        "download_status": "downloaded",
                        "is_inline": False,
                    },
                ],
            }
            output_path = root / "rendered-email.pdf"

            create_email_pdf(email, output_path)

            reader = PdfReader(str(output_path))
            self.assertEqual(len(reader.pages), 3)
            first_page_text = reader.pages[0].extract_text()
            header_positions = [
                first_page_text.index(label)
                for label in ("From:", "Sent:", "To:", "Subject:")
            ]
            self.assertEqual(header_positions, sorted(header_positions))
            self.assertIn("Subject: Rendered message", first_page_text)
            self.assertIn("Example Sender", first_page_text)
            self.assertIn(
                "To: Example Recipient <recipient@example.com>", first_page_text
            )
            formatted_sent = _format_datetime("2026-10-17T10:00:00Z")
            self.assertRegex(
                formatted_sent,
                r"^[A-Za-z]+, [A-Za-z]+ \d{2}, 2026 "
                r"\d{1,2}:\d{2} (AM|PM) [A-Z]{2,8}$",
            )
            self.assertIn(
                f"Sent: {formatted_sent}", first_page_text
            )
            self.assertNotIn("2026-10-17T10:00:00Z", first_page_text)
            self.assertIn("Message body text", first_page_text)
            image_resources = reader.pages[0]["/Resources"].get("/XObject")
            self.assertTrue(image_resources)
            self.assertIn("Appended text attachment", reader.pages[1].extract_text())
            self.assertIn("Existing PDF attachment", reader.pages[2].extract_text())

    def test_appends_attached_email_and_its_attachments(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            attached_email = EmailMessage()
            attached_email["Subject"] = "Attached message"
            attached_email["From"] = "Nested Sender <nested@example.com>"
            attached_email["To"] = "Nested Recipient <to@example.com>"
            attached_email.set_content("Nested email body")
            attached_email.add_attachment(
                "Nested attachment text",
                subtype="plain",
                filename="nested.txt",
            )
            eml_path = root / "attached-message.eml"
            eml_path.write_bytes(attached_email.as_bytes(policy=SMTP))

            email = {
                "subject": "Outer message",
                "body_content_type": "text",
                "body_content": "Outer email body",
                "attachments": [
                    {
                        "name": "attached-message.eml",
                        "local_path": str(eml_path),
                        "content_type": "message/rfc822",
                        "download_status": "downloaded",
                        "is_inline": False,
                    }
                ],
            }
            output_path = root / "outer-email.pdf"

            create_email_pdf(email, output_path)

            reader = PdfReader(str(output_path))
            self.assertEqual(len(reader.pages), 3)
            self.assertIn("Outer email body", reader.pages[0].extract_text())
            self.assertIn("Attached message", reader.pages[1].extract_text())
            self.assertIn("Nested email body", reader.pages[1].extract_text())
            self.assertIn("Nested attachment text", reader.pages[2].extract_text())

    def test_unsupported_attachment_does_not_replace_existing_pdf(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            attachment_path = root / "archive.bin"
            attachment_path.write_bytes(b"not a supported document")
            output_path = root / "existing.pdf"
            output_path.write_bytes(b"previous PDF")

            email = {
                "subject": "Message",
                "body_content_type": "text",
                "body_content": "Body",
                "attachments": [
                    {
                        "name": "archive.bin",
                        "local_path": str(attachment_path),
                        "download_status": "downloaded",
                        "is_inline": False,
                    }
                ],
            }
            with self.assertRaisesRegex(RuntimeError, "Unsupported attachment type"):
                create_email_pdf(email, output_path)
            self.assertEqual(output_path.read_bytes(), b"previous PDF")


if __name__ == "__main__":
    unittest.main()
