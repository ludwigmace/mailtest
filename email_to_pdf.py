"""Render stored Outlook messages and attachments into a single PDF."""

import html
import mimetypes
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime, timezone
from email import policy as email_policy
from email.parser import BytesParser
from pathlib import Path
from email.utils import getaddresses, parsedate_to_datetime
from urllib.parse import unquote, urlparse
from urllib.request import url2pathname


PAGE_CSS = """
@page { size: letter; margin: 0.65in 0.7in; }
body { font-family: Arial, sans-serif; font-size: 10pt; color: #242424; }
.outlook-header { margin-bottom: 16px; }
.outlook-header-line { padding: 2px 0; }
.outlook-label { font-weight: bold; }
.outlook-subject { border-bottom: 1px solid #000000; padding: 2px 0 8px 0; margin-bottom: 12px; }
.outlook-body { overflow-wrap: break-word; }
img { max-width: 100%; }
pre { white-space: pre-wrap; overflow-wrap: break-word; font-family: Arial, sans-serif; }
"""

OFFICE_EXTENSIONS = {
    ".doc", ".docm", ".docx", ".dot", ".dotx", ".odt", ".ott", ".rtf",
    ".xls", ".xlsb", ".xlsm", ".xlsx", ".xlt", ".xltx", ".ods", ".ots",
    ".ppt", ".pptm", ".pptx", ".pps", ".ppsx", ".odp", ".otp",
}
TEXT_EXTENSIONS = {".txt", ".csv", ".tsv", ".log", ".md", ".xml", ".json"}
HTML_EXTENSIONS = {".html", ".htm", ".xhtml"}
MAX_NESTED_EMAIL_DEPTH = 5


def _display_address(person):
    address = (person or {}).get("emailAddress") or {}
    name = address.get("name", "").strip()
    email_address = address.get("address", "").strip()
    if name and email_address:
        return f"{html.escape(name)} &lt;{html.escape(email_address)}&gt;"
    return html.escape(name or email_address)


def _display_recipients(people):
    return "; ".join(_display_address(person) for person in people or [])


def _timezone_abbreviation(local_time):
    name = local_time.tzname() or "UTC"
    normalized = name.casefold()
    if normalized in ("utc", "coordinated universal time"):
        return "UTC"
    if normalized in ("gmt standard time", "gmt daylight time"):
        return "BST" if local_time.dst() else "GMT"
    return "".join(word[0] for word in re.findall(r"[A-Za-z0-9]+", name)).upper()


def _format_datetime(value):
    if not value:
        return ""
    timestamp = str(value).strip()
    if timestamp.upper().endswith(" UTC"):
        timestamp = timestamp[:-4] + "+00:00"
    elif timestamp.endswith("Z"):
        timestamp = timestamp[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(timestamp)
    except ValueError:
        return html.escape(str(value))
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    local_time = parsed.astimezone()
    formatted = local_time.strftime("%A, %B %d, %Y %I:%M %p")
    formatted = re.sub(r"(?<= )0(?=\d:)", "", formatted)
    return html.escape(f"{formatted} {_timezone_abbreviation(local_time)}")


def _clean_html(content):
    content = re.sub(r"(?is)<(script|iframe|object|embed)\b[^>]*>.*?</\1\s*>", "", content)
    content = re.sub(r"(?is)<(script|iframe|object|embed)\b[^>]*/\s*>", "", content)
    content = re.sub(r"(?i)\s+on[a-z]+\s*=\s*(?:\"[^\"]*\"|'[^']*'|[^\s>]+)", "", content)

    styles = "\n".join(re.findall(r"(?is)<style\b[^>]*>(.*?)</style\s*>", content))
    body_match = re.search(r"(?is)<body\b[^>]*>(.*?)</body\s*>", content)
    if body_match:
        body = body_match.group(1)
    else:
        body = re.sub(r"(?is)<head\b[^>]*>.*?</head\s*>", "", content)
        body = re.sub(r"(?is)</?html\b[^>]*>|<!doctype[^>]*>", "", body)
    body = re.sub(r"(?is)<style\b[^>]*>.*?</style\s*>", "", body)
    return styles, body


def _cid_key(value):
    value = unquote(str(value or "")).strip()
    if value.lower().startswith("cid:"):
        value = value[4:]
    return value.strip("<>").strip().lower()


def _attachment_lookup(email):
    by_cid = {}
    for attachment in email.get("attachments", []):
        content_id = attachment.get("content_id")
        local_path = attachment.get("local_path")
        if content_id and local_path and Path(local_path).is_file():
            by_cid[_cid_key(content_id)] = Path(local_path).resolve()
    return by_cid


def _referenced_content_ids(email):
    if (email.get("body_content_type") or "").lower() != "html":
        return set()
    body = email.get("body_content") or ""
    return {
        _cid_key(content_id)
        for content_id in re.findall(r"(?i)cid:([^\"'\s)>]+)", body)
    }


def _link_callback(base_dir, by_cid):
    base_dir = Path(base_dir).resolve()

    def resolve(uri, _relative_to=None):
        if not uri:
            return None
        if uri.lower().startswith("cid:"):
            path = by_cid.get(_cid_key(uri))
            return str(path) if path else None
        parsed = urlparse(uri)
        if parsed.scheme and parsed.scheme.lower() != "file":
            return None
        if parsed.scheme.lower() == "file":
            candidate = Path(url2pathname(parsed.path))
        else:
            candidate = base_dir / unquote(parsed.path)
        try:
            candidate = candidate.resolve(strict=True)
            candidate.relative_to(base_dir)
        except (OSError, ValueError):
            return None
        return str(candidate) if candidate.is_file() else None

    return resolve


def _html_document(email):
    subject = html.escape(email.get("subject") or "(No Subject)")
    fields = (
        ("From", _display_address(email.get("sender"))),
        ("Sent", _format_datetime(email.get("sent_datetime"))),
        ("To", _display_recipients(email.get("to_recipients"))),
    )
    header_lines = "".join(
        f'<div class="outlook-header-line"><span class="outlook-label">'
        f"{label}:</span>{f' {value}' if value else ''}</div>"
        for label, value in fields
    )
    body_content = email.get("body_content") or ""
    if (email.get("body_content_type") or "text").lower() == "html":
        message_styles, message_body = _clean_html(body_content)
    else:
        message_styles = ""
        message_body = f"<pre>{html.escape(body_content)}</pre>"

    document = f"""<!doctype html>
<html><head><meta charset="utf-8"><style>{PAGE_CSS}\n{message_styles}</style></head>
<body><div class="outlook-header">{header_lines}
<div class="outlook-subject"><span class="outlook-label">Subject:</span> {subject}</div></div>
<div class="outlook-body">{message_body}</div></body></html>"""
    return document


def _render_html_pdf(content, target, base_dir, by_cid=None):
    from xhtml2pdf import pisa
    from xhtml2pdf.files import ResourceAccessPolicy

    target = Path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    by_cid = by_cid or {}
    resource_roots = tuple({path.parent for path in by_cid.values()})
    policy = ResourceAccessPolicy.server(
        base_dir=base_dir,
        extra_roots=resource_roots,
        allow_remote=False,
    )
    with target.open("wb") as output:
        result = pisa.CreatePDF(
            content,
            dest=output,
            path=str(Path(base_dir) / "document.html"),
            encoding="utf-8",
            link_callback=_link_callback(base_dir, by_cid),
            resource_policy=policy,
        )
    if result.err:
        target.unlink(missing_ok=True)
        raise RuntimeError(f"HTML-to-PDF conversion failed for {target.name}")


def _render_email_pdf(email, target):
    target = Path(target)
    base_dir = target.parent
    document = _html_document(email)
    _render_html_pdf(
        document,
        target,
        base_dir,
        by_cid=_attachment_lookup(email),
    )


def _render_text_attachment(source, target):
    content = Path(source).read_text(encoding="utf-8", errors="replace")
    title = html.escape(Path(source).name)
    document = (
        "<!doctype html><html><head><meta charset='utf-8'></head><body>"
        f"<h2>{title}</h2><pre>{html.escape(content)}</pre></body></html>"
    )
    _render_html_pdf(document, target, Path(source).parent)


def _render_html_attachment(source, target):
    content = Path(source).read_text(encoding="utf-8", errors="replace")
    styles, body = _clean_html(content)
    document = (
        "<!doctype html><html><head><meta charset='utf-8'><style>"
        f"{PAGE_CSS}\n{styles}</style></head><body>{body}</body></html>"
    )
    _render_html_pdf(document, target, Path(source).parent)


def _render_image_attachment(source, target):
    from PIL import Image
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas

    page_width, page_height = letter
    margin = 36
    output = canvas.Canvas(str(target), pagesize=letter)
    with Image.open(source) as image:
        frame_count = getattr(image, "n_frames", 1)
        for frame_number in range(frame_count):
            image.seek(frame_number)
            frame = image.convert("RGBA")
            background = Image.new("RGBA", frame.size, "white")
            background.alpha_composite(frame)
            image_reader = ImageReader(background.convert("RGB"))
            image_width, image_height = image.size
            scale = min(
                (page_width - margin * 2) / image_width,
                (page_height - margin * 2) / image_height,
            )
            drawn_width, drawn_height = image_width * scale, image_height * scale
            output.drawImage(
                image_reader,
                (page_width - drawn_width) / 2,
                (page_height - drawn_height) / 2,
                width=drawn_width,
                height=drawn_height,
                preserveAspectRatio=True,
            )
            output.showPage()
    output.save()


def _find_libreoffice():
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable:
        return executable
    if os.name == "nt":
        candidates = (
            Path(os.environ.get("PROGRAMFILES", "C:/Program Files"))
            / "LibreOffice/program/soffice.exe",
            Path(os.environ.get("PROGRAMFILES(X86)", "C:/Program Files (x86)"))
            / "LibreOffice/program/soffice.exe",
        )
        for candidate in candidates:
            if candidate.is_file():
                return str(candidate)
    return None


def _render_office_attachment(source, target, libreoffice):
    target = Path(target)
    profile = target.parent / "libreoffice-profile"
    command = [
        libreoffice,
        "--headless",
        f"-env:UserInstallation={profile.resolve().as_uri()}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(target.parent),
        str(Path(source).resolve()),
    ]
    result = subprocess.run(command, capture_output=True, text=True, timeout=180)
    converted = target.parent / f"{Path(source).stem}.pdf"
    if result.returncode or not converted.is_file():
        detail = (result.stderr or result.stdout).strip()
        raise RuntimeError(
            f"LibreOffice could not convert {Path(source).name}"
            + (f": {detail}" if detail else "")
        )
    if converted != target:
        converted.replace(target)


def _email_people(headers):
    people = []
    for name, address in getaddresses(headers):
        people.append(
            {"emailAddress": {"name": name, "address": address}}
        )
    return people


def _email_from_eml(source, attachment_directory):
    message = BytesParser(policy=email_policy.default).parsebytes(
        Path(source).read_bytes()
    )
    body_part = message.get_body(preferencelist=("html", "plain"))
    if body_part is None:
        body_content_type = "text"
        body_content = ""
    else:
        body_content_type = "html" if body_part.get_content_subtype() == "html" else "text"
        try:
            body_content = body_part.get_content()
        except (LookupError, UnicodeDecodeError):
            body_content = body_part.get_payload(decode=True).decode(
                "utf-8", errors="replace"
            )

    attachment_directory.mkdir(parents=True, exist_ok=True)
    attachments = []
    for index, part in enumerate(message.iter_attachments(), start=1):
        name = part.get_filename() or f"attachment-{index}"
        name = Path(name.replace("\\", "/")).name
        content_type = part.get_content_type()
        if not Path(name).suffix:
            name += mimetypes.guess_extension(content_type) or ".bin"
        local_path = attachment_directory / f"{index:04d}_{name}"
        payload = part.get_payload(decode=True)
        if payload is None:
            payload = part.as_bytes(policy=email_policy.SMTP)
        local_path.write_bytes(payload)
        attachments.append(
            {
                "name": name,
                "content_type": content_type,
                "content_id": part.get("Content-ID"),
                "is_inline": part.get_content_disposition() == "inline",
                "local_path": str(local_path),
                "download_status": "downloaded",
            }
        )

    date_header = message.get("Date")
    try:
        sent_datetime = parsedate_to_datetime(date_header).isoformat() if date_header else ""
    except (TypeError, ValueError, OverflowError):
        sent_datetime = str(date_header or "")
    sender = _email_people([message.get("From", "")])
    return {
        "subject": str(message.get("Subject") or "(No Subject)"),
        "sender": sender[0] if sender else {},
        "to_recipients": _email_people(message.get_all("To", [])),
        "cc_recipients": _email_people(message.get_all("Cc", [])),
        "bcc_recipients": _email_people(message.get_all("Bcc", [])),
        "sent_datetime": sent_datetime,
        "body_content_type": body_content_type,
        "body_content": body_content,
        "attachments": attachments,
    }


def _render_eml_attachment(source, target, libreoffice, depth):
    if depth >= MAX_NESTED_EMAIL_DEPTH:
        raise RuntimeError(
            f"Attached email nesting exceeds {MAX_NESTED_EMAIL_DEPTH} levels"
        )
    target = Path(target)
    with tempfile.TemporaryDirectory(
        prefix="attached-email-", dir=str(target.parent)
    ) as temporary_directory:
        temporary_directory = Path(temporary_directory)
        email = _email_from_eml(source, temporary_directory / "attachments")
        _create_email_pdf(email, target, libreoffice, depth + 1)


def _attachment_format(attachment):
    path = Path(attachment.get("local_path") or "")
    suffix = path.suffix.lower()
    if suffix:
        return suffix
    content_type = (attachment.get("content_type") or "").split(";", 1)[0].strip()
    return mimetypes.guess_extension(content_type) or ""


def _convert_attachment(attachment, target, libreoffice, depth):
    source = Path(attachment.get("local_path") or "")
    if attachment.get("download_status") != "downloaded" or not source.is_file():
        raise RuntimeError(
            f"Attachment {attachment.get('name') or source.name!r} was not downloaded"
        )

    extension = _attachment_format(attachment)
    if extension == ".pdf":
        return source
    if extension in TEXT_EXTENSIONS:
        _render_text_attachment(source, target)
    elif extension in HTML_EXTENSIONS:
        _render_html_attachment(source, target)
    elif extension in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"}:
        _render_image_attachment(source, target)
    elif extension == ".eml":
        _render_eml_attachment(source, target, libreoffice, depth)
    elif extension in OFFICE_EXTENSIONS:
        if not libreoffice:
            raise RuntimeError(
                f"Converting {source.name} requires LibreOffice; install it and "
                "ensure soffice is on PATH"
            )
        _render_office_attachment(source, target, libreoffice)
    else:
        raise RuntimeError(
            f"Unsupported attachment type for {source.name}: "
            f"{extension or attachment.get('content_type') or 'unknown'}"
        )
    return Path(target)


def _create_email_pdf(email, output_path, libreoffice, depth):
    from pypdf import PdfReader, PdfWriter

    if depth > MAX_NESTED_EMAIL_DEPTH:
        raise RuntimeError(
            f"Attached email nesting exceeds {MAX_NESTED_EMAIL_DEPTH} levels"
        )
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(
        prefix="email-pdf-", dir=str(output_path.parent)
    ) as temporary_directory:
        temporary_directory = Path(temporary_directory)
        email_pdf = temporary_directory / "email.pdf"
        _render_email_pdf(email, email_pdf)

        converted_pdfs = []
        embedded_content_ids = _referenced_content_ids(email)
        for index, attachment in enumerate(email.get("attachments", []), start=1):
            content_id = attachment.get("content_id")
            # CID-referenced parts are already rendered in the message body.
            if content_id and _cid_key(content_id) in embedded_content_ids:
                continue
            target = temporary_directory / f"attachment-{index:04d}.pdf"
            converted_pdfs.append(
                _convert_attachment(attachment, target, libreoffice, depth)
            )

        writer = PdfWriter()
        try:
            writer.append(str(email_pdf))
            for pdf_path in converted_pdfs:
                writer.append(str(pdf_path))
            writer.add_metadata({"/Title": email.get("subject") or "Email"})
            temporary_output = temporary_directory / "combined.pdf"
            with temporary_output.open("wb") as output:
                writer.write(output)
            # Validate the output before replacing the last successful PDF.
            with temporary_output.open("rb") as generated:
                PdfReader(generated).pages[0]
            temporary_output.replace(output_path)
        finally:
            writer.close()
    return output_path


def create_email_pdf(email, output_path):
    """Create one PDF containing the message followed by converted attachments."""
    return _create_email_pdf(email, output_path, _find_libreoffice(), depth=0)
