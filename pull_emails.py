"""Sync unflagged Outlook messages into a local SQLite database."""

import argparse
import hashlib
import json
import re
from pathlib import Path
from urllib.parse import quote

from email_to_pdf import create_email_pdf
from email_processing import process_email
from email_store import EmailStore
from graph_client import GraphClient


ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "config.json"
DATA_ROOT = ROOT / "data"
DATABASE_PATH = DATA_ROOT / "emails.sqlite3"
ATTACHMENT_ROOT = DATA_ROOT / "attachments"
PDF_ROOT = DATA_ROOT / "pdfs"
CACHE_PATH = ROOT / "token_cache.bin"
ALLOWED_FLAG_STATUSES = ("flagged", "notFlagged", "complete")


def parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "Download notFlagged Microsoft Graph emails and attachments, "
            "store them locally, process them, and update their flag. "
            "Convert selected stored emails to PDF with --convert-to-pdf."
        )
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="List the five most recent Exchange emails",
    )
    parser.add_argument(
        "--list-stored",
        action="store_true",
        help="List recently synced emails available for PDF conversion",
    )
    parser.add_argument(
        "--convert-to-pdf",
        nargs="+",
        metavar="GRAPH_MESSAGE_ID",
        help="Convert one or more stored emails and append their attachments",
    )
    parser.add_argument(
        "--message-id",
        help="Set the follow-up flag on a message instead of syncing",
    )
    parser.add_argument(
        "--flag-status",
        choices=ALLOWED_FLAG_STATUSES,
        help="Flag state to set with --message-id",
    )
    args = parser.parse_args()

    if bool(args.message_id) != bool(args.flag_status):
        parser.error("--message-id and --flag-status must be provided together")
    selected_modes = sum(
        bool(value)
        for value in (
            args.list,
            args.list_stored,
            args.convert_to_pdf,
            args.message_id,
        )
    )
    if selected_modes > 1:
        parser.error(
            "--list, --list-stored, --convert-to-pdf, and --message-id "
            "are mutually exclusive"
        )
    return args


def load_config():
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"Missing {CONFIG_PATH.name}; copy config.example.json to config.json "
            "and set your client ID."
        )
    with CONFIG_PATH.open(encoding="utf-8") as config_file:
        config = json.load(config_file)

    client_id = config["client_id"]
    if not client_id or client_id == "YOUR_APPLICATION_CLIENT_ID_HERE":
        raise ValueError("Set your Microsoft application client ID in config.json")

    scopes = [scope for scope in config["scopes"] if scope != "Mail.Read"]
    if "Mail.ReadWrite" not in scopes:
        scopes.append("Mail.ReadWrite")
    return client_id, config["authority"], scopes


def acquire_access_token(client_id, authority, scopes):
    from msal import PublicClientApplication
    from msal_extensions import FilePersistenceWithDataProtection, PersistedTokenCache

    token_cache = PersistedTokenCache(
        FilePersistenceWithDataProtection(str(CACHE_PATH))
    )
    app = PublicClientApplication(
        client_id,
        authority=authority,
        token_cache=token_cache,
    )

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
    if not result:
        print("Opening browser for authentication...")
        result = app.acquire_token_interactive(scopes=scopes)
    if "access_token" not in result:
        raise RuntimeError(
            "Failed to acquire token: "
            + result.get("error_description", "unknown authentication error")
        )
    return result["access_token"]


def _safe_filename(name):
    name = Path(name or "attachment").name
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return name or "attachment"


def _attachment_target(message_id, attachment_id, name):
    message_folder = hashlib.sha256(message_id.encode("utf-8")).hexdigest()[:24]
    attachment_prefix = hashlib.sha256(
        attachment_id.encode("utf-8")
    ).hexdigest()[:12]
    return ATTACHMENT_ROOT / message_folder / f"{attachment_prefix}_{_safe_filename(name)}"


def _download_attachment(
    graph, store, message_id, attachment, attachment_type, attachment_index
):
    attachment_id = attachment["id"]
    store.upsert_attachment(
        message_id, attachment, attachment_type, attachment_index
    )

    downloadable = attachment_type in ("fileAttachment", "itemAttachment")
    if not downloadable:
        store.set_attachment_status(message_id, attachment_id, "not_downloadable")
        return

    existing = store.get_attachment(message_id, attachment_id)
    if (
        existing
        and existing["download_status"] == "downloaded"
        and existing["local_path"]
        and Path(existing["local_path"]).is_file()
    ):
        return

    target = _attachment_target(
        message_id, attachment_id, attachment.get("name") or "attachment"
    )
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = target.with_name(target.name + ".part")
    response = None
    digest = hashlib.sha256()
    try:
        response = graph.download_attachment(message_id, attachment_id)
        with temporary_path.open("wb") as output_file:
            for chunk in response.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    output_file.write(chunk)
                    digest.update(chunk)
        temporary_path.replace(target)
        store.set_attachment_downloaded(
            message_id, attachment_id, target, digest.hexdigest()
        )
    except Exception as error:
        temporary_path.unlink(missing_ok=True)
        store.set_attachment_status(
            message_id, attachment_id, "failed", str(error)
        )
        raise
    finally:
        if response is not None:
            response.close()


def sync_emails(graph, store):
    succeeded = 0
    failed = 0

    for message in graph.list_not_flagged_messages():
        message_id = message.get("id")
        if not message_id:
            print("Skipping a Graph message without an ID")
            failed += 1
            continue

        stage = "download_failed"
        try:
            store.upsert_email(message)
            # Query explicitly because Graph's hasAttachments flag can omit
            # inline attachments.
            for attachment_index, attachment in enumerate(
                graph.list_attachments(message_id)
            ):
                raw_type = attachment.get("@odata.type", "")
                attachment_type = raw_type.rsplit(".", 1)[-1].lstrip("#")
                if not attachment_type:
                    attachment_type = "fileAttachment"
                _download_attachment(
                    graph,
                    store,
                    message_id,
                    attachment,
                    attachment_type,
                    attachment_index,
                )

            email = store.get_email(message_id)
            stage = "processing_failed"
            flag_status = process_email(email)
            if flag_status not in ALLOWED_FLAG_STATUSES or flag_status == "notFlagged":
                raise ValueError(
                    "process_email must return 'flagged' or 'complete' "
                    "so the message is not pulled again"
                )

            stage = "flag_update_failed"
            graph.set_flag(message_id, flag_status)
            store.set_processing_status(
                message_id, "processed", flag_status=flag_status
            )
            succeeded += 1
            print(f"Processed {message_id}: flag set to {flag_status}")
        except Exception as error:
            store.set_processing_status(message_id, stage, error=str(error))
            failed += 1
            print(f"Failed to process {message_id}: {error}")

    print(f"\nSync finished: {succeeded} processed, {failed} failed.")


def list_recent_emails(graph):
    messages = list(graph.list_recent_messages(limit=5))[:5]
    print(f"\nLatest {len(messages)} emails:\n")
    for message in messages:
        sender = (message.get("from") or {}).get("emailAddress", {}).get(
            "name", "Unknown Sender"
        )
        subject = message.get("subject") or "(No Subject)"
        received = message.get("receivedDateTime", "")
        flag_status = (message.get("flag") or {}).get("flagStatus", "unknown")
        print(f"- ID: {message.get('id', '')}")
        print(
            f"  [{received[:10]}] From: {sender} | Subject: {subject} "
            f"| Flag: {flag_status}"
        )


def list_stored_emails(store, limit=20):
    emails = store.list_recent_emails(limit)
    print(f"\nLatest {len(emails)} stored emails:\n")
    for email in emails:
        sender = (email.get("sender") or {}).get("emailAddress") or {}
        subject = email.get("subject") or "(No Subject)"
        received = email.get("received_datetime") or ""
        pdf_path = email.get("pdf_path")
        pdf_status = "PDF available" if pdf_path and Path(pdf_path).is_file() else "no PDF"
        print(f"- ID: {email['graph_id']}")
        print(
            f"  [{received[:10]}] From: {sender.get('name') or sender.get('address', 'Unknown Sender')} "
            f"| Subject: {subject} | {pdf_status}"
        )


def convert_email_to_pdf(store, message_id):
    email = store.get_email(message_id)
    if email is None:
        raise ValueError(
            f"Email {message_id!r} is not stored locally; run a sync first"
        )
    output_path = PDF_ROOT / (
        hashlib.sha256(message_id.encode("utf-8")).hexdigest()[:24] + ".pdf"
    )
    create_email_pdf(email, output_path)
    store.set_pdf_path(message_id, output_path)
    return output_path


def main():
    args = parse_args()

    if args.list_stored or args.convert_to_pdf:
        store = EmailStore(DATABASE_PATH)
        try:
            if args.list_stored:
                list_stored_emails(store)
            else:
                failed = False
                for message_id in args.convert_to_pdf:
                    try:
                        output_path = convert_email_to_pdf(store, message_id)
                        print(f"Created {output_path}")
                    except Exception as error:
                        failed = True
                        print(f"Failed to convert {message_id}: {error}")
                if failed:
                    raise SystemExit(1)
        finally:
            store.close()
        return

    client_id, authority, scopes = load_config()
    access_token = acquire_access_token(client_id, authority, scopes)
    graph = GraphClient(access_token)

    if args.message_id:
        graph.set_flag(args.message_id, args.flag_status)
        print(f"Updated message flag to {args.flag_status}.")
    elif args.list:
        list_recent_emails(graph)
    else:
        DATA_ROOT.mkdir(parents=True, exist_ok=True)
        store = EmailStore(DATABASE_PATH)
        try:
            sync_emails(graph, store)
        finally:
            store.close()


if __name__ == "__main__":
    main()
