"""Small Microsoft Graph client for email synchronization."""

from urllib.parse import quote

import requests


GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
MESSAGE_SELECT = ",".join(
    (
        "id",
        "internetMessageId",
        "subject",
        "from",
        "toRecipients",
        "ccRecipients",
        "bccRecipients",
        "replyTo",
        "body",
        "receivedDateTime",
        "sentDateTime",
        "isRead",
        "importance",
        "conversationId",
        "hasAttachments",
        "flag",
    )
)


class GraphClient:
    def __init__(self, access_token):
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
            'Prefer': 'IdType="ImmutableId"',
        }

    def _get_json(self, url, params=None):
        response = requests.get(
            url, headers=self.headers, params=params, timeout=60
        )
        response.raise_for_status()
        return response.json()

    def _iter_pages(self, first_page):
        page = first_page
        while page:
            yield from page.get("value", [])
            next_link = page.get("@odata.nextLink")
            if not next_link:
                break
            page = self._get_json(next_link)

    def list_not_flagged_messages(self):
        url = f"{GRAPH_ROOT}/me/messages"
        params = {
            "$top": 50,
            "$select": MESSAGE_SELECT,
            "$filter": "flag/flagStatus eq 'notFlagged'",
        }
        try:
            first_page = self._get_json(url, params=params)
        except requests.HTTPError as error:
            # Some Graph tenants reject filtering on the nested flag property.
            # Fall back to paging through messages and filtering locally.
            if not error.response or error.response.status_code != 400:
                raise
            params.pop("$filter")
            first_page = self._get_json(url, params=params)

        for message in self._iter_pages(first_page):
            flag_status = (message.get("flag") or {}).get("flagStatus")
            if flag_status == "notFlagged":
                yield message

    def list_recent_messages(self, limit=5):
        page = self._get_json(
            f"{GRAPH_ROOT}/me/messages",
            params={
                "$top": limit,
                "$select": "id,subject,from,receivedDateTime,isRead,flag",
                "$orderby": "receivedDateTime desc",
            },
        )
        yield from self._iter_pages(page)

    def list_attachments(self, message_id):
        encoded_message_id = quote(message_id, safe="")
        url = f"{GRAPH_ROOT}/me/messages/{encoded_message_id}/attachments"
        page = self._get_json(url, params={"$top": 50})
        yield from self._iter_pages(page)

    def download_attachment(self, message_id, attachment_id):
        encoded_message_id = quote(message_id, safe="")
        encoded_attachment_id = quote(attachment_id, safe="")
        url = (
            f"{GRAPH_ROOT}/me/messages/{encoded_message_id}/attachments/"
            f"{encoded_attachment_id}/$value"
        )
        response = requests.get(
            url,
            headers=self.headers,
            timeout=120,
            stream=True,
        )
        response.raise_for_status()
        return response

    def set_flag(self, message_id, flag_status):
        encoded_message_id = quote(message_id, safe="")
        url = f"{GRAPH_ROOT}/me/messages/{encoded_message_id}"
        response = requests.patch(
            url,
            headers=self.headers,
            json={"flag": {"flagStatus": flag_status}},
            timeout=60,
        )
        response.raise_for_status()
