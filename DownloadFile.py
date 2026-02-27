import os
import json
import requests
import msal
import base64
from datetime import datetime, timedelta, timezone
from urllib.parse import quote


# ========================
# CONFIGURATION
# ========================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(BASE_DIR, "mejlovi.json"), "r", encoding="utf-8") as f:
    mejlovi = json.load(f)

with open(os.path.join(BASE_DIR, "credentials.json"), "r", encoding="utf-8") as f:
    creds = json.load(f)

SAVE_FOLDER = r"C:\Users\petarnik\skripta_neotstraneti\skripta_neotstraneti"
MAILBOX = creds["shared_mailbox"]

EXPECTED_SENDER = mejlovi["svc"].lower()
SUBJECT_KEYWORD = "otvoreniprecki"
DAYS_BACK = 1

DEBUG = True


# ========================
# AUTHENTICATION
# ========================

def get_access_token():
    app = msal.ConfidentialClientApplication(
        creds["client_id"],
        authority=f"https://login.microsoftonline.com/{creds['tenant_id']}",
        client_credential=creds["client_secret"],
    )

    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    if "access_token" not in result:
        raise Exception(f"Could not acquire token: {result}")

    return result["access_token"]


# ========================
# GRAPH HELPER
# ========================

def graph_get(url, headers):
    response = requests.get(url, headers=headers)

    if DEBUG:
        print("\nREQUEST:", url)
        print("STATUS:", response.status_code)

    if response.status_code >= 400:
        raise Exception(f"Graph API error: {response.text}")

    return response.json()


# ========================
# FIND MATCHING MESSAGE
# ========================

def find_matching_message(headers):

    date_limit = (
        datetime.now(timezone.utc) - timedelta(days=DAYS_BACK)
    ).isoformat().replace("+00:00", "Z")

    # Only use SAFE Graph filters
    filter_query = (
        f"receivedDateTime ge {date_limit} "
        f"and hasAttachments eq true"
    )

    encoded_filter = quote(filter_query)

    url = (
        f"https://graph.microsoft.com/v1.0/users/{MAILBOX}/messages"
        f"?$filter={encoded_filter}"
        "&$orderby=receivedDateTime desc"
        "&$top=10"
    )

    data = graph_get(url, headers)
    messages = data.get("value", [])

    if not messages:
        raise Exception("No recent messages found.")

    for message in messages:

        subject = (message.get("subject") or "").lower()
        sender = (
            message.get("from", {})
            .get("emailAddress", {})
            .get("address", "")
            .lower()
        )

        if DEBUG:
            print("\nChecking message:")
            print("Subject:", subject)
            print("Sender:", sender)

        if SUBJECT_KEYWORD.lower() in subject and sender == EXPECTED_SENDER:
            print("\n Matching email found.")
            return message

    raise Exception("No matching email found after filtering.")


# ========================
# DOWNLOAD ATTACHMENT
# ========================

def download_excel_attachment(message_id, headers):

    url = f"https://graph.microsoft.com/v1.0/users/{MAILBOX}/messages/{message_id}/attachments"
    data = graph_get(url, headers)

    attachments = data.get("value", [])

    if not attachments:
        raise Exception("Email found but has no attachments.")

    os.makedirs(SAVE_FOLDER, exist_ok=True)

    for attachment in attachments:

        if attachment.get("@odata.type") != "#microsoft.graph.fileAttachment":
            continue

        filename = attachment.get("name", "")

        if not filename.lower().endswith((".xlsx", ".xls")):
            continue

        print(f"\nFound Excel attachment: {filename}")

        # Large attachment support
        content_bytes = attachment.get("contentBytes")

        if content_bytes:
            file_bytes = base64.b64decode(content_bytes)
        else:
            # Fallback for large files
            download_url = attachment.get("@microsoft.graph.downloadUrl")
            file_response = requests.get(download_url)
            file_bytes = file_response.content

        save_path = os.path.join(SAVE_FOLDER, "otvoreniprecki.xlsx")

        with open(save_path, "wb") as f:
            f.write(file_bytes)

        print(f"File saved to: {save_path}")
        return

    raise Exception("No Excel attachment found.")


# ========================
# MAIN
# ========================

def main():

    print("Authenticating...")
    token = get_access_token()

    headers = {
        "Authorization": f"Bearer {token}"
    }

    print("Searching for email...")
    message = find_matching_message(headers)

    print("Downloading attachment...")
    download_excel_attachment(message["id"], headers)

    print("\nDone successfully.")


if __name__ == "__main__":
    main()