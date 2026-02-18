import os
import json
import requests
import msal
from datetime import datetime, timedelta
import base64

# setup

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(BASE_DIR, 'mejlovi.json'), 'r', encoding='utf-8') as f:
    mejlovi = json.load(f)

with open(os.path.join(BASE_DIR, 'credentials.json'), 'r', encoding='utf-8') as f:
    creds = json.load(f)

SAVE_FOLDER = r"C:\Users\petarnik\skripta_neotstraneti\skripta_neotstraneti"
SENDER_EMAIL = mejlovi['Pero']
SUBJECT_KEYWORD = "otvoreniprecki"
DAYS_BACK = 1
DEBUG = True


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
        raise Exception(f"Could not get token: {result}")

    return result["access_token"]


def main():
    token = get_access_token()

    headers = {
        "Authorization": f"Bearer {token}"
    }

    date_limit = (datetime.now() - timedelta(days=DAYS_BACK)).isoformat() + "Z"

    # filter emails
    filter_query = (
        f"receivedDateTime ge {date_limit} "
        f"and from/emailAddress/address eq '{SENDER_EMAIL}' "
        f"and contains(subject,'{SUBJECT_KEYWORD}')"
    )

    url = (
        "https://graph.microsoft.com/v1.0/users/"
        f"{creds['shared_mailbox']}/mailFolders/inbox/messages"
        f"?$filter={filter_query}"
        "&$orderby=receivedDateTime desc"
        "&$top=1"
    )

    response = requests.get(url, headers=headers)
    data = response.json()

    if DEBUG:
        print("Graph response:", data)

    messages = data.get("value", [])

    if not messages:
        raise Exception("No matching email found!")

    message = messages[0]
    message_id = message["id"]

    print("Found email:", message.get("subject"))

    # get attachment
    att_url = f"https://graph.microsoft.com/v1.0/users/{creds['shared_mailbox']}/messages/{message_id}/attachments"
    att_response = requests.get(att_url, headers=headers)
    att_data = att_response.json()
    print("Attachment response:", att_data)

    attachments = att_data.get("value", [])

    if not attachments:
        raise Exception("Matching email found, but no attachments!")

    os.makedirs(SAVE_FOLDER, exist_ok=True)

    saved = False

    for attachment in attachments:
        name = attachment.get("name", "")
        if name.lower().endswith(".xlsx"):

            content_bytes = attachment.get("contentBytes")

            if content_bytes:
                file_bytes = base64.b64decode(content_bytes)

                dest_path = os.path.join(SAVE_FOLDER, "otvoreniprecki.xlsx")

                with open(dest_path, "wb") as f:
                    f.write(file_bytes)

                print("✅ Saved:", dest_path)
                saved = True
                break

    if not saved:
        raise Exception("Matching email found, but no Excel attachment!")


if __name__ == "__main__":
    main()
    