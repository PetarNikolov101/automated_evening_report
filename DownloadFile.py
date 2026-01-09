import win32com.client
import os
from datetime import datetime, timedelta
import json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

with open(os.path.join(BASE_DIR, 'mejlovi.json'), 'r', encoding='utf-8') as f:
    mejlovi = json.load(f)

SAVE_FOLDER = r"C:\Users\petarnik\skripta_neotstraneti\skripta_neotstraneti"
SENDER_EMAIL = mejlovi['Branka']
SUBJECT_KEYWORD = "otvoreniprecki"
DAYS_BACK = 1  # last 24h
DEBUG = True

outlook = win32com.client.Dispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")
inbox = ns.GetDefaultFolder(6)  # Inbox

date_limit = datetime.now() - timedelta(days=DAYS_BACK)
target_email = None

items = inbox.Items
try:
    items.Sort("[ReceivedTime]", True)  # newest first
except Exception:
    pass

for msg in items:
    try:
        if getattr(msg, "Class", None) != 43:  # MailItem
            continue

        subject = getattr(msg, "Subject", "") or ""
        sender = getattr(msg, "SenderEmailAddress", "") or ""

        try:
            received = msg.ReceivedTime
        except Exception:
            received = None

        if received is not None:
            try:
                if received.replace(tzinfo=None) < date_limit:
                    break
            except Exception:
                pass

        if SENDER_EMAIL.lower() == sender.lower() and \
           SUBJECT_KEYWORD.lower() in subject.lower():
            target_email = msg
            break

    except Exception:
        continue

if not target_email:
    raise Exception("No matching email found!")

print("Found email:", getattr(target_email, "Subject", "<no subject>"))

# ---- SAVE EXCEL ATTACHMENT ----

os.makedirs(SAVE_FOLDER, exist_ok=True)

saved = False
today = datetime.now().strftime("%Y-%m-%d")

for i in range(1, target_email.Attachments.Count + 1):
    try:
        attachment = target_email.Attachments.Item(i)
        filename = attachment.FileName or ""

        if filename.lower().endswith(".xlsx"):
            dest_name = f"otvoreniprecki.xlsx"
            dest_path = os.path.join(SAVE_FOLDER, dest_name)

            attachment.SaveAsFile(dest_path)
            print("✅ Saved:", dest_path)
            saved = True
            break

    except Exception as e:
        if DEBUG:
            print("Attachment error:", e)

if not saved:
    raise Exception("Matching email found, but no Excel attachment!")