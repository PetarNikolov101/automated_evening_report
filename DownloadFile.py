import win32com.client
import os
from datetime import datetime, timedelta

SAVE_FOLDER = r"C:\Users\petarnik\skripta_neotstraneti\skripta_neotstraneti"
SENDER_EMAIL = "Branka.Atanasovska@telekom.mk"
SUBJECT_KEYWORD = "otvoreniprecki"
DAYS_BACK = 1  # 24h

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox

date_limit = (datetime.now() - timedelta(days=DAYS_BACK)).strftime("%m/%d/%Y %H:%M %p")

messages = inbox.Items
messages = messages.Restrict(f"[ReceivedTime] >= '{date_limit}'")
messages.Sort("[ReceivedTime]", True)

found = False
for message in messages:
    try:
        if message.Class != 43:  # skip non-mail items
            continue

        sender = message.SenderEmailAddress.strip()
        subject = message.Subject.strip()

        if SENDER_EMAIL in sender and SUBJECT_KEYWORD in subject:
            if message.Attachments.Count > 0:
                for i in range(1, message.Attachments.Count + 1):
                    attachment = message.Attachments.Item(i)
                    filename = attachment.FileName
                    if filename.lower().endswith(".xlsx"):
                        save_path = os.path.join(SAVE_FOLDER, filename)
                        attachment.SaveAsFile(save_path)
                        print(f"✅ Saved: {save_path}")
                        found = True
                        break  # ✅ stop after saving one attachment
            else:
                print("⚠️ Found mail but no attachment.")
            break  # ✅ stop after the first matching email
    except Exception as e:
        print(f"❌ Error processing message: {e}")
        break

if not found:
    print("No matching emails found.")
