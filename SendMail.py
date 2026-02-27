import os
import json
import re
import base64
import pandas as pd
from datetime import datetime
import requests
import msal


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# read emails and credentials from jsons
with open(os.path.join(BASE_DIR, 'mejlovi.json'), 'r', encoding='utf-8') as f:
    mejlovi = json.load(f)

with open(os.path.join(BASE_DIR, 'credentials.json'), 'r', encoding='utf-8') as f:
    credentials = json.load(f)

date = datetime.today()
date_str = date.strftime("%d-%m-%Y")

name_of_excel = f"Lista na precki - TT {date_str}.xlsx"
excel_path = os.path.join(
    r'C:\Users\petarnik\skripta_neotstraneti\skripta_neotstraneti\kreirani datoteki',
    name_of_excel
)

#styling for the tables
def widen_column_by_header(html, header_name, width_px=220):
    headers = re.findall(r'<th.*?>(.*?)</th>', html, flags=re.DOTALL)
    headers = [re.sub('<.*?>', '', h).strip() for h in headers]

    if header_name not in headers:
        return html

    col_index = headers.index(header_name) + 1  # nth-child is 1-based

    html = re.sub(
        fr'(<th[^>]*>:?{header_name}</th>)',
        fr'<th style="width:{width_px}px; min-width:{width_px}px;">{header_name}</th>',
        html
    )

    html = re.sub(
        fr'(<tr[^>]*>(?:.*?</td>){{{col_index-1}}})(<td)',
        fr'\1<td style="width:{width_px}px; min-width:{width_px}px; white-space:nowrap;">',
        html,
        flags=re.DOTALL
    )

    return html

def style_table(html):
    return (
        html.replace(
            '<table ',
            '<table style="font-family: Aptos Narrow, Aptos, Calibri, Arial, sans-serif; '
            'font-size:10pt; border-collapse:collapse;" '
        )
        .replace('<th>', '<th style="padding:4px;">')
        .replace('<td>', '<td style="padding:4px;">')
    )

#this reads the excel file

df_summery = pd.read_excel(excel_path, sheet_name="Summery", engine='openpyxl')
edinechni = 0
try:
    mask = df_summery['A'].astype(str).str.strip() == "Вкупно отворени пречки на Ниво 3"
    if mask.any():
        edinechni = int(pd.to_numeric(df_summery.loc[mask, 'B'].iloc[0], errors='coerce') or 0)
except Exception:
    edinechni = 0

df_edinechni = pd.read_excel(excel_path, sheet_name="Overdue Utre", engine='openpyxl')
html_table_edinechni = df_edinechni.to_html(index=False, border=1, justify='left', na_rep='')
html_table_edinechni = widen_column_by_header(html_table_edinechni, "Техничар", width_px=240)
html_table_edinechni = style_table(html_table_edinechni)

df_grupni = pd.read_excel(excel_path, sheet_name="Overdue Grupni Utre", engine='openpyxl')
html_table_grupni = df_grupni.to_html(index=False, border=1, justify='left', na_rep='')
html_table_grupni = style_table(html_table_grupni)

grupni = 0
try:
    first_col = df_grupni.columns[0]
    mask = df_grupni[first_col].astype(str).str.strip().str.upper() == 'ВКУПНО'
    if mask.any() and 'Вкупно' in df_grupni.columns:
        grupni = int(pd.to_numeric(df_grupni.loc[mask, 'Вкупно'].iloc[0], errors='coerce') or 0)
except Exception:
    grupni = 0

df_csod = pd.read_excel(excel_path, sheet_name="CSOD", engine='openpyxl')
html_table_csod = df_csod.to_html(index=False, border=1, justify='left', na_rep='')

html_table_csod = style_table(html_table_csod)

csod = 0
try:
    first_col = df_csod.columns[0]
    mask = df_csod[first_col].astype(str).str.strip().str.upper() == 'ВКУПНО'
    if mask.any() and 'Вкупно' in df_csod.columns:
        csod = int(pd.to_numeric(df_csod.loc[mask, 'Вкупно'].iloc[0], errors='coerce') or 0)
except Exception:
    csod = 0

#create the mail body

html_body = f"""
<html>
  <body style="font-family: Aptos Narrow, Aptos, Calibri, Arial, sans-serif; font-size:10pt;">
    <p>Колеги,</p>

    <p>Во моментот во SSOD имаме <b>{edinechni}</b> незатворени единечни пречки.</p>

    <p>Отворените пречки со промашен таргет, пречки со таргет до утре до 16 часот
       и тековна се прикажани по регион и доделен техничар:</p>

    {html_table_edinechni}

    <p>Ве молам пристапете кон решавање на пречките за да не истече таргет времето
       за корисниците (дефинирано до утре до 16 часот).</p>

    <p>Во моментот имаме <b>{grupni}</b> единечни пречки кои се поврзани со групен прекин:</p>

    {html_table_grupni}

    <p>Ве молам пристапете кон решавање / проверка на групните пречки.</p>

    <p>Во моментот има <b>{csod}</b> пречки отворени за CSOD.</p>

    {html_table_csod}

    <p>Поздрав,<br>Петар Николов</p>
  </body>
</html>
"""


recipients = [
    mejlovi['Snezhana'],
    mejlovi['Klimentina'],
    mejlovi['Dimitar'],
    mejlovi['Maja'],
    mejlovi['Elizabeta'],
    mejlovi['Regionalni_Ofisi'],
    mejlovi['CTSO'],
    mejlovi['Anastas'],
    mejlovi['Kelmend'],
    mejlovi['Goran'],
    mejlovi['Irena'],
    mejlovi['Tatjana'],
    mejlovi['Zanet'],
    mejlovi['Emilija'],
    mejlovi['CTSO_disp'],
    mejlovi['CSODGPON'],
    mejlovi['CSODADSL'],
    mejlovi['Pero']
]


tenant_id = credentials["tenant_id"]
client_id = credentials["client_id"]
client_secret = credentials["client_secret"]

authority = f"https://login.microsoftonline.com/{tenant_id}"
scope = ["https://graph.microsoft.com/.default"]

app = msal.ConfidentialClientApplication(
    client_id,
    authority=authority,
    client_credential=client_secret
)

token_response = app.acquire_token_for_client(scopes=scope)
if "access_token" not in token_response:
    raise Exception(f"Cannot acquire token: {token_response}")

access_token = token_response["access_token"]

with open(excel_path, "rb") as f:
    file_content = base64.b64encode(f.read()).decode("utf-8")


graph_message = {
    "message": {
        "subject": f"Lista na precki - TT {date_str}",
        "body": {"contentType": "HTML", "content": html_body},
        "toRecipients": [{"emailAddress": {"address": r}} for r in recipients],
        "attachments": [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": name_of_excel,
                "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "contentBytes": file_content
            }
        ]
    },
    "saveToSentItems": True
}

# sender mailbox address
sender = credentials["shared_mailbox"]

send_url = f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

response = requests.post(send_url, headers=headers, json=graph_message)
if response.status_code != 202:
    raise Exception(f"Mail send failed: {response.status_code} {response.text}")

print("Mail sent successfully via Graph.")


os.remove(excel_path)
os.remove(os.path.join(BASE_DIR, "otvoreniprecki.xlsx"))