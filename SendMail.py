import pandas as pd
from datetime import datetime
import os
import json
import smtplib
from email.message import EmailMessage
from email.utils import formatdate

# setup

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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

# read data

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

df_grupni = pd.read_excel(excel_path, sheet_name="Overdue Grupni Utre", engine='openpyxl')
html_table_grupni = df_grupni.to_html(index=False, border=1, justify='left', na_rep='')

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

csod = 0
try:
    first_col = df_csod.columns[0]
    mask = df_csod[first_col].astype(str).str.strip().str.upper() == 'ВКУПНО'
    if mask.any() and 'Вкупно' in df_csod.columns:
        csod = int(pd.to_numeric(df_csod.loc[mask, 'Вкупно'].iloc[0], errors='coerce') or 0)
except Exception:
    csod = 0

# styloing

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

html_table_edinechni = style_table(html_table_edinechni)
html_table_grupni = style_table(html_table_grupni)
html_table_csod = style_table(html_table_csod)

# html mail

html_body = f"""
<html>
  <body style="font-family: Aptos Narrow, Aptos, Calibri, Arial, sans-serif; font-size:10pt;">
    <p>Колеги,</p>

    <p>Во моментот во SSOD имаме <b>{edinechni}</b> незатворени единечни пречки.</p>

    <p>
      Отворените пречки со промашен таргет, пречки со таргет до утре до 16 часот
      и тековна се прикажани по регион и доделен техничар:
    </p>

    {html_table_edinechni}

    <p>
      Ве молам пристапете кон решавање на пречките за да не истече таргет времето
      за корисниците (дефинирано до утре до 16 часот).
    </p>

    <p>
      Во моментот имаме <b>{grupni}</b> единечни пречки кои се поврзани со групен прекин:
    </p>

    {html_table_grupni}

    <p>
      Ве молам пристапете кон решавање / проверка на групните пречки.
    </p>

    <p>
      Во моментот има <b>{csod}</b> пречки отворени за CSOD.
    </p>

    {html_table_csod}

    <p>Поздрав,<br>Петар Николов</p>
  </body>
</html>
"""


# recipients = [
#     mejlovi['Snezhana'],
#     mejlovi['Klimentina'],
#     mejlovi['Dimitar'],
#     mejlovi['Maja'],
#     mejlovi['Elizabeta'],
#     mejlovi['Regionalni_Ofisi'],
#     mejlovi['CTSO'],
#     mejlovi['Anastas'],
#     mejlovi['Kelmend'],
#     mejlovi['Goran'],
#     mejlovi['Irena'],
#     mejlovi['Tatjana'],
#     mejlovi['Zanet'],
#     mejlovi['Emilija'],
#     mejlovi['CTSO_disp'],
#     mejlovi['CSODGPON'],
#     mejlovi['CSODADSL'],
# ]

recipients = [
    mejlovi['Pero']
]

msg = EmailMessage()
msg['To'] = ", ".join(recipients)
msg['Subject'] = f"Lista na precki - TT {date_str}"
msg['Date'] = formatdate(localtime=True)

msg.set_content("Вашиот клиент не поддржува HTML.")
msg.add_alternative(html_body, subtype='html')

# attach excel file

with open(excel_path, 'rb') as f:
    msg.add_attachment(
        f.read(),
        maintype='application',
        subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename=name_of_excel
    )

# send

SMTP_SERVER = credentials["SMTP_server"]
SMTP_PORT = credentials["SMTP_port"]

with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
    server.send_message(msg)

#cleanup

os.remove(excel_path)
os.remove(os.path.join(BASE_DIR, "otvoreniprecki.xlsx"))