import win32com.client
import pythoncom
import pandas as pd
from datetime import date, timedelta, datetime
import os

pythoncom.CoInitialize()

outlook = win32com.client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)

date = datetime.today()
date_str = date.strftime("%d-%m-%Y")
name_of_excel = f"Lista na precki - TT {date_str}.xlsx"
excel_path = os.path.join(r'C:\Users\petarnik\skripta_neotstraneti\skripta_neotstraneti\kreirani datoteki', name_of_excel)
#message.To = "Klimentina.Tunteva@telekom.mk;Snezana.Tilovska@telekom.mk;Maja.Ivanovska@telekom.mk;Elizabeta.Arsova@telekom.mk;Regionalni_Ofisi.mk@telekom.mk;CTSOTechnicalSupport@telekom.mk;Emilija.Andonovska@telekom.mk;Anastas.Miskovski@telekom.mk;Dimitar.Mukaetov@telekom.mk;Kelmend.Ziba@telekom.mk;Goran.Jovanoski@telekom.mk;Irena.Lokvenec@telekom.mk;Tatjana.Korpaseva@telekom.mk;Zanet.Tanturovska@telekom.mk;CTSO_Skopje_Dispatchers.mk@telekom.mk;CSODGPON@telekom.mk;CSODADSL@telekom.mk" 
message.To = "Petar.Nikolov@telekom.mk; Snezana.Tilovska@telekom.mk;CSODGPON@telekom.mk;CSODADSL@telekom.mk;Klimentina.Tunteva@telekom.mk"
#message.To = "Petar.Nikolov@telekom.mk"
message.Subject = f'Lista na precki - TT {date_str}'

df_summery = pd.read_excel(excel_path, sheet_name="Summery", engine='openpyxl')
edinechni = 0
try:
    # Find the row where column 'A' is "Вкупно отворени пречки на Ниво 3"
    mask = df_summery['A'].astype(str).str.strip() == "Вкупно отворени пречки на Ниво 3"
    if mask.any():
        edinechni = int(pd.to_numeric(df_summery.loc[mask, 'B'].iloc[0], errors='coerce') or 0)
except Exception:
    edinechni = 0

df_edinechni = pd.read_excel(excel_path, sheet_name = "Overdue Utre", engine = 'openpyxl')
html_table_edinechni = df_edinechni.to_html(index=False, border=1, justify='left', na_rep='')

df_grupni = pd.read_excel(excel_path, sheet_name="Overdue Grupni Utre", engine='openpyxl')
html_table_grupni = df_grupni.to_html(index = False,  border=1, justify='left', na_rep='')
grupni = 0
try:
    first_col = df_grupni.columns[0]
    # look for a row where the first column equals 'ВКУПНО' (case-insensitive)
    mask = df_grupni[first_col].astype(str).str.strip().str.upper() == 'ВКУПНО'
    if mask.any() and 'Вкупно' in df_grupni.columns:
        grupni = int(pd.to_numeric(df_grupni.loc[mask, 'Вкупно'].iloc[0], errors='coerce') or 0)
    elif 'Вкупно' in grupni.columns:
        # fallback: use last non-null value in the 'Вкупно' column
        last = df_grupni['Вкупно'].dropna()
        if not last.empty:
            edinechni = int(pd.to_numeric(last.iloc[-1], errors='coerce') or 0)
except Exception:
    grupni = 0

df_csod = pd.read_excel(excel_path, sheet_name="CSOD", engine = "openpyxl")
html_table_csod = df_csod.to_html(index = False, border = 1, justify="left", na_rep="")
csod = 0

try:
    first_col = df_csod.columns[0]  # should be "Класификација"
    mask = df_csod[first_col].astype(str).str.strip().str.upper() == 'ВКУПНО'

    if mask.any() and 'Вкупно' in df_csod.columns:
        csod = int(pd.to_numeric(df_csod.loc[mask, 'Вкупно'].iloc[0], errors='coerce') or 0)
    elif 'Вкупно' in df_csod.columns:
        last = df_csod['Вкупно'].dropna()
        if not last.empty:
            csod = int(pd.to_numeric(last.iloc[-1], errors='coerce') or 0)
except Exception as e:
    print("Error reading CSOD total:", e)
    csod = 0

html_body = f"""
<html>
  <body>
    <p>Колеги,</p>
    <p> Во моментот во SSOD имаме {edinechni} незатворени единечни пречки.</p>

    <p> Отворените пречки со промашен таргет во моментот, пречки со таргет до утре до 16 часот и тековна се прикажани по регион и доделен техничар: </p>

    {html_table_edinechni}

    <p>Ве молам пристапете кон решавање на пречките за да не истече таргет времето за корисниците (дефинирано до утре до 16 часот).</p>
    <p>Во моментот имаме {grupni} единечни пречки кои се поврзани со групен прекин со класификација:</p>
    {html_table_grupni}
    <p>Ве молам пристапете кон решавање/ проверка  на групните пречките за да не истече таргет времето за корисниците.</p>

    <p> Во моментот има {csod} пречки отворени за CSOD.
    {html_table_csod}
    <p> Поздрав, </p>
    <p> Петар Николов.</p>
  
  </body>
</html>
"""

message.HTMLBody = html_body

message.Attachments.Add(excel_path)

message.Send()

os.remove(excel_path)
os.remove("otvoreniprecki.xlsx")
