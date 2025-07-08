import pandas as pd
from datetime import date, timedelta, datetime
import numpy as np

# read from html
file_path = input("Внеси име на датотека од која се читаат податоци: ")

dfs = pd.read_html(file_path + ".xls")
df = dfs[0]
df.to_excel('kreirani datoteki/converted_neotstraneti.xlsx', index=False)
copy_from = pd.read_excel("kreirani datoteki/converted_neotstraneti.xlsx", index_col=None)

today = datetime.today()
tomorrow = date.today() + timedelta(days=1)
after_tomorrow = date.today() + timedelta(days=2)
tomorrow_4pm = datetime.combine(tomorrow, datetime.strptime("16:00:00", "%H:%M:%S").time())

excel_name = input("Внеси име на датотека: ")

print("Extracting columns...")
# columns to extract
columns_to_extract = [
    {"columnName": "Пречка", "index": 0, "letterOfColumn": "A"},
    {"columnName": "LineID", "index": 1, "letterOfColumn": "B"},
    {"columnName": "Статус", "index": 2, "letterOfColumn": "C"},
    {"columnName": "Дата на пријава", "index": 4, "letterOfColumn": "D"},
    {"columnName": "Категорија", "index": 9, "letterOfColumn": "J"},
    {"columnName": "Групна", "index": 12, "letterOfColumn": "M"},
    {"columnName": "Групна класификација", "index": 13, "letterOfColumn": "N"},
    {"columnName": "Техничар", "index": 17, "letterOfColumn": "R"},
    {"columnName": "Last remark WFM", "index": 18, "letterOfColumn": "S"},
    {"columnName": "ТТ налог креиран", "index": 20, "letterOfColumn": "U"},
    {"columnName": "Статус налог", "index": 22, "letterOfColumn": "W"},
    {"columnName": "Име на регион", "index": 23, "letterOfColumn": "X"},
    {"columnName": "Доделена група", "index": 24, "letterOfColumn": "Y"},
    {"columnName": "Посакуван крај", "index": 44, "letterOfColumn": "AS"}
]

print("Filtering SSOD...")
# extract columns
selected_columns = {
    col["columnName"]: copy_from.iloc[1:, col["index"]]
    for col in columns_to_extract
}

selected_df = pd.DataFrame(selected_columns)

# convert selected_columns back to a dictionary
selected_columns = selected_df.to_dict('list')

pomoshen_df = pd.DataFrame(selected_columns)
# convert Посакуван крај to datetime
pomoshen_df['Посакуван крај'] = pd.to_datetime(pomoshen_df['Посакуван крај'], errors='coerce', dayfirst=True)

print("Filtering overdue issues...")

# filter overdue for SSOD
data_df = pomoshen_df[(pomoshen_df['Доделена група'] == "SSOD")&(pomoshen_df["Статус налог"] != "ОТКАЖАН")]
overdue_df = data_df[(data_df['Посакуван крај'].dt.date < today.date())]

print("Grouping by region, technician, and counting their overdue issues...")
# group by region and technician, count overdue issues
overdue_summary = overdue_df.groupby(['Име на регион', 'Техничар']).size().reset_index(name='Вкупно')

# group issues for Overdue Grupni Utre
overdue_grupni_df = pomoshen_df[(pomoshen_df['Групна'].notna()) & (pomoshen_df['Посакуван крај'] < tomorrow_4pm)]
overdue_grupni = overdue_grupni_df.groupby(['Име на регион', 'Техничар']).size().reset_index(name='Вкупно')

# filter Overdue Utre
overdue_utre_df = data_df[(data_df['Групна'].isna()) & (pomoshen_df['Посакуван крај'].notna())]

# calculate Пробиена утре 4PM (not overdue now, but will be by tomorrow 4 PM)
probieni_utre = overdue_utre_df[
    (overdue_utre_df['Посакуван крај'] > tomorrow_4pm)
].groupby(['Име на регион', 'Техничар']).size().reset_index(name='Пробиена утре 4PM')

# calculate Тековна (current issues, not overdue by tomorrow 4 PM) ?
tekovni = overdue_utre_df[overdue_utre_df['Посакуван крај'] >= tomorrow_4pm].groupby(
    ['Име на регион', 'Техничар']).size().reset_index(name='Тековна')

# merge the counts into overdue_utre
overdue_utre = probieni_utre.merge(tekovni, on=['Име на регион', 'Техничар'], how='outer').fillna(0)
overdue_utre['Вкупно'] = overdue_utre['Пробиена утре 4PM'] + overdue_utre['Тековна']

# add more columns
columns_to_add_utre = [
    {"columnName": "Пробиена утре 4PM", "index": 2, "letterOfColumn": "C"},
    {"columnName": "Тековна", "index": 3, "letterOfColumn": "D"},
    {"columnName": "Пробиена денес", "index": 4, "letterOfColumn": "E"}
]

for column in columns_to_add_utre:
    if column['columnName'] not in overdue_utre.columns:
        overdue_utre[column['columnName']] = np.nan

# reorder columns
overdue_utre = overdue_utre[['Име на регион', 'Техничар', 'Пробиена утре 4PM', 'Тековна', 'Вкупно']]
print("Creatig Summery...")
total_nivo_3 = len(data_df)  # total number of rows in data_df
total_overdue_today = len(overdue_df)#overdue_df["Вкупно"].sum()  # total number of overdue issues today
total_overdue_tomorrow_4pm = overdue_utre['Вкупно'].sum()  # sum of Вкупно in overdue_utre
total_overdue_grupni_tomorrow_4pm = overdue_grupni['Вкупно'].sum()  # sum of Вкупно in overdue_grupni
summery_df = {
    'A': ["Вкупно отворени пречки на Ниво 3", 
          "Вкупно пречки чиј рок е пробиен денес", 
          "Вкупно пречки чиј рок ќе биде пробиен утре во 16 часот", 
          "Вкупно групни пречки чиј рок ќе биде пробиен утре во 16 часот"],
    'B':[total_nivo_3, total_overdue_today, total_overdue_tomorrow_4pm, total_overdue_grupni_tomorrow_4pm]
}

summery_df = pd.DataFrame(summery_df)

print("Creating Excel file...")
# create excel file
with pd.ExcelWriter("kreirani datoteki/" + excel_name + ".xlsx", engine='openpyxl') as writer:

    # write all the sheets
    pomoshen_df.to_excel(writer, sheet_name="Pomoshen", index=False)
    summery_df.to_excel(writer, sheet_name="Summery", index=False)
    overdue_utre.to_excel(writer, sheet_name="Overdue Utre", index=False)
    overdue_grupni.to_excel(writer, sheet_name="Overdue Grupni Utre", index=False)
    overdue_summary.to_excel(writer, sheet_name="Overdue Sega", index=False)
    data_df.to_excel(writer, sheet_name="Data", index=False)

    # get workbook and sheets
    workbook = writer.book
    pomoshen_sheet = writer.sheets["Pomoshen"]
    summery_sheet = writer.sheets["Summery"]
    overdue_utre_sheet = writer.sheets["Overdue Utre"]
    overdue_grupni_sheet = writer.sheets["Overdue Grupni Utre"]
    overdue_sheet = writer.sheets["Overdue Sega"]
    data_sheet = writer.sheets["Data"]
    
    # set column widths and height for all sheets
    column_widths = {
        "A": 15, "B": 15, "C": 30, "D": 20, "E": 33, "F": 12, "G": 25,
        "H": 20, "I": 20, "J": 18, "K": 15, "L": 15, "M": 15, "N": 20
    }
    overdue_column_widths = {
        "A": 20, "B": 20, "C": 10   
    }
    overdue_utre_widths = {
        "A": 20, "B": 20, "C": 20, "D": 20, "E": 20
    }
    summery_widths ={
        "A":58, "B":10
    }
    column_height = 20

    for col, width in column_widths.items():
        pomoshen_sheet.column_dimensions[col].width = width
        data_sheet.column_dimensions[col].width = width

    for row in range(1, pomoshen_sheet.max_row + 1):
        pomoshen_sheet.row_dimensions[row].height = column_height
        data_sheet.row_dimensions[row].height = column_height

    for col, width in overdue_column_widths.items():
        overdue_sheet.column_dimensions[col].width = width
        overdue_grupni_sheet.column_dimensions[col].width = width 
    for row in range(1, overdue_sheet.max_row + 1):
        overdue_sheet.row_dimensions[row].height = column_height
        overdue_grupni_sheet.row_dimensions[row].height = column_height

    for col, width in overdue_utre_widths.items():
        overdue_utre_sheet.column_dimensions[col].width = width
    for row in range(1, overdue_utre_sheet.max_row + 1):
        overdue_utre_sheet.row_dimensions[row].height = column_height


    for col, width in summery_widths.items():
        summery_sheet.column_dimensions[col].width = width
    for row in range(1, summery_sheet.max_row + 1):
        summery_sheet.row_dimensions[row].height = column_height
    
    pomoshen_sheet.sheet_state = 'hidden'

print("Excel file created successfully.")