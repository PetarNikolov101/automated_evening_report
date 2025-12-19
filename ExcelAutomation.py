import pandas as pd
from datetime import date, timedelta, datetime
import numpy as np
import os

#pivot  DataFrame.pivot(index=None, columns=None, values=None)

# read from html
file_path = "otvoreniprecki"

# dfs = pd.read_html(file_path + ".xls")
# df = dfs[0]
# df.to_excel('kreirani datoteki/converted_neotstraneti.xlsx', index=False)
# copy_from = pd.read_excel("kreirani datoteki/converted_neotstraneti.xlsx", index_col=None)

df = pd.read_excel(r"C:\Users\petarnik\skripta_neotstraneti\skripta_neotstraneti\otvoreniprecki.xlsx", header = 1, engine='openpyxl')
copy_from = pd.read_excel(r"C:\Users\petarnik\skripta_neotstraneti\skripta_neotstraneti\otvoreniprecki.xlsx", header=1, engine='openpyxl')

today = datetime.today()
tomorrow = date.today() + timedelta(days=1)
after_tomorrow = date.today() + timedelta(days=2)
tomorrow_4pm = datetime.combine(tomorrow, datetime.strptime("16:00:00", "%H:%M:%S").time())
date_str = today.strftime("%d-%m-%Y")  # safe for filenames

excel_name = f"Lista na precki - TT {date_str}.xlsx"

print("Extracting columns...")
# columns to extract
columns_to_extract = [
    {"columnName": "Пречка", "index": 0, "letterOfColumn": "A"},
    {"columnName": "LineID", "index": 1, "letterOfColumn": "B"},
    {"columnName": "Статус", "index": 2, "letterOfColumn": "C"},
    {"columnName": "Дата на пријава", "index": 4, "letterOfColumn": "E"},
    {"columnName": "Категорија", "index": 9, "letterOfColumn": "J"},
    {"columnName": "Групна", "index": 12, "letterOfColumn": "M"},
    {"columnName": "Групна Kласификација", "index": 13, "letterOfColumn": "N"},
    {"columnName": "Име на регион", "index": 16, "letterOfColumn": "Q"},
    {"columnName": "Last remark WFM", "index": 17, "letterOfColumn": "R"},
    {"columnName": "Техничар", "index": 18, "letterOfColumn": "S"},
    {"columnName": "Статус налог", "index": 19, "letterOfColumn": "T"},
    {"columnName": "Доделена група", "index": 20, "letterOfColumn": "U"},
    {"columnName": "Посакуван крај", "index": 21, "letterOfColumn": "V"}
]

global columns_to_extract_for_csod 
columns_to_extract_for_csod=[
    {"columnName": "Класификација", "index": 17, "letterOfColumn": "P"}
]

# extract columns
selected_columns = {
    col["columnName"]: copy_from.iloc[:, col["index"]]
    for col in columns_to_extract
}

selected_df = pd.DataFrame(selected_columns)

# convert selected_columns back to a dictionary
selected_columns = selected_df.to_dict('list')

pomoshen_df = pd.DataFrame(selected_columns)
# convert Посакуван крај to datetime
pomoshen_df['Посакуван крај'] = pd.to_datetime(pomoshen_df['Посакуван крај'], errors='coerce', dayfirst=True)

# add status column
def categorize_status(deadline):
    if pd.isna(deadline):
        return "unknown"
    if deadline <= today:
        return "overdue"
    elif deadline.date() == tomorrow and deadline <= tomorrow_4pm:
        return "will be overdue tomorrow at 4"
    else:
        return "not overdue"

pomoshen_df['Status'] = pomoshen_df['Посакуван крај'].apply(categorize_status)

print("Filtering overdue issues...")

# filter overdue for SSOD
data_df = pomoshen_df[(pomoshen_df['Доделена група'] == "SSOD")&(pomoshen_df["Статус налог"] != "ОТКАЖАН")]
overdue_df = data_df[(data_df['Посакуван крај'].dt.date <= today.date())]

print("Grouping by region, technician, and counting their overdue issues...")
# group by region and technician, count overdue issues
overdue_summary = overdue_df.groupby(['Име на регион', 'Техничар']).size().reset_index(name='Вкупно')

# group issues for Overdue Grupni Utre
overdue_grupni_df = pomoshen_df[
    (pomoshen_df['Групна'].notna()) & (pomoshen_df['Посакуван крај'] <= tomorrow_4pm)
]
overdue_grupni = overdue_grupni_df.groupby(['Име на регион', 'Групна Kласификација']).size().reset_index(name='Вкупно')

# Filter for non-group tasks with a valid 'Посакуван крај'
overdue_utre_df = data_df[
    (data_df['Групна'].isna()) & (data_df['Посакуван крај'].notna())
]

# Пробиена утре 4PM (will become overdue by tomorrow 4 PM)
# probieni_utre = overdue_utre_df[
#     overdue_utre_df['Посакуван крај'] < tomorrow_4pm,
# ].groupby(['Име на регион', 'Техничар']).size().reset_index(name='Пробиена утре 4PM')

probieni_utre = overdue_utre_df[
    (overdue_utre_df['Посакуван крај'] > today) &
    (overdue_utre_df['Посакуван крај'] <= tomorrow_4pm)
].groupby(['Име на регион', 'Техничар']).size().reset_index(name='Пробиена утре 4PM')

probieni_utre['Вкупно'] = probieni_utre['Пробиена утре 4PM']

# Тековна (still valid after tomorrow 4 PM)
tekovni = overdue_utre_df[
    overdue_utre_df['Посакуван крај'] > tomorrow_4pm
].groupby(['Име на регион', 'Техничар']).size().reset_index(name='Тековна')

# Веќе пробиена (already overdue before today)
veke_probieni = overdue_utre_df[
    overdue_utre_df['Посакуван крај'] < today
].groupby(['Име на регион', 'Техничар']).size().reset_index(name='Веќе пробиена')

# Merge all three dataframes together
overdue_utre = probieni_utre.merge(tekovni, on=['Име на регион', 'Техничар'], how='outer')
overdue_utre = overdue_utre.merge(veke_probieni, on=['Име на регион', 'Техничар'], how='outer')
overdue_utre = overdue_utre.fillna(0)

# Calculate total
overdue_utre['Вкупно'] = (
    overdue_utre['Пробиена утре 4PM']
    + overdue_utre['Тековна']
    + overdue_utre['Веќе пробиена']
)

# add more columns
columns_to_add_utre = [
    {"columnName": "Веќе пробиена", "index": 2, "letterOfColumn": "C"},
    {"columnName": "Пробиена утре 4PM", "index": 2, "letterOfColumn": "D"},
    {"columnName": "Тековна", "index": 4, "letterOfColumn": "E"},
    {"columnName": "Пробиена денес", "index": 5, "letterOfColumn": "F"}
]

for column in columns_to_add_utre:
    if column['columnName'] not in overdue_utre.columns:
        overdue_utre[column['columnName']] = np.nan

# reorder columns
overdue_utre = overdue_utre[['Име на регион', 'Техничар', 'Веќе пробиена', 'Пробиена утре 4PM', 'Тековна', 'Вкупно']]
print("Creatig Summery...")
total_nivo_3 = len(data_df)  # total number of rows in data_df
total_overdue_today = len(overdue_df)#overdue_df["Вкупно"].sum()  # total number of overdue issues today
total_overdue_tomorrow_4pm = probieni_utre['Пробиена утре 4PM'].sum()  # sum of will be overdue tomorrow
total_overdue_grupni_tomorrow_4pm = overdue_grupni['Вкупно'].sum()  # sum of Вкупно in overdue_grupni
whole = total_overdue_grupni_tomorrow_4pm + total_overdue_tomorrow_4pm
percentile = (total_overdue_tomorrow_4pm/whole)*100
rounded_percentile = round(percentile, 2)
percentile_string = f"{rounded_percentile}%"
summery_df = {
    'A': ["Вкупно отворени пречки на Ниво 3", 
          "Вкупно пречки чиј рок е пробиен денес", 
          "Вкупно пречки чиј рок ќе биде пробиен утре во 16 часот", 
          "Вкупно групни пречки чиј рок ќе биде пробиен утре во 16 часот",
          "Процент на единечни пречки кои не се поврзани со групен прекин од вкупниот број на пречки чиј рок ќе биде пробиен утре во 16 часот"],
    'B':[total_nivo_3, total_overdue_today, total_overdue_tomorrow_4pm, total_overdue_grupni_tomorrow_4pm, percentile_string]
}

summery_df = pd.DataFrame(summery_df)

def merge_region_cells(sheet, region_col_letter='A'):
    """
    Merge vertically adjacent cells in the 'region' column
    if they have the same value (e.g., Bitola appearing multiple times).
    """
    max_row = sheet.max_row
    start_row = 2  # Assuming row 1 is headers

    while start_row <= max_row:
        region_value = sheet[f"{region_col_letter}{start_row}"].value
        end_row = start_row

        # Find the last consecutive row with the same region
        while end_row + 1 <= max_row and sheet[f"{region_col_letter}{end_row + 1}"].value == region_value:
            end_row += 1

        # Merge cells if more than one row has the same region
        if end_row > start_row:
            sheet.merge_cells(f"{region_col_letter}{start_row}:{region_col_letter}{end_row}")

        start_row = end_row + 1

def add_totals(df, region_col='Име на регион', tech_col='Техничар'):
    # Make a copy to avoid modifying original
    df = df.copy()

    # Identify numeric columns
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()

    # Prepare new list to store rows (including region totals)
    new_rows = []

    # Group by region and add subtotal rows
    for region, group in df.groupby(region_col, sort=False):
        new_rows.append(group)
        subtotal = group[numeric_cols + ['Вкупно']].sum()
        subtotal_row = {
            region_col: f"{region} - Вкупно",
            tech_col: "",
            **subtotal.to_dict()
        }
        new_rows.append(pd.DataFrame([subtotal_row]))

    # Concatenate all rows
    df_with_totals = pd.concat(new_rows, ignore_index=True)

    # Add a final grand total row
    grand_total = df_with_totals[numeric_cols + ['Вкупно']].sum()/2
    grand_total_row = {
        region_col: "ВКУПНО",
        tech_col: "",
        **grand_total.to_dict()
    }
    df_with_totals = pd.concat([df_with_totals, pd.DataFrame([grand_total_row])], ignore_index=True)

    return df_with_totals

#namesto po region po klasifikacija
#selected_columns = {
#    col["columnName"]: copy_from.iloc[:, col["index"]]
#    for col in columns_to_extract
#}
def get_CSOD():
    df_subset = df.copy()

    csod_table = df_subset[
        (df_subset['Доделена група'] == "CSOD") &
        (df_subset["Статус налог"] != "ОТКАЖАН") &
        (df_subset["Групна"].isna())
    ]

    selected_cols = [col["columnName"] for col in columns_to_extract_for_csod]
    csod_table = csod_table[selected_cols].copy()

    global csod_summary
    csod_summary = (
        csod_table["Класификација"]
        .value_counts()
        .reset_index()
    )
    csod_summary.columns = ["Класификација", "Вкупно"]

    # Ensure Вкупно is int
    csod_summary["Вкупно"] = csod_summary["Вкупно"].astype(int)

    # Add total row
    total_value = csod_summary["Вкупно"].sum()
    total_row = pd.DataFrame({
        "Класификација": ["ВКУПНО"],
        "Вкупно": [total_value]
    })

    # Concatenate without float upcasting
    csod_summary = pd.concat([csod_summary, total_row], ignore_index=True)
    


# Apply to all three tables
overdue_utre = add_totals(overdue_utre)
overdue_summary = add_totals(overdue_summary)
overdue_grupni = add_totals(overdue_grupni, tech_col = None)       

get_CSOD()

print("Creating Excel file...")
# create excel file
with pd.ExcelWriter(os.path.join(r"C:\Users\petarnik\skripta_neotstraneti\skripta_neotstraneti\kreirani datoteki", excel_name), engine='openpyxl') as writer:

    # write all the sheets
    pomoshen_df.to_excel(writer, sheet_name="Pomoshen", index=False)
    summery_df.to_excel(writer, sheet_name="Summery", index=False)
    overdue_utre.to_excel(writer, sheet_name="Overdue Utre", index=False)
    overdue_grupni.to_excel(writer, sheet_name="Overdue Grupni Utre", index=False)
    overdue_summary.to_excel(writer, sheet_name="Overdue Sega", index=False)
    data_df.to_excel(writer, sheet_name="Data", index=False)
    csod_summary.to_excel(writer, sheet_name = "CSOD", index=False)


    # get workbook and sheets
    workbook = writer.book
    pomoshen_sheet = writer.sheets["Pomoshen"]
    summery_sheet = writer.sheets["Summery"]
    overdue_utre_sheet = writer.sheets["Overdue Utre"]
    overdue_grupni_sheet = writer.sheets["Overdue Grupni Utre"]
    overdue_sheet = writer.sheets["Overdue Sega"]
    data_sheet = writer.sheets["Data"]
    hidden_csod = writer.sheets["CSOD"]
    
    # set column widths and height for all sheets
    column_widths = {
        "A": 15, "B": 15, "C": 30, "D": 20, "E": 33, "F": 12, "G": 25,
        "H": 20, "I": 20, "J": 18, "K": 15, "L": 15, "M": 20, "N": 20
    }
    overdue_column_widths = {
        "A": 20, "B": 20, "C": 10   
    }
    overdue_utre_widths = {
        "A": 20, "B": 20, "C": 20, "D": 20, "E": 20
    }
    summery_widths ={
        "A":120, "B":10
    }
    csod_widths = {
        "A":30, "B":20
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

    for col, width in csod_widths.items():
        hidden_csod.column_dimensions[col].width = width
    for row in range(1, hidden_csod.max_row + 1):
        hidden_csod.row_dimensions[row].height = column_height

    for col, width in summery_widths.items():
        summery_sheet.column_dimensions[col].width = width
    for row in range(1, summery_sheet.max_row + 1):
        summery_sheet.row_dimensions[row].height = column_height
    
    merge_region_cells(overdue_utre_sheet, region_col_letter='A')
    merge_region_cells(overdue_sheet, region_col_letter='A')
    merge_region_cells(overdue_grupni_sheet, region_col_letter='A')

    pomoshen_sheet.sheet_state = 'hidden'
    csod_summary.sheet_state = 'hidden'
    overdue_sheet.sheet_state = "hidden"

print("Excel file created successfully.")