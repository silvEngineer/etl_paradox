import pyodbc
import csv
import os
import json

def export_table_to_csv(connection, table_name, query, output_dir):
    cursor = connection.cursor()
    
    cursor.execute(query)
    rows = cursor.fetchall()

    columns = [column[0] for column in cursor.description]

    output_file = os.path.join(output_dir, f"{table_name}.csv")

    with open(output_file, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        
        writer.writerow(columns)

        writer.writerows(rows)

    print(f"Tabela '{table_name}' exportada para {output_file}.")

LOCATION = r"C:\USUARIO\BDE"
cnxn = pyodbc.connect(
    r"Driver={{Microsoft Paradox Driver (*.db )\}};DriverID=538;Fil=Paradox 5.X;DefaultDir={0};Dbq={0};CollatingSequence=ASCII;".format(LOCATION),
    autocommit=True
)

config_file = r"C:\Users\Usuario\Desktop\paradox_extrator\config.json"
with open(config_file, "r", encoding="utf-8") as file:
    tables = json.load(file)

output_directory = r"C:\Users\Usuario\Desktop\paradox_extrator\dados"

os.makedirs(output_directory, exist_ok=True)

for table in tables:
    export_table_to_csv(cnxn, table["tabela"], table["query"], output_directory)