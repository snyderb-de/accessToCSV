import pyodbc
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

def get_lookup_fields(connection, table_name):
    # Placeholder function - needs specific implementation
    return []

def construct_query(table_name, lookup_fields):
    # Placeholder function - needs specific implementation
    return f"SELECT * FROM {table_name}"

def export_to_files(connection, query, output_dir, table_name):
    df = pd.read_sql_query(query, connection)
    csv_output = os.path.join(output_dir, f"{table_name}.csv")
    xlsx_output = os.path.join(output_dir, f"{table_name}.xlsx")
    df.to_csv(csv_output, index=False)
    df.to_excel(xlsx_output, index=False)
    print(f"Exported {table_name} to CSV and XLSX in {output_dir}")

def process_database(db_path, base_output_dir):
    connection = pyodbc.connect(f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path}')
    cursor = connection.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = [row[0] for row in cursor.fetchall()]

    db_name = os.path.splitext(os.path.basename(db_path))[0]
    output_dir = os.path.join(base_output_dir, db_name)
    os.makedirs(output_dir, exist_ok=True)

    for table_name in tables:
        lookup_fields = get_lookup_fields(connection, table_name)
        query = construct_query(table_name, lookup_fields)
        export_to_files(connection, query, output_dir, table_name)

    connection.close()

def main():
    root = tk.Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory(title="Select Folder with Access Database Files")
    
    if not folder_selected:
        print("No folder selected. Exiting.")
        return

    converted_dir = os.path.join(folder_selected, 'converted')
    os.makedirs(converted_dir, exist_ok=True)

    for file in os.listdir(folder_selected):
        if file.endswith(".accdb") or file.endswith(".mdb"):
            db_path = os.path.join(folder_selected, file)
            process_database(db_path, converted_dir)

main()
