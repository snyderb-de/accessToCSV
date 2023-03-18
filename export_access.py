#!/usr/bin/env python3

import os
import sys
import argparse
import subprocess
from tkinter import Tk, filedialog

def install_dependencies():
    if sys.platform.startswith("linux"):
        package_managers = ["apt-get", "yum", "pacman", "dnf"]
        for pm in package_managers:
            if subprocess.run(["which", pm], capture_output=True).returncode == 0:
                subprocess.run(["sudo", pm, "install", "mdbtools"])
                break
        else:
            print("Package manager not found. Please install the required packages manually.")
            sys.exit(1)

    elif sys.platform == "darwin":
        subprocess.run(["brew", "install", "mdbtools"])

    else:
        print("Unsupported platform. Please install the required packages manually.")
        sys.exit(1)

def is_mdbtools_installed():
    return subprocess.run(["which", "mdb-export"], capture_output=True).returncode == 0

def export_tables_to_csv(folder):
    for filename in os.listdir(folder):
        if not (filename.endswith(".accdb") or filename.endswith(".mdb")):
            continue

        fullfilename = os.path.join(folder, filename)
        if not os.path.isfile(fullfilename):
            continue

        dbname, _ = os.path.splitext(filename)
        db_output_dir = os.path.join(args.output, dbname)
        os.makedirs(db_output_dir, exist_ok=True)

        table_names = subprocess.run(["mdb-tables", "-1", fullfilename], capture_output=True, text=True)
        if table_names.returncode != 0:
            print(f"Skipping invalid file: {fullfilename}")
            continue

        for table in table_names.stdout.splitlines():
            print(f"Exporting table {table} from {filename}")
            csv_file = os.path.join(db_output_dir, f"{table}.csv")
            with open(csv_file, "w") as csvfile:
                subprocess.run(["mdb-export", fullfilename, table], stdout=csvfile)

def main(args):
    if not is_mdbtools_installed():
        print("The mdbtools package is required but is not installed.")
        choice = input("Do you want to install it now? [y/N] ").lower()
        if choice.startswith("y"):
            install_dependencies()
        else:
            print("Aborting.")
            sys.exit(1)

    if args.folder is None:
        root = Tk()
        root.withdraw()
        folder = filedialog.askdirectory(title="Select Folder with Access Database Files")
        if not folder:
            print("User cancelled. Exiting.")
            sys.exit(1)
    else:
        folder = args.folder

    export_tables_to_csv(folder)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Export Access database tables to CSV")
    parser.add_argument("-f", "--folder", help="Specify the folder containing Access files to export.")
    parser.add_argument("-o", "--output", default="exported_csv", help="Specify the output folder to store the exported CSV files. Defaults to a folder named 'exported_csv' in the current working directory.")
    args = parser.parse_args()

    main(args)
