#!/usr/bin/env bash

function check_python_and_tkinter() {

  echo "Checking for Python 3 and Tkinter..."

  if ! command -v python3 >/dev/null 2>&1; then
    echo "Python 3 is required but not installed on your system."
    echo "Please visit https://www.python.org/downloads/ to download and install Python 3."
    exit 1
  fi

  if ! python3 -c "import tkinter" >/dev/null 2>&1; then
    echo "Tkinter is required but not installed on your system."
    echo "Please install Tkinter for Python 3. You can learn more at https://tkdocs.com/tutorial/install.html"
    exit 1
  fi

  echo "Python 3 and Tkinter are installed. Proceeding with the script..."
  sleep 3

}

check_python_and_tkinter

function install_dependencies() {
  if command -v apt-get >/dev/null 2>&1; then
    PKG_MANAGER="apt-get"
  elif command -v yum >/dev/null 2>&1; then
    PKG_MANAGER="yum"
  elif command -v pacman >/dev/null 2>&1; then
    PKG_MANAGER="pacman"
  elif command -v dnf >/dev/null 2>&1; then
    PKG_MANAGER="dnf"
  elif command -v brew >/dev/null 2>&1; then
    PKG_MANAGER="brew"
  else
    echo "Package manager not found. Please install the required packages manually."
    exit 1
  fi

  case $PKG_MANAGER in
    apt-get)
      sudo apt-get update
      sudo apt-get install mdbtools
      ;;
    yum)
      sudo yum install mdbtools
      ;;
    pacman)
      sudo pacman -S mdbtools
      ;;
    dnf)
      sudo dnf install mdbtools
      ;;
    brew)
      brew install mdbtools
      ;;
  esac
}

function is_mdbtools_installed() {
  echo "Checking for mdbtols..."
  command -v mdb-export >/dev/null 2>&1
}

function export_tables_to_csv() {
  local folder="$1"
  local total_files="$2"
  local processed_files=0
  log_file="$output/error_log.txt"

  # Ensure the log file is created
  touch "$log_file"

  for fullfilename in "$folder"/*; do
    extension="${fullfilename##*.}"
    if [[ "$extension" != "accdb" && "$extension" != "mdb" && "$extension" != "adp" ]]; then
      echo "Skipped non-Access file: $fullfilename" >> "$log_file"
      continue
    fi

    if [ ! -f "$fullfilename" ] || ! mdb-tables -1 "$fullfilename" >/dev/null 2>&1; then
      echo "Skipped invalid file: $fullfilename" >> "$log_file"
      continue
    fi

    filename=$(basename "$fullfilename")
    dbname=${filename%.*}
    output_dir="$output/$dbname"
    mkdir -p "$output_dir"

    IFS=$'\n'
    for table in $(mdb-tables -1 "$fullfilename"); do
      echo "Exporting table $table from $filename"
      csv_output="$output_dir/$table.csv"
      xlsx_output="$output_dir/$table.xlsx"
      mdb-export "$fullfilename" "$table" > "$csv_output"
      # Convert CSV to Excel with pandas
      python3 -c "import pandas as pd; df = pd.read_csv('$csv_output'); df.to_excel('$xlsx_output', index=False)"
    done
    # Increment the processed files count and print the progress
    ((processed_files++))
    echo -ne "Processing files: $processed_files/$total_files\r"
  done
  echo
}


if ! is_mdbtools_installed; then
  echo "The mdbtools package is required but is not installed."
  read -rp "Do you want to install it now? [y/N] " answer
  if [[ "$answer" =~ [yY](es)* ]]; then
    install_dependencies
  else
    echo "No? ... alright, then."
    exit 1
  fi
fi

if [[ "$(uname)" == "Darwin" ]]; then
  folder=$(osascript -e 'try' -e 'set _folder to choose folder with prompt "Select Folder with Access Database Files"' -e 'POSIX path of _folder' -e 'on error' -e '' -e 'end try')
  output=$(osascript -e 'try' -e 'set _folder to choose folder with prompt "Select Output Folder"' -e 'POSIX path of _folder' -e 'on error' -e '' -e 'end try')
else
  folder=$(python3 -c 'import tkinter as tk; from tkinter import filedialog; root = tk.Tk(); root.withdraw(); folder = filedialog.askdirectory(title="Select Folder with Access Database Files"); print(folder)' 2>/dev/null)
  output=$(python3 -c 'import tkinter as tk; from tkinter import filedialog; root = tk.Tk(); root.withdraw(); folder = filedialog.askdirectory(title="Select Output Folder"); print(folder)' 2>/dev/null)
fi

if [ -z "$folder" ] || [ -z "$output" ]; then
  echo "User cancelled. Exiting."
  exit 1
fi


total_files=$(find "$folder" -type f \( -iname "*.accdb" -o -iname "*.mdb" -o -iname "*.adp" \) | wc -l)

export_tables_to_csv "$folder" "$total_files"
