#!/usr/bin/env bash

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
  command -v mdb-export >/dev/null 2>&1
}

function export_tables_to_csv() {
  local folder="$1"
  log_file="$output/error_log.txt"
  for fullfilename in "$folder"/*.{accdb,mdb}; do
    if [ ! -f "$fullfilename" ] || ! mdb-tables -1 "$fullfilename" >/dev/null 2>&1; then
      echo "Skipping invalid file: $fullfilename" >> "$log_file"
      continue
    fi

    filename=$(basename "$fullfilename")
    dbname=${filename%.*}
    output_dir="$output/$dbname"
    mkdir -p "$output_dir"

    IFS=$'\n'
    for table in $(mdb-tables -1 "$fullfilename"); do
      echo "Exporting table $table from $filename" >> "$log_file"
      mdb-export "$fullfilename" "$table" > "$output_dir/$table.csv"
    done
  done
}


if ! is_mdbtools_installed; then
  echo "The mdbtools package is required but is not installed."
  read -rp "Do you want to install it now? [y/N] " answer
  if [[ "$answer" =~ [yY](es)* ]]; then
    install_dependencies
  else
    echo "Aborting."
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

export_tables_to_csv "$folder"
