#!/usr/bin/env bash

#!/usr/bin/env bash

: <<'COMMENT'
This script uses MDB Tools to loop through a directory of MS Access files and creates
a folder for each Access file and a CSV for each table in each file.

Dependencies:
- mdbtools
- zenity
- awk
- sed

Usage: ./export_access.sh [OPTIONS]

Options:
  -f, --folder     Specify the folder containing Access files to export.
                   Defaults to the current working directory.

  -o, --output     Specify the output folder to store the exported CSV files.
                   Defaults to a folder named "exported_csv" in the current working directory.

  -s, --single     Specify a single Access file to export. The folder option will be ignored.

  -h, --help       Show this help message and exit.
COMMENT



# Detect package manager
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

# Check if mdbtools package is installed
if ! dpkg -s mdbtools >/dev/null 2>&1 && ! rpm -q mdbtools >/dev/null 2>&1 && ! pacman -Q mdbtools >/dev/null 2>&1 && ! dnf list installed mdbtools >/dev/null 2>&1 && ! brew ls --versions mdbtools >/dev/null 2>&1; then
  echo "The mdbtools package is required but is not installed."
  read -rp "Do you want to install it now? [y/N] " answer
  if [[ "$answer" =~ [yY](es)* ]]; then
    case $PKG_MANAGER in
      apt-get) sudo apt-get update && sudo apt-get install mdbtools ;;
      yum) sudo yum install mdbtools ;;
      pacman) sudo pacman -S mdbtools ;;
      dnf) sudo dnf install mdbtools ;;
      brew) brew install mdbtools ;;
      *)
        echo "Package manager not found. Please install the required packages manually."
        exit 1
        ;;
    esac
  else
    echo "Aborting."
    exit 1
  fi
fi

# Check if required commands are installed
missing_commands=()
for cmd in awk sed zenity; do
  if ! command -v "$cmd" >/dev/null 2>&1; then
    missing_commands+=("$cmd")
  fi
done

# If any required commands are missing, prompt the user to install them
if [ ${#missing_commands[@]} -gt 0 ]; then
  echo "The following commands are required but are not installed: ${missing_commands[*]}"
  read -rp "Do you want to install them now? [y/N] " answer
  if [[ "$answer" =~ [yY](es)* ]]; then
    case $PKG_MANAGER in
      apt-get) sudo apt-get update && sudo apt-get install "${missing_commands[@]}" ;;
      yum) sudo yum install "${missing_commands[@]}" ;;
      pacman) sudo pacman -S "${missing_commands[@]}" ;;
      dnf) sudo dnf install "${missing_commands[@]}" ;;
      brew) brew install "${missing_commands[@]}" ;;
      *)
        echo "Package manager not found. Please install the required packages manually."
        exit 1
        ;;
    esac
  else
    echo "Aborting."
    exit 1
  fi
fi


# Open a file chooser dialog to let the user select the folder containing the Access database files
folder=$(zenity --file-selection --directory --title="Select Folder with Access Database Files")

# Check if the user clicked the "Cancel" button in the file chooser dialog
if [ -z "$folder" ]; then
  echo "User cancelled. Exiting."
  exit 1
fi

# Loop through each file in the selected folder and export all tables to CSV format
for fullfilename in "$folder"/*.{accdb,mdb}; do
  # Check if the file exists and is a valid Access database file
  if [ ! -f "$fullfilename" ] || ! mdb-tables -1 "$fullfilename" >/dev/null 2>&1; then
    echo "Skipping invalid file: $fullfilename"
    continue
  fi

  # Extract the filename and database name from the full path to the database file
  filename=$(basename "$fullfilename")
  dbname=${filename%.*}

  # Create a new directory with the database name
  mkdir "$dbname"

  # Loop through each table in the database and export to CSV format
  IFS=$'\n'
  for table in $(mdb-tables -1 "$fullfilename"); do
    echo "Exporting table $table from $filename"
    mdb-export "$fullfilename" "$table" > "$dbname/$table.csv"
  done
done

