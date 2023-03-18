#!/usr/bin/env bash

# uses MDB Tools to loop through a directory of MS Access files and creates
# a folder for each Access file
# and a CSV for each table in each file

#!/usr/bin/env bash

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
  if [ ! -f "$fullfilename" ] || ! command -v mdb-tables >/dev/null 2>&1 || ! mdb-tables -1 "$fullfilename" >/dev/null 2>&1; then
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

