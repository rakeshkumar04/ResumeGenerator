#!/bin/bash

# Directory containing the Excel files
EXCEL_DIR="excel_files"  # Replace with the actual directory containing your Excel files

# Check if the directory exists
if [ ! -d "$EXCEL_DIR" ]; then
  echo "Directory $EXCEL_DIR does not exist."
  exit 1
fi

# Iterate over each Excel file in the directory
for excel_file in "$EXCEL_DIR"/*.xlsx; do
  # Check if the file exists (in case there are no .xlsx files)
  if [ ! -f "$excel_file" ]; then
    echo "No Excel files found in $EXCEL_DIR."
    exit 1
  fi

  # Run the main.py script for each Excel file
  echo "Generating tex files for $excel_file..."
  python3 main.py "$excel_file"

  # Check if the main.py script executed successfully
  if [ $? -ne 0 ]; then
    echo "Error occurred while processing $excel_file. Aborting."
    exit 1
  fi
done

echo "All tex files generated successfully."
