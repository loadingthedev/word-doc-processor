#!/bin/bash

# Get the directory where the script is located
ScriptDir="$(dirname "$(realpath "$0")")"

# Define the directory where the processed files will be saved
ProcessedFolder="$ScriptDir/processeddoc"

# Create the processed folder if it doesn't exist
if [[ ! -d "$ProcessedFolder" ]]; then
    mkdir "$ProcessedFolder"
fi

# Get all .docx files in the script's directory
for file in "$ScriptDir"/*.docx; do
    if [[ -f "$file" ]]; then
        # Extract the file name
        FileName=$(basename "$file")

        # Remove leading and trailing underscores
        FileName=${FileName##*_}
        FileName=${FileName%%_*}

        # Replace hyphens with spaces
        NewFileName=${FileName//-/ }

        # Full path for the new file
        NewFilePath="$ProcessedFolder/$NewFileName"

        # Copy the file with the new name to the processed folder
        cp "$file" "$NewFilePath"
    fi
done

echo "Process complete."
