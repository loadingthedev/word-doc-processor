# The directory where the script is located
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# The directory where the processed files will be saved
$ProcessedFolder = Join-Path $ScriptDir "processeddoc"

# Create the processed folder if it doesn't exist
if (-not (Test-Path -Path $ProcessedFolder)) {
    New-Item -ItemType Directory -Path $ProcessedFolder
}

# Get all .docx files in the script's directory
Get-ChildItem -Path $ScriptDir -Filter "*.docx" | ForEach-Object {
    # Original file name
    $FileName = $_.Name

    # Remove leading and trailing underscores
    $FileName = $FileName.TrimStart('_').TrimEnd('_')

    # Replace hyphens with spaces
    $NewFileName = $FileName -replace '-', ' '

    # Full path for the new file
    $NewFilePath = Join-Path $ProcessedFolder $NewFileName

    # Copy the file with the new name to the processed folder
    Copy-Item -Path $_.FullName -Destination $NewFilePath
}

Write-Host "Process complete."
