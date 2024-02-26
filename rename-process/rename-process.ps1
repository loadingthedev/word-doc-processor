# The directory where the script is located
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition

# The directory where the processed files will be saved
$destinationFolder = Join-Path $scriptDirectory "processed_doc"

# Create the destination folder if it doesn't exist
if (-not (Test-Path -Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder
    Write-Host "Created destination folder: $destinationFolder"
} else {
    Write-Host "Destination folder already exists: $destinationFolder"
}

# Initialize Word application
$word = New-Object -ComObject Word.Application
$word.Visible = $false
Write-Host "Word application opened in background"

# Get all .docx files in the script's directory
$docxFiles = Get-ChildItem -Path $scriptDirectory -Filter "*.docx"

$fileCount = $docxFiles.Count
$processedCount = 0

if ($fileCount -eq 0) {
    Write-Host "No .docx files found in the directory: $scriptDirectory"
} else {
    Write-Host "Found $fileCount .docx file(s) to process"

    foreach ($file in $docxFiles) {
        # Original file name
        $fileName = $file.Name

        # Log the file name
        Write-Host "Processing file: $fileName"

        # Open the document
        $document = $word.Documents.Open($file.FullName)

        # Remove leading and trailing underscores
        $fileName = $fileName.TrimStart('_').TrimEnd('_')

        # Replace hyphens with spaces and _ with spaces
        $newFileName = $fileName -replace '-', ' '
        $newFileName = $newFileName -replace '_', ' '

        # Trim leading and trailing spaces
        $newFileName = $newFileName.Trim()

        # Full path for the new file
        $newFilePath = Join-Path $destinationFolder $newFileName

        # Save the document with the new name
        $document.SaveAs([ref] [string] $newFilePath, [ref] 16) # 16 is the WdSaveFormat for Word Document
        $document.Close()

        Write-Host "Saved processed file to: $newFilePath"

        $processedCount++
        Write-Host "Processed ${processedCount} of ${fileCount}: $newFileName"
    }

    $word.Quit()
    Write-Host "Word application closed"
}

Write-Host "Processing complete. Total files processed: ${processedCount}"
