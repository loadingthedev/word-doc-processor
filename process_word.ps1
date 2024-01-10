$scriptDirectory = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$destinationFolder = Join-Path -Path $scriptDirectory -ChildPath "processed_doc"

Write-Host "Starting Word Document Processing"

# Create the destination folder if it doesn't exist
if (-not (Test-Path -Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder
    Write-Host "Created destination folder: $destinationFolder"
} else {
    Write-Host "Destination folder already exists: $destinationFolder"
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
Write-Host "Word application opened in background"

# Retrieve all the .docx files
$docxFiles = Get-ChildItem -Path $scriptDirectory -Filter *.docx

$fileCount = $docxFiles.Count
$processedCount = 0

if ($fileCount -eq 0) {
    Write-Host "No .docx files found in the directory: $scriptDirectory"
} else {
    Write-Host "Found $fileCount .docx file(s) to process"

    foreach ($file in $docxFiles) {
        $docPath = $file.FullName
        $newDocPath = Join-Path -Path $destinationFolder -ChildPath $file.Name

        Write-Host "Processing file: $($file.Name)"
        $document = $word.Documents.Open($docPath)
        $document.SaveAs([ref] [string] $newDocPath, [ref] 16) # 16 is the WdSaveFormat for Word Document
        $document.Close()
        Write-Host "Saved processed file to: $newDocPath"

        $processedCount++
        Write-Host "Processed ${processedCount} of ${fileCount}: $($file.Name)"
    }

    $word.Quit()
    Write-Host "Word application closed"
}

Write-Host "Processing complete. Total files processed: ${processedCount}"
