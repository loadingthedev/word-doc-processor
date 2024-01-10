$scriptDirectory = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$destinationFolder = Join-Path -Path $scriptDirectory -ChildPath "processed_doc"

# Create the destination folder if it doesn't exist
if (-not (Test-Path -Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Retrieve all the .docx files
$docxFiles = Get-ChildItem -Path $scriptDirectory -Filter *.docx

$fileCount = $docxFiles.Count
$processedCount = 0

foreach ($file in $docxFiles) {
    $docPath = $file.FullName
    $newDocPath = Join-Path -Path $destinationFolder -ChildPath $file.Name

    $document = $word.Documents.Open($docPath)
    $document.SaveAs([ref] [string] $newDocPath, [ref] 16) # 16 is the WdSaveFormat for Word Document
    $document.Close()

    $processedCount++
    Write-Host "Processed ${processedCount} of ${fileCount}: $($file.Name)"
}

$word.Quit()

Write-Host "Processing complete. Total files processed: ${processedCount}"
