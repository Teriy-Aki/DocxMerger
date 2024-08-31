# DocxMerger - AScript to merge selected .docx files using Word Interop

# Error handling: trap any errors and keep the console open
trap {
    Write-Error "An error occurred: $_"
    Read-Host "Press Enter to exit..."
    exit 1
}

# Get the files passed as arguments (from the "Send To" context)
$selectedFiles = $args | Where-Object { $_ -match "\.docx$" }

if ($selectedFiles.Count -eq 0) {
    Write-Error "No .docx files were selected."
    Read-Host "Press Enter to exit..."
    exit 1
}

# Prompt user for sorting option
Write-Host "Choose how to sort the files before merging:"
Write-Host "1 - By Name"
Write-Host "2 - By Date Created"
Write-Host "3 - By Date Modified"
$sortOption = Read-Host "Enter your choice (1/2/3)"

# Sort the selected files based on user input
switch ($sortOption) {
    '1' { $docxFiles = $selectedFiles | Sort-Object { (Get-Item $_).Name } }
    '2' { $docxFiles = $selectedFiles | Sort-Object { (Get-Item $_).CreationTime } }
    '3' { $docxFiles = $selectedFiles | Sort-Object { (Get-Item $_).LastWriteTime } }
    default {
        Write-Error "Invalid selection. Please enter 1, 2, or 3."
        Read-Host "Press Enter to exit..."
        exit 1
    }
}

# Specify the temporary output file name (on the desktop)
$tempFilePath = "USERPATH_HERE\MergedOutput.docx"

# Final desired location for the merged document in the same directory as the first selected file
$directory = (Get-Item $docxFiles[0]).DirectoryName
$finalOutputFile = Join-Path $directory "MergedOutput.docx"

# Initialize Word Application
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    # Create a new document for the merged output
    Write-Output "Creating new merged document..."
    $mergedDoc = $word.Documents.Add()

    # Remove any initial empty paragraph
    $mergedDoc.Content.Text = ""

    $isFirstDocument = $true

    foreach ($file in $docxFiles) {
        Write-Output "Processing file: '$($file)'"

        # Open each document
        $doc = $word.Documents.Open($file)
        Write-Output "Opened document: '$($file)'"

        # If this is not the first document, insert a page break before adding new content
        if (-not $isFirstDocument) {
            $endRange = $mergedDoc.Content
            $endRange.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
            $endRange.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdPageBreak)
            Write-Output "Inserted page break."
        }

        # Select the entire content of the document
        $range = $doc.Content
        $range.Copy()
        Write-Output "Copied content from: '$($file)'"

        # Move to the end of the merged document and paste the content
        $endRange = $mergedDoc.Content
        $endRange.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)
        $endRange.PasteAndFormat([Microsoft.Office.Interop.Word.WdRecoveryType]::wdFormatOriginalFormatting)
        Write-Output "Pasted content into merged document."

        # Close the document without saving
        $doc.Close([ref]$false)
        Write-Output "Closed document: '$($file)'"
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null

        $isFirstDocument = $false  # After processing the first document, set this flag to false
    }

    # Attempt to save the merged document to the temporary location
    Write-Output "Attempting to save the merged document temporarily as: '$tempFilePath'"
    try {
        $null = $mergedDoc.SaveAs($tempFilePath, [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault)
        Write-Output "Document saved successfully to temp location: '$tempFilePath'"
    } catch {
        Write-Error "Failed to save document to temp location: $_"
        throw $_
    } finally {
        $mergedDoc.Close()
        Write-Output "Closed merged document."
    }

    # Now move the file to the final desired location
    Write-Output "Moving merged document to final location: '$finalOutputFile'"
    Move-Item -Path $tempFilePath -Destination $finalOutputFile -Force
    Write-Output "Merged document successfully moved to: '$finalOutputFile'"

} catch {
    Write-Error "An unexpected error occurred: $_"
    Read-Host "Press Enter to exit..."
    exit 1
} finally {
    # Clean up Word application
    if ($word -ne $null) {
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

# If no errors occurred, still wait for user input before closing
Read-Host "Press Enter to exit..."
