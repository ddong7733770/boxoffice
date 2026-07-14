param (
    [string]$FilePath
)

if (-not $FilePath) {
    Write-Error "FilePath parameter is required"
    exit 1
}

if (-not (Test-Path $FilePath)) {
    Write-Error "File not found: $FilePath"
    exit 1
}

$absolutePath = (Resolve-Path $FilePath).Path

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0 # WdAlertLevel.wdAlertsNone
    
    $doc = $word.Documents.Open($absolutePath, $false, $true) # FileName, ConfirmConversions, ReadOnly
    $text = $doc.Content.Text
    $doc.Close($false) # Close without saving changes
    $word.Quit()
    
    # Clean COM object from memory to prevent winword process leak
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # Output text with UTF-8 to keep Korean intact
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    Write-Output $text
}
catch {
    Write-Error "Failed to extract text from docx: $_"
    if ($word) {
        try { $word.Quit() } catch {}
    }
    exit 1
}
