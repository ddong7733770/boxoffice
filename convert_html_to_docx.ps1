param (
    [string]$HtmlPath,
    [string]$DocxPath
)

if (-not $HtmlPath -or -not $DocxPath) {
    Write-Error "Both HtmlPath and DocxPath parameters are required."
    exit 1
}

if (-not (Test-Path $HtmlPath)) {
    Write-Error "HTML file not found: $HtmlPath"
    exit 1
}

$absoluteHtml = (Resolve-Path $HtmlPath).Path

# Resolve target parent directory path
$parentDir = Split-Path -Parent $DocxPath
if (-not $parentDir) { $parentDir = "." }
$absoluteParent = (Resolve-Path $parentDir).Path
$docxFileName = Split-Path -Leaf $DocxPath
$absoluteDocx = Join-Path $absoluteParent $docxFileName

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    
    # Open HTML file
    $doc = $word.Documents.Open($absoluteHtml, $false, $true)
    
    # Save as DOCX format (wdFormatXMLDocument = 12)
    $doc.SaveAs2($absoluteDocx, 12)
    $doc.Close($false)
    $word.Quit()
    
    # Clean COM objects from memory
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    
    Write-Output "Successfully converted HTML to DOCX: $absoluteDocx"
}
catch {
    Write-Error "Failed to convert HTML to DOCX: $_"
    if ($word) {
        try { $word.Quit() } catch {}
    }
    exit 1
}
