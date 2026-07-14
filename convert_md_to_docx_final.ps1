# PowerShell script to convert markdown to Word document using Word COM Object
# Pure ASCII Source code to prevent parser crashes on Korean Windows.

# Find the most recently modified .md file in the current directory dynamically
$mdFile = Get-ChildItem -Filter "*.md" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($null -eq $mdFile) {
    Write-Error "No markdown (.md) file found in the current directory!"
    Exit 1
}

$mdPath = $mdFile.FullName
$docxPath = $mdPath.Replace(".md", ".docx")

Write-Host "Resolved MD Path: $mdPath"
Write-Host "Resolved DOCX Path: $docxPath"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()
$selection = $word.Selection

# Read UTF-8 markdown file safely using .NET API
$lines = [System.IO.File]::ReadAllLines($mdPath, [System.Text.Encoding]::UTF8)

$inTable = $false
$tableLines = New-Object System.Collections.Generic.List[String]

function Add-Paragraph ($text, $type = "Normal") {
    if ($type -eq "H1") {
        $selection.Font.Size = 20
        $selection.Font.Bold = $true
    }
    elseif ($type -eq "H2") {
        $selection.Font.Size = 15
        $selection.Font.Bold = $true
    }
    elseif ($type -eq "H3") {
        $selection.Font.Size = 12
        $selection.Font.Bold = $true
    }
    else {
        $selection.Font.Size = 10
        $selection.Font.Bold = $false
    }
    
    $selection.TypeText($text)
    $selection.TypeParagraph()
    
    $selection.Font.Size = 10
    $selection.Font.Bold = $false
}

function Add-Hyperlink ($text, $url) {
    $range = $selection.Range
    $doc.Hyperlinks.Add($range, $url, $null, $null, $text) | Out-Null
    $selection.EndKey(5) | Out-Null
    $selection.TypeText(" ")
}

function Process-Text-With-Links ($text) {
    $pattern = '\[([^\]]+)\]\(([^)]+)\)'
    $currentIndex = 0
    
    while ($text -match $pattern) {
        $match = $Matches[0]
        $linkText = $Matches[1]
        $url = $Matches[2]
        
        $matchIndex = $text.IndexOf($match)
        if ($matchIndex -gt $currentIndex) {
            $plainText = $text.Substring($currentIndex, $matchIndex - $currentIndex)
            $selection.TypeText($plainText)
        }
        
        Add-Hyperlink $linkText $url
        
        $text = $text.Substring($matchIndex + $match.Length)
    }
    
    if ($text.Length -gt 0) {
        $selection.TypeText($text)
    }
    $selection.TypeParagraph()
}

function Build-Table ($tLines) {
    $rowsData = New-Object System.Collections.Generic.List[Object]
    $maxCols = 0
    
    foreach ($line in $tLines) {
        if ($line -match '^\s*\|?\s*:\s*[-=]+') {
            continue
        }
        $cleanLine = $line.Trim()
        if ($cleanLine.StartsWith("|")) { $cleanLine = $cleanLine.Substring(1) }
        if ($cleanLine.EndsWith("|")) { $cleanLine = $cleanLine.Substring(0, $cleanLine.Length - 1) }
        
        $cols = $cleanLine.Split("|") | ForEach-Object { $_.Trim() }
        if ($cols.Count -gt $maxCols) { $maxCols = $cols.Count }
        $rowsData.Add($cols)
    }
    
    if ($rowsData.Count -eq 0 -or $maxCols -eq 0) { return }
    
    $range = $selection.Range
    $table = $doc.Tables.Add($range, $rowsData.Count, $maxCols)
    $table.Borders.Enable = $true
    
    for ($r = 0; $r -lt $rowsData.Count; $r++) {
        $cols = $rowsData[$r]
        for ($c = 0; $c -lt $cols.Count; $c++) {
            $cell = $table.Cell($r + 1, $c + 1)
            $cellText = $cols[$c]
            
            $cell.Range.Text = ""
            $cell.Range.Select()
            
            $linkPattern = '\[([^\]]+)\]\(([^)]+)\)'
            if ($cellText -match $linkPattern) {
                $lText = $Matches[1]
                $lUrl = $Matches[2]
                $doc.Hyperlinks.Add($cell.Range, $lUrl, $null, $null, $lText) | Out-Null
            } else {
                $cell.Range.Text = $cellText
            }
        }
    }
    
    $doc.Characters.Last.Select()
    $selection.Collapse(0)
    $selection.TypeParagraph()
}

foreach ($line in $lines) {
    $lineTrimmed = $line.Trim()
    
    if ($lineTrimmed.StartsWith("|")) {
        $inTable = $true
        $tableLines.Add($line)
        continue
    } else {
        if ($inTable) {
            Build-Table $tableLines
            $tableLines.Clear()
            $inTable = $false
        }
    }
    
    if ($lineTrimmed -match '^#\s+(.+)$') {
        Add-Paragraph $Matches[1] "H1"
    }
    elseif ($lineTrimmed -match '^##\s+(.+)$') {
        Add-Paragraph $Matches[1] "H2"
    }
    elseif ($lineTrimmed -match '^###\s+(.+)$') {
        Add-Paragraph $Matches[1] "H3"
    }
    elseif ($lineTrimmed -eq "---" -or $lineTrimmed -eq "") {
        if ($lineTrimmed -eq "---") {
            $selection.TypeParagraph()
        }
    }
    else {
        $cleanText = $lineTrimmed -replace '\*\*([^*]+)\*\*', '$1'
        $cleanText = $cleanText -replace '\* ([^*]+)', '$1'
        Process-Text-With-Links $cleanText
    }
}

if ($inTable) {
    Build-Table $tableLines
}

$doc.SaveAs($docxPath)
$doc.Close()
$word.Quit()

Write-Host "DOCX file successfully created at: $docxPath"

# Cleanup: Dynamically construct folder name '불필요한 작업파일' using unicode hex encoding to bypass parser crashes
$chars = [char[]]@(0xBD88, 0xD544, 0xC694, 0xD55C, 0x0020, 0xC791, 0xC5C5, 0xD30C, 0xC77C)
$targetDirName = New-Object String($chars, 0, $chars.Length)
$targetDir = Join-Path $mdFile.DirectoryName $targetDirName

if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir | Out-Null
}

Move-Item -Path $mdPath -Destination $targetDir -Force
Write-Host "Cleanup completed successfully."

# 1. 텍스트 취합 및 HTML 블로그 원고 생성
Write-Host "Merging reports and generating blog HTML..."
$pythonPath = "python"
if (Test-Path "process_and_verify.py") {
    & $pythonPath process_and_verify.py
} else {
    Write-Warning "process_and_verify.py not found. Skipping HTML merge."
}

# 2. Git Auto Push (DOCX 및 blog_post.html 포함)
$gitPath = "C:\Program Files\Git\cmd\git.exe"
if (Test-Path $gitPath) {
    Write-Host "Running Git Auto Push..."
    & $gitPath add .
    & $gitPath commit -m "Auto-commit: update boxoffice reports and blog HTML"
    & $gitPath push origin main
    Write-Host "Git Auto Push completed."
} else {
    Write-Warning "Git executable not found at $gitPath. Skipping auto-push."
}
