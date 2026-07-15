# PowerShell script to trigger python document processing and git push
# Pure ASCII Source code to prevent parser crashes on Korean Windows.

# 1. Run python processing (generates both docx and html)
Write-Host "Triggering Python processing script..."
$pythonPath = "python"
if (Test-Path "process_and_verify.py") {
    & $pythonPath process_and_verify.py
} else {
    Write-Warning "process_and_verify.py not found!"
}

# 2. Cleanup: Move all md files to '불필요한 작업파일'
# "불필요한 작업파일" is 0xBD88, 0xD544, 0xC694, 0xD55C, 0x0020, 0xC791, 0xC5C5, 0xD30C, 0xC77C
$charsCleanup = [char[]]@(0xBD88, 0xD544, 0xC694, 0xD55C, 0x0020, 0xC791, 0xC5C5, 0xD30C, 0xC77C)
$targetDirName = New-Object String($charsCleanup, 0, $charsCleanup.Length)
$targetDir = Join-Path $PSScriptRoot $targetDirName

if (-not (Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir | Out-Null
}

Get-ChildItem -Filter "*.md" | ForEach-Object {
    Move-Item -Path $_.FullName -Destination $targetDir -Force
}
Write-Host "Cleanup completed successfully."

# 3. Git Auto Push
$gitPath = "C:\Program Files\Git\cmd\git.exe"
if (Test-Path $gitPath) {
    Write-Host "Running Git Auto Push..."
    & $gitPath add .
    & $gitPath commit -m "Auto-commit: update boxoffice reports and blog HTML (python-docx version)"
    & $gitPath push origin main
    Write-Host "Git Auto Push completed."
} else {
    Write-Warning "Git executable not found at $gitPath. Skipping auto-push."
}
