$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path -LiteralPath $PSScriptRoot).Path
Set-Location -LiteralPath $projectRoot
$distTarget = [IO.Path]::GetFullPath((Join-Path $projectRoot 'dist\ThaVaDoi'))
if (-not $distTarget.StartsWith($projectRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Đường dẫn đóng gói nằm ngoài dự án.'
}

$python = Join-Path $projectRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python)) {
    python -m venv .venv
}
& $python -m pip install -r requirements-build.txt
if ($LASTEXITCODE -ne 0) { throw 'Không cài được thư viện.' }

$icon = Join-Path $projectRoot 'assets\icon.ico'
if (-not (Test-Path -LiteralPath $icon)) {
    & $python tools\make_icon.py
    if ($LASTEXITCODE -ne 0) { throw 'Không tạo được biểu tượng.' }
}

$ocrCandidates = @()
if ($env:TESSERACT_HOME) { $ocrCandidates += $env:TESSERACT_HOME }
$ocrCandidates += @('C:\Program Files\Tesseract-OCR', 'C:\Program Files (x86)\Tesseract-OCR')
$ocrOnPath = Get-Command tesseract.exe -ErrorAction SilentlyContinue
if ($ocrOnPath) { $ocrCandidates += (Split-Path -Parent $ocrOnPath.Source) }
$ocrHome = $ocrCandidates | Where-Object { Test-Path -LiteralPath (Join-Path $_ 'tesseract.exe') } | Select-Object -First 1
if (-not $ocrHome) { throw 'Cần cài Tesseract OCR trên máy đóng gói hoặc đặt TESSERACT_HOME.' }

# Qt uses Windows' ICU DLL. A Poppler directory on PATH can make PyInstaller
# collect an incompatible ICU DLL and prevent the GUI from starting.
$env:PATH = (($env:PATH -split ';' | Where-Object {
    $_ -and ($_ -notmatch '(?i)[\\/]poppler[\\/].*[\\/]bin[\\/]?$')
}) -join ';')

& $python -m PyInstaller `
    --noconfirm `
    --workpath build-sanitized `
    --onedir `
    --windowed `
    --name ThaVaDoi `
    --icon $icon `
    --add-data 'ocr_data;ocr_data' `
    --exclude-module PySide6.QtWebEngineCore `
    --exclude-module PySide6.QtWebEngineWidgets `
    --exclude-module PySide6.QtWebEngineQuick `
    --exclude-module tkinter `
    run.py
if ($LASTEXITCODE -ne 0) { throw 'Đóng gói thất bại.' }

$ocrTarget = Join-Path $distTarget '_internal\tesseract'
New-Item -ItemType Directory -Path $ocrTarget -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $ocrHome 'tesseract.exe') -Destination $ocrTarget -Force
Get-ChildItem -LiteralPath $ocrHome -File -Filter '*.dll' | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $ocrTarget -Force
}
$ocrLicense = Join-Path $ocrHome 'doc\LICENSE'
if (Test-Path -LiteralPath $ocrLicense) {
    Copy-Item -LiteralPath $ocrLicense -Destination (Join-Path $ocrTarget 'LICENSE-Tesseract.txt') -Force
}

Write-Host "Đã tạo: $projectRoot\dist\ThaVaDoi\ThaVaDoi.exe"
$archive = Join-Path $projectRoot 'ThaVaDoi-Windows-0.5.0.zip'
Compress-Archive -LiteralPath $distTarget -DestinationPath $archive -CompressionLevel Optimal -Force
Write-Host "Đã tạo: $archive"
