# FreeConvert için tek dosyalık, taşınabilir Windows .exe üretir.
# Tesseract OCR motoru ve tur/eng dil verisi tesseract_runtime/ klasöründen
# exe içine gömülür. Word, LibreOffice, OnlyOffice, Java veya sistem Tesseract
# kurulumu gerekmez. Tüm dönüşümler uygulama içindeki açık kaynak motorlarla
# yapılır.
#
# Kullanım:
#   powershell -ExecutionPolicy Bypass -File .\build_portable.ps1
#   powershell -ExecutionPolicy Bypass -File .\build_portable.ps1 -SkipInstall

[CmdletBinding()]
param(
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
$projectRoot = $PSScriptRoot
if (-not $projectRoot) { $projectRoot = (Get-Location).Path }
Set-Location -LiteralPath $projectRoot

function Invoke-Native {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )
    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$FilePath failed with exit code $LASTEXITCODE"
    }
}

if (-not [Environment]::Is64BitOperatingSystem) {
    throw "This portable build requires 64-bit Windows."
}

$python = (Get-Command python -ErrorAction Stop).Source
$pythonArch = (& $python -c "import struct; print(struct.calcsize('P') * 8)").Trim()
if ($LASTEXITCODE -ne 0 -or $pythonArch -ne "64") {
    throw "64-bit Python is required (found: $pythonArch)."
}

$runtime = Join-Path $projectRoot "tesseract_runtime"
$requiredRuntime = @(
    "tesseract.exe",
    "libtesseract-5.dll",
    "libleptonica-6.dll",
    "tessdata\eng.traineddata",
    "tessdata\tur.traineddata"
)
foreach ($relative in $requiredRuntime) {
    $path = Join-Path $runtime $relative
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Missing bundled OCR runtime file: $path"
    }
}

$fontFiles = @(
    "assets\fonts\NotoSans-Regular.ttf",
    "assets\fonts\NotoSans-Bold.ttf",
    "assets\fonts\NotoSans-Italic.ttf",
    "assets\fonts\OFL.txt",
    "THIRD_PARTY_NOTICES.md",
    "LICENSE"
)
foreach ($relative in $fontFiles) {
    if (-not (Test-Path -LiteralPath (Join-Path $projectRoot $relative) -PathType Leaf)) {
        throw "Missing bundled font/license file: $relative"
    }
}

$tesseract = Join-Path $runtime "tesseract.exe"
$tessdata = Join-Path $runtime "tessdata"
& $tesseract --tessdata-dir $tessdata --list-langs
if ($LASTEXITCODE -ne 0) {
    throw "Bundled Tesseract validation failed."
}

if (-not $SkipInstall) {
    Invoke-Native $python @(
        "-m", "pip", "install", "--disable-pip-version-check",
        "-r", (Join-Path $projectRoot "requirements.txt")
    )
}

$spec = Join-Path $projectRoot "FreeConvert.spec"
if (-not (Test-Path -LiteralPath $spec -PathType Leaf)) {
    throw "PyInstaller spec is missing: $spec"
}
Invoke-Native $python @(
    "-m", "PyInstaller", "--noconfirm", "--clean", $spec
)

$output = Join-Path $projectRoot "dist\FreeConvert.exe"
if (-not (Test-Path -LiteralPath $output -PathType Leaf)) {
    throw "Portable output was not created: $output"
}

$uninstallSpec = Join-Path $projectRoot "Uninstall.spec"
Invoke-Native $python @(
    "-m", "PyInstaller", "--noconfirm", "--clean", $uninstallSpec
)
$uninstaller = Join-Path $projectRoot "dist\uninstall.exe"
if (-not (Test-Path -LiteralPath $uninstaller -PathType Leaf)) {
    throw "Uninstaller was not created: $uninstaller"
}

Invoke-Native $python @(
    (Join-Path $projectRoot "tools\verify_portable.py"), $output
)

$hash = (Get-FileHash -LiteralPath $output -Algorithm SHA256).Hash
$size = (Get-Item -LiteralPath $output).Length
Write-Host "Portable output: $output"
Write-Host ("Size: {0:N0} bytes" -f $size)
Write-Host "SHA-256: $hash"
