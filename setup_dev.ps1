# Geliştirme ortamını tek komutla kurar: Python 3.12 kontrolü, .venv, bağımlılıklar.
# Kullanım: powershell -ExecutionPolicy Bypass -File .\setup_dev.ps1

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$py = $null
foreach ($candidate in @("py -3.12", "python")) {
    try {
        $ver = & cmd /c "$candidate --version" 2>$null
        if ($ver -match "Python 3\.12") { $py = $candidate; break }
    } catch { }
}
if (-not $py) {
    throw "Python 3.12 bulunamadı. https://www.python.org/downloads/ adresinden 64-bit Python 3.12 kurun."
}

if (-not (Test-Path ".venv")) {
    & cmd /c "$py -m venv .venv"
}

& ".\.venv\Scripts\python.exe" -m pip install --upgrade pip
& ".\.venv\Scripts\python.exe" -m pip install -r requirements.txt
& ".\.venv\Scripts\python.exe" -m unittest discover -s tests

Write-Host ""
Write-Host "Kurulum tamam. Uygulamayı çalıştırmak için: .\.venv\Scripts\python.exe app.py"
