param(
    [string]$ServiceName = "AdmissionDiagStreamlit",
    [string]$AppDir = "C:\Users\chris\Desktop\입시위치진단서비스",
    [string]$Port = "8501"
)

$nssm = "C:\tools\nssm\nssm.exe"
if (-not (Test-Path -LiteralPath $nssm)) {
    Write-Host "NSSM not found: $nssm"
    Write-Host "Download NSSM and place nssm.exe at C:\tools\nssm\nssm.exe"
    exit 1
}

$pythonExe = Join-Path $AppDir ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $pythonExe)) {
    Write-Host "Python venv not found: $pythonExe"
    Write-Host "Run: python -m venv .venv && .\.venv\Scripts\activate && pip install -r requirements.txt"
    exit 1
}

& $nssm install $ServiceName $pythonExe "-m streamlit run app.py --server.port $Port --server.address 0.0.0.0"
& $nssm set $ServiceName AppDirectory $AppDir
& $nssm set $ServiceName Start SERVICE_AUTO_START
& $nssm start $ServiceName

Write-Host "Service installed and started: $ServiceName"
