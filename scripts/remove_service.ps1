param(
    [string]$ServiceName = "AdmissionDiagStreamlit"
)

$nssm = "C:\tools\nssm\nssm.exe"
if (-not (Test-Path -LiteralPath $nssm)) {
    Write-Host "NSSM not found: $nssm"
    exit 1
}

& $nssm stop $ServiceName
& $nssm remove $ServiceName confirm

Write-Host "Service removed: $ServiceName"
