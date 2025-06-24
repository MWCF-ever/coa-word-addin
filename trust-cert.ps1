# Trust localhost certificate for Office Add-in
$certPath = Join-Path $PSScriptRoot "certs\localhost.crt"

if (Test-Path $certPath) {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
    $store.Open("ReadWrite")
    $store.Add($cert)
    $store.Close()
    Write-Host "Certificate trusted successfully!" -ForegroundColor Green
} else {
    Write-Host "Certificate not found at: $certPath" -ForegroundColor Red
    Write-Host "Please run 'npm run generate-cert' first" -ForegroundColor Yellow
}

# Also add to IE trusted sites
$registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\localhost"
if (!(Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force
}
New-ItemProperty -Path $registryPath -Name "https" -Value 2 -PropertyType DWord -Force

Write-Host "Added localhost to IE trusted sites" -ForegroundColor Green