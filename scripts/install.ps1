<#
  Sets up the SQL Query add-in for Excel desktop (current user only, no admin needed):
    1. installs npm dependencies
    2. creates a localhost HTTPS certificate and trusts it for the current user
    3. registers (sideloads) manifest.xml with Excel
    4. optionally starts the server at logon (-AutoStart)
#>
param([switch]$AutoStart)
$ErrorActionPreference = "Stop"
$root = Resolve-Path (Join-Path $PSScriptRoot "..")
$certs = Join-Path $root "certs"
$manifest = Join-Path $root "manifest.xml"
$addinId = ([xml](Get-Content $manifest -Raw)).OfficeApp.Id

$node = (Get-Command node -ErrorAction SilentlyContinue).Source
if (-not $node) { throw "Node.js not found on PATH." }

Push-Location $root
try {
    if (-not (Test-Path (Join-Path $root "node_modules\sql.js"))) {
        Write-Host "Installing npm packages..."
        npm install --omit=dev
        if ($LASTEXITCODE) { throw "npm install failed" }
    }
} finally { Pop-Location }

if (-not (Test-Path (Join-Path $root "public\assets\icon-32.png"))) { & (Join-Path $PSScriptRoot "make-icons.ps1") }

# --- certificate -----------------------------------------------------------
New-Item -ItemType Directory -Force $certs | Out-Null
$pfxPath = Join-Path $certs "localhost.pfx"
if (-not (Test-Path $pfxPath)) {
    Write-Host "Creating localhost certificate..."
    $cert = New-SelfSignedCertificate -Subject "CN=localhost" -FriendlyName "Excel SQL Add-in (localhost)" `
        -TextExtension @("2.5.29.17={text}DNS=localhost&IPAddress=127.0.0.1&IPAddress=::1", "2.5.29.37={text}1.3.6.1.5.5.7.3.1") `
        -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable -KeyAlgorithm RSA -KeyLength 2048 `
        -HashAlgorithm SHA256 -NotAfter (Get-Date).AddYears(5)
    $pass = [Convert]::ToBase64String((1..24 | ForEach-Object { Get-Random -Maximum 256 }))
    Set-Content -Path (Join-Path $certs "pfx-pass.txt") -Value $pass -Encoding ascii -NoNewline
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password (ConvertTo-SecureString $pass -AsPlainText -Force) `
        -CryptoAlgorithmOption AES256_SHA256 | Out-Null
    $cer = Join-Path $certs "localhost.cer"
    Export-Certificate -Cert $cert -FilePath $cer | Out-Null
    Set-Content -Path (Join-Path $certs "thumbprint.txt") -Value $cert.Thumbprint -Encoding ascii -NoNewline
    Write-Host "Trusting the certificate for your user. Windows will ask you to confirm - click Yes." -ForegroundColor Yellow
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store "Root", "CurrentUser"
    $store.Open("ReadWrite")
    try { $store.Add((New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $cer)) } finally { $store.Close() }
}

# --- sideload --------------------------------------------------------------
$dev = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
New-Item -Force $dev -ErrorAction SilentlyContinue | Out-Null
Set-ItemProperty -Path $dev -Name $addinId -Value $manifest
Write-Host "Registered add-in $addinId with Excel."

# --- auto start ------------------------------------------------------------
$startup = Join-Path ([Environment]::GetFolderPath("Startup")) "Excel SQL Add-in Server.lnk"
if ($AutoStart) {
    $sh = New-Object -ComObject WScript.Shell
    $lnk = $sh.CreateShortcut($startup)
    $lnk.TargetPath = $node
    $lnk.Arguments = "`"$(Join-Path $root 'server.js')`""
    $lnk.WorkingDirectory = "$root"
    $lnk.WindowStyle = 7 # minimized
    $lnk.Description = "Serves the Excel SQL Query add-in on https://localhost:3021"
    $lnk.Save()
    Write-Host "Server will start automatically at logon."
}

# --- start server now ------------------------------------------------------
$running = Get-NetTCPConnection -LocalPort 3021 -State Listen -ErrorAction SilentlyContinue
if (-not $running) {
    Start-Process -FilePath $node -ArgumentList "`"$(Join-Path $root 'server.js')`"" -WorkingDirectory $root -WindowStyle Minimized
    Write-Host "Server started on https://localhost:3021"
}

Write-Host ""
Write-Host "Done. Restart Excel, then open Data > SQL Query." -ForegroundColor Green
Write-Host "If the button is missing: Home > Add-ins > More Add-ins > Developer Add-ins (or Shared Folder) > SQL Query."
