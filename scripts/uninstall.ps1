# Removes the add-in registration, the trusted localhost certificate and the logon shortcut.
$root = Resolve-Path (Join-Path $PSScriptRoot "..")
$addinId = ([xml](Get-Content (Join-Path $root "manifest.xml") -Raw)).OfficeApp.Id

Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Name $addinId -ErrorAction SilentlyContinue
Write-Host "Unregistered add-in."

$thumbFile = Join-Path $root "certs\thumbprint.txt"
if (Test-Path $thumbFile) {
    $thumb = (Get-Content $thumbFile -Raw).Trim()
    foreach ($store in "Cert:\CurrentUser\Root", "Cert:\CurrentUser\My") {
        Get-ChildItem $store | Where-Object Thumbprint -eq $thumb | Remove-Item -ErrorAction SilentlyContinue
    }
    Remove-Item -Recurse -Force (Join-Path $root "certs")
    Write-Host "Removed certificate."
}

$startup = Join-Path ([Environment]::GetFolderPath("Startup")) "Excel SQL Add-in Server.lnk"
if (Test-Path $startup) { Remove-Item $startup; Write-Host "Removed logon shortcut." }

Get-CimInstance Win32_Process -Filter "Name='node.exe'" |
    Where-Object { $_.CommandLine -like "*excel-sql-addin*server.js*" } |
    ForEach-Object { Stop-Process -Id $_.ProcessId -Force; Write-Host "Stopped server." }
