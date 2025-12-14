# Uninstall and stop the service
$publishDir = Join-Path (Get-Location) "artifacts\publish"
if (-not (Test-Path $publishDir)) { $publishDir = Get-Location }

Write-Host "Stopping service (if running)..."
sc.exe stop ServizioAntiCopieMultiple | Out-Null
Start-Sleep -Seconds 1

Write-Host "Deleting service (if exists)..."
sc.exe delete ServizioAntiCopieMultiple | Out-Null

Write-Host "Service removed (if existed)."
