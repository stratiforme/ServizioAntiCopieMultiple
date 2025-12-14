param(
    [switch]$Uninstall
)

$publishDir = Join-Path (Get-Location) "artifacts\publish"
if (-not (Test-Path $publishDir)) { $publishDir = Get-Location }

$serviceExe = Join-Path $publishDir "ServizioAntiCopieMultiple.exe"
$toolExe = Join-Path $publishDir "gestionesacm.exe"

if ($Uninstall)
{
    Write-Host "Stopping and uninstalling service..."
    sc.exe stop ServizioAntiCopieMultiple | Out-Null
    sc.exe delete ServizioAntiCopieMultiple | Out-Null
    Write-Host "Service removed (if existed)."
    exit 0
}

if (-not (Test-Path $serviceExe))
{
    Write-Host "Servizio executable not found in $publishDir"
    Write-Host "Listing files in $publishDir"
    Get-ChildItem $publishDir
    Read-Host "Place ServizioAntiCopieMultiple.exe in this folder and press Enter to continue"
}

if (-not (Test-Path $serviceExe)) { Write-Error "Service executable not found. Aborting."; exit 1 }

Write-Host "Creating service..."
sc.exe create ServizioAntiCopieMultiple binPath= `"$serviceExe`" start= auto DisplayName= "Servizio Anti Copie Multiple" | Write-Host
Write-Host "Starting service..."
sc.exe start ServizioAntiCopieMultiple | Write-Host
Write-Host "Done. Use gestionesacm.exe to manage the service interactively."
