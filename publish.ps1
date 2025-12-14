# Publish both projects into a single artifacts/publish folder for Windows x64 self-contained single-file
$ErrorActionPreference = 'Stop'

$solutionRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$publishDir = Join-Path $solutionRoot 'artifacts\publish'

Write-Host "Cleaning publish folder: $publishDir"
if (Test-Path $publishDir) { Remove-Item $publishDir -Recurse -Force }
New-Item -ItemType Directory -Path $publishDir | Out-Null

Write-Host "Publishing ServizioAntiCopieMultiple (service)"
dotnet publish "ServizioAntiCopieMultiple/ServizioAntiCopieMultiple.csproj" -c Release -r win-x64 -p:PublishSingleFile=true -p:SelfContained=true -o $publishDir

Write-Host "Publishing GestioneSACM (tool)"
dotnet publish "GestioneSACM/GestioneSACM.csproj" -c Release -r win-x64 -p:PublishSingleFile=true -p:SelfContained=true -o $publishDir

Write-Host "Publish complete. Outputs in: $publishDir"