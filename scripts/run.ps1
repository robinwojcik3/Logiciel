<#
  Lance l'application via l'environnement virtuel local (.venv).

  Utilisation:
    powershell -ExecutionPolicy Bypass -File scripts/run.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$RepoRoot = Split-Path -Parent $PSScriptRoot
$VenvPy   = Join-Path $RepoRoot ".venv/Scripts/python.exe"
$Entry    = Join-Path $RepoRoot "Start.py"

if (-not (Test-Path $VenvPy)) {
  Write-Host ".venv introuvable. Exécutez d'abord scripts\\setup.ps1" -ForegroundColor Yellow
  exit 1
}

& $VenvPy $Entry

