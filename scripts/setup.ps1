<#
  Configure un environnement virtuel local (.venv) et installe
  les dépendances Python listées dans requirements.txt.

  Utilisation:
    powershell -ExecutionPolicy Bypass -File scripts/setup.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Section($text) {
  Write-Host "`n=== $text ===" -ForegroundColor Cyan
}

# Détermine les chemins utiles
$RepoRoot = Split-Path -Parent $PSScriptRoot
$VenvDir  = Join-Path $RepoRoot ".venv"
$PyExe    = $null

Write-Section "Détection de Python"
try {
  # Préférence au lanceur Windows pour Python 3.13 s'il est disponible
  $pyPath = & py -3.13 -c "import sys; print(sys.executable)" 2>$null
  if ($LASTEXITCODE -eq 0 -and $pyPath) {
    $PyExe = $pyPath.Trim()
  }
} catch {}

if (-not $PyExe) {
  try {
    $pyPath = & python -c "import sys; print(sys.executable)" 2>$null
    if ($LASTEXITCODE -eq 0 -and $pyPath) { $PyExe = $pyPath.Trim() }
  } catch {}
}

if (-not $PyExe) {
  throw "Python introuvable. Installez Python 3.11+ et relancez."
}

Write-Host "Python: $PyExe"

Write-Section "Création/MAJ de l'environnement virtuel (.venv)"
if (-not (Test-Path $VenvDir)) {
  & $PyExe -m venv $VenvDir
} else {
  Write-Host ".venv déjà présent, réutilisation."
}

$VenvPy = Join-Path $VenvDir "Scripts/python.exe"
$VenvPip = Join-Path $VenvDir "Scripts/pip.exe"

Write-Section "Mise à jour de pip/setuptools/wheel"
& $VenvPy -m pip install --upgrade pip setuptools wheel

Write-Section "Installation des dépendances (requirements.txt)"
$ReqFile = Join-Path $RepoRoot "requirements.txt"
if (-not (Test-Path $ReqFile)) { throw "requirements.txt introuvable à la racine du dépôt." }
& $VenvPip install -r $ReqFile
if ($LASTEXITCODE -ne 0) {
  throw "Échec d'installation de requirements.txt (vérifiez la connexion réseau)."
}

Write-Section "Vérification rapide des imports"
$testScript = @'
import importlib, sys
mods = [
    "geopandas","requests","PIL","pillow_heif","selenium","docx","bs4","openpyxl"
]
missing = []
for m in mods:
    try:
        importlib.import_module(m)
    except Exception as e:
        missing.append(f"{m}: {e}")
if missing:
    print("Manquants/erreurs:\n - " + "\n - ".join(missing))
    raise SystemExit(1)
print("Tous les imports de base OK")
'@

$tmp = Join-Path $env:TEMP "venv_import_test.py"
Set-Content -Path $tmp -Value $testScript -Encoding UTF8
try {
  & $VenvPy $tmp
} finally {
  Remove-Item -Path $tmp -ErrorAction SilentlyContinue
}
if ($LASTEXITCODE -ne 0) {
  throw "Vérification des imports échouée (voir détails ci-dessus)."
}

Write-Section "Terminé"
Write-Host "Environnement prêt. Pour lancer l'application :" -ForegroundColor Green
Write-Host "  scripts\\run.ps1" -ForegroundColor Green

Write-Host "\nNotes:" -ForegroundColor DarkGray
Write-Host "- QGIS 3.x doit être installé (chemin configuré dans modules\\main_app.py)." -ForegroundColor DarkGray
Write-Host "- Selenium nécessite Google Chrome; le driver est géré automatiquement (Selenium Manager)." -ForegroundColor DarkGray
