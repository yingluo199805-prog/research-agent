param(
  [string]$PythonCommand = "python",
  [switch]$SkipWebDependencies
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$venv = Join-Path $repoRoot ".venv"

& $PythonCommand -m venv $venv
& (Join-Path $venv "Scripts\\python.exe") -m pip install --upgrade pip
& (Join-Path $venv "Scripts\\python.exe") -m pip install -r (Join-Path $repoRoot "requirements.txt")

if (-not $SkipWebDependencies) {
  $npm = Get-Command npm -ErrorAction SilentlyContinue
  if (-not $npm) { throw "Node.js 20+ and npm are required for the Vercel application." }
  Push-Location (Join-Path $repoRoot "web")
  try { & npm install } finally { Pop-Location }
}

Write-Output "Environment ready: $venv"
