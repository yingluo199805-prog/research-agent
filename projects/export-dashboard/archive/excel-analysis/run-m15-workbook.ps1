param(
  [Parameter(Mandatory = $true)][string]$Cache,
  [Parameter(Mandatory = $true)][string]$BaseWorkbook,
  [Parameter(Mandatory = $true)][string]$TemplateWorkbook,
  [Parameter(Mandatory = $true)][string]$OutputWorkbook,
  [string]$EuropeWorkbook
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $repoRoot ".venv\\Scripts\\python.exe"
if (-not (Test-Path -LiteralPath $python)) { $python = "python" }

$pythonArgs = @(
  (Join-Path $repoRoot "tools\\run_full_m15_workbook.py"),
  "--cache", $Cache,
  "--base-workbook", $BaseWorkbook,
  "--template-workbook", $TemplateWorkbook,
  "--output-workbook", $OutputWorkbook
)
if ($EuropeWorkbook) { $pythonArgs += @("--europe-workbook", $EuropeWorkbook) }

& $python @pythonArgs
