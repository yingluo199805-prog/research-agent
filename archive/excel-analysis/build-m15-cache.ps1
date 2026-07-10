param(
  [Parameter(Mandatory = $true)][string]$SourceMkls,
  [Parameter(Mandatory = $true)][string]$TargetWorkbook,
  [Parameter(Mandatory = $true)][string]$OutputCache
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $repoRoot ".venv\\Scripts\\python.exe"
if (-not (Test-Path -LiteralPath $python)) { $python = "python" }

& $python (Join-Path $repoRoot "tools\\build_mkls_m15_cache.py") `
  --source-mkls $SourceMkls `
  --target-workbook $TargetWorkbook `
  --output $OutputCache
