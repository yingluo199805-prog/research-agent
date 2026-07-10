param(
  [Parameter(Mandatory = $true)][string]$MainWorkbook,
  [Parameter(Mandatory = $true)][string]$MklsRaw,
  [Parameter(Mandatory = $true)][string]$CustomsRaw,
  [Parameter(Mandatory = $true)][string]$Report,
  [switch]$Writeback
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $repoRoot ".venv\\Scripts\\python.exe"
if (-not (Test-Path -LiteralPath $python)) { $python = "python" }

$pythonArgs = @(
  (Join-Path $repoRoot "src\\data_cleaning\\refresh_master_workbook.py"),
  "--main", $MainWorkbook,
  "--mkls-raw", $MklsRaw,
  "--hg-raw", $CustomsRaw,
  "--report", $Report
)
if ($Writeback) { $pythonArgs += "--writeback" }

& $python @pythonArgs
