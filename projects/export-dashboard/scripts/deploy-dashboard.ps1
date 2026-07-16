param(
  [Parameter(Mandatory = $true)][string]$DataWorkbook,
  [string]$DataFileDate,
  [int]$MklsMaxYear,
  [int]$MklsMaxMonth,
  [string]$Scope
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$buildArgs = @("-DataWorkbook", $DataWorkbook)
if ($DataFileDate) { $buildArgs += @("-DataFileDate", $DataFileDate) }
if ($MklsMaxYear) { $buildArgs += @("-MklsMaxYear", $MklsMaxYear) }
if ($MklsMaxMonth) { $buildArgs += @("-MklsMaxMonth", $MklsMaxMonth) }
& (Join-Path $PSScriptRoot "build-dashboard.ps1") @buildArgs

$vercel = Get-Command vercel -ErrorAction SilentlyContinue
if (-not $vercel) { throw "Install and authenticate the Vercel CLI before deployment." }
$env:NODE_USE_ENV_PROXY = "1"
$vercelArgs = @("--prod", "--yes")
if ($Scope) { $vercelArgs += @("--scope", $Scope) }

Push-Location (Join-Path $repoRoot "web")
try { & vercel @vercelArgs } finally { Pop-Location }
