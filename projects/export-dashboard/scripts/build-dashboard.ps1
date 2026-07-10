param(
  [Parameter(Mandatory = $true)][string]$DataWorkbook,
  [string]$DataFileDate
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$builderRoot = Join-Path $repoRoot "src\\dashboard_build"
$python = Join-Path $repoRoot ".venv\\Scripts\\python.exe"
if (-not (Test-Path -LiteralPath $python)) { $python = "python" }

$resolvedWorkbook = (Resolve-Path -LiteralPath $DataWorkbook).Path
if (-not $DataFileDate) {
  $name = [IO.Path]::GetFileNameWithoutExtension($resolvedWorkbook)
  if ($name -notmatch '(?<stamp>\d{6})') { throw "Use -DataFileDate YYYY-MM-DD when the workbook name has no YYMMDD stamp." }
  $stamp = $Matches['stamp']
  $DataFileDate = "20$($stamp.Substring(0,2))-$($stamp.Substring(2,2))-$($stamp.Substring(4,2))"
}

$env:DASHBOARD_DATA_FILE = $resolvedWorkbook
$env:DASHBOARD_DATA_FILE_DATE = $DataFileDate
Push-Location $builderRoot
try {
  foreach ($script in @('gen_data.py', 'gen_data_caam.py', 'gen_data_mkls.py', 'build.py')) {
    & $python (Join-Path $builderRoot $script)
    if ($LASTEXITCODE -ne 0) { throw "$script failed with exit code $LASTEXITCODE" }
  }
} finally {
  Pop-Location
}

$html = Join-Path $builderRoot "index.html"
if (-not (Test-Path -LiteralPath $html) -or (Get-Item -LiteralPath $html).Length -lt 1MB) {
  throw "Dashboard HTML was not generated or is unexpectedly small."
}
Copy-Item -LiteralPath $html -Destination (Join-Path $repoRoot "web\\public\\dashboard.html") -Force
Get-Item -LiteralPath $html, (Join-Path $repoRoot "web\\public\\dashboard.html") | Select-Object FullName,Length,LastWriteTime
