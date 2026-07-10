param(
  [Parameter(Mandatory=$true)][string]$MainPath,
  [Parameter(Mandatory=$true)][string]$ConfigPath
)

$ErrorActionPreference = "Stop"

$cfg = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$brand = [string]$cfg.brand
$targetOem = [string]$cfg.targetOem
$year = [int]$cfg.year
$months = @{}
foreach ($m in $cfg.months) { $months[[int]$m] = $true }

$backupPath = $MainPath -replace '\.xlsx$', ("_preCaamLeapmotorFix_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
Copy-Item -LiteralPath $MainPath -Destination $backupPath -Force

$excel = $null
$wb = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false
  $excel.EnableEvents = $false

  $wb = $excel.Workbooks.Open($MainPath)

  # Sheet 1: 中汽协CV ; columns A/B/F/J = 年/月/品牌/车企划分.
  $wsCv = $wb.Worksheets.Item(1)
  $xlUp = -4162
  $lastRow = $wsCv.Cells.Item($wsCv.Rows.Count, 1).End($xlUp).Row
  $fixedRows = 0
  for ($r = 2; $r -le $lastRow; $r++) {
    $y = [int]$wsCv.Cells.Item($r, 1).Value2
    if ($y -ne $year) { continue }
    $m = [int]$wsCv.Cells.Item($r, 2).Value2
    if (-not $months.ContainsKey($m)) { continue }
    $b = [string]$wsCv.Cells.Item($r, 6).Value2
    if ($b -eq $brand) {
      $wsCv.Cells.Item($r, 10).Value2 = $targetOem
      $fixedRows++
    }
  }

  # Sheet 5: 中汽协匹配字段 ; columns D/E = brand -> OEM mapping.
  $wsMap = $wb.Worksheets.Item(5)
  $lastMapRow = $wsMap.Cells.Item($wsMap.Rows.Count, 4).End($xlUp).Row
  $fixedMapRows = 0
  for ($r = 1; $r -le $lastMapRow; $r++) {
    $b = [string]$wsMap.Cells.Item($r, 4).Value2
    if ($b -eq $brand) {
      $wsMap.Cells.Item($r, 5).Value2 = $targetOem
      $fixedMapRows++
    }
  }

  $wb.Save()

  [pscustomobject]@{
    status = "fixed"
    backupPath = $backupPath
    fixedCvRows = $fixedRows
    fixedMappingRows = $fixedMapRows
    targetOem = $targetOem
  } | ConvertTo-Json -Depth 5
}
finally {
  if ($wb -ne $null) { $wb.Close($true) | Out-Null }
  if ($excel -ne $null) { $excel.Quit() | Out-Null }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
