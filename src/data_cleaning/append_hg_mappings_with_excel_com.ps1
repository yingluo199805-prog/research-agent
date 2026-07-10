param(
  [Parameter(Mandatory=$true)][string]$MainPath,
  [Parameter(Mandatory=$true)][string]$MappingsPath
)

$ErrorActionPreference = "Stop"
$mappings = Get-Content -LiteralPath $MappingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $mappings -or $mappings.Count -eq 0) { throw "No mappings found." }

$backupPath = $MainPath -replace '\.xlsx$', ("_preHgMappingBackup_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
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
  # Current workbook order: sheet 6 = 海关匹配字段.
  $ws = $wb.Worksheets.Item(6)
  $xlUp = -4162
  $seriesRow = $ws.Cells.Item($ws.Rows.Count, 1).End($xlUp).Row + 1
  $oemRow = $ws.Cells.Item($ws.Rows.Count, 4).End($xlUp).Row + 1

  foreach ($m in $mappings) {
    $existsSeries = $false
    $existsOem = $false
    $lastSeries = $ws.Cells.Item($ws.Rows.Count, 1).End($xlUp).Row
    for ($r = 3; $r -le $lastSeries; $r++) {
      if ([string]$ws.Cells.Item($r, 1).Value2 -eq [string]$m.brand) {
        $ws.Cells.Item($r, 2).Value2 = $m.series
        $existsSeries = $true
      }
    }
    if (-not $existsSeries) {
      $ws.Cells.Item($seriesRow, 1).Value2 = $m.brand
      $ws.Cells.Item($seriesRow, 2).Value2 = $m.series
      $seriesRow++
    }

    $lastOem = $ws.Cells.Item($ws.Rows.Count, 4).End($xlUp).Row
    for ($r = 3; $r -le $lastOem; $r++) {
      if ([string]$ws.Cells.Item($r, 4).Value2 -eq [string]$m.brand) {
        $ws.Cells.Item($r, 5).Value2 = $m.oem
        $existsOem = $true
      }
    }
    if (-not $existsOem) {
      $ws.Cells.Item($oemRow, 4).Value2 = $m.brand
      $ws.Cells.Item($oemRow, 5).Value2 = $m.oem
      $oemRow++
    }
  }

  $wb.Save()
  [pscustomobject]@{
    status = "written"
    backupPath = $backupPath
    mappingCount = $mappings.Count
    sheetName = $ws.Name
  } | ConvertTo-Json -Depth 5
}
finally {
  if ($wb -ne $null) { $wb.Close($true) | Out-Null }
  if ($excel -ne $null) { $excel.Quit() | Out-Null }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
