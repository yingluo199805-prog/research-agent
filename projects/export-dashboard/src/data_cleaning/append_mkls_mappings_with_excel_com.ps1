param(
  [Parameter(Mandatory=$true)][string]$MainPath,
  [Parameter(Mandatory=$true)][string]$MappingsPath
)

$ErrorActionPreference = "Stop"

$mappings = Get-Content -LiteralPath $MappingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $mappings -or $mappings.Count -eq 0) {
  throw "No mappings found."
}

$backupPath = $MainPath -replace '\.xlsx$', ("_preMklsMappingBackup_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
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
  # Current workbook order: sheet 9 = marklines matching fields.
  $ws = $wb.Worksheets.Item(9)

  $xlUp = -4162
  $seriesRow = $ws.Cells.Item($ws.Rows.Count, 10).End($xlUp).Row + 1
  $oemRow = $ws.Cells.Item($ws.Rows.Count, 13).End($xlUp).Row + 1

  foreach ($m in $mappings) {
    $ws.Cells.Item($seriesRow, 10).Value2 = $m.brand
    $ws.Cells.Item($seriesRow, 11).Value2 = $m.series
    $seriesRow++

    $ws.Cells.Item($oemRow, 13).Value2 = $m.brand
    $ws.Cells.Item($oemRow, 14).Value2 = $m.oem
    $oemRow++
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
