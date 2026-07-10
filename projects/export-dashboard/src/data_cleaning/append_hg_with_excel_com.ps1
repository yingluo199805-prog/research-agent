param(
  [Parameter(Mandatory=$true)][string]$MainPath,
  [Parameter(Mandatory=$true)][string]$PayloadPath
)

$ErrorActionPreference = "Stop"

$payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
$rows = $payload.cleaned
if (-not $rows -or $rows.Count -eq 0) { throw "No cleaned rows found." }

$backupPath = $MainPath -replace '\.xlsx$', ("_preHgAppendBackup_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
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
  # Current workbook order: sheet 2 = 海关CV.
  $ws = $wb.Worksheets.Item(2)
  $xlUp = -4162
  $lastRow = $ws.Cells.Item($ws.Rows.Count, 1).End($xlUp).Row

  $existing = 0
  for ($r = 2; $r -le $lastRow; $r++) {
    if ([int]$ws.Cells.Item($r, 1).Value2 -eq 2026 -and [int]$ws.Cells.Item($r, 2).Value2 -eq 4) {
      $existing++
    }
  }
  if ($existing -gt 0) { throw "2026-04 already exists in 海关CV: $existing rows" }

  $rowCount = $rows.Count
  $colCount = 11
  $data = New-Object 'object[,]' $rowCount, $colCount
  for ($i = 0; $i -lt $rowCount; $i++) {
    $rowValues = @($rows[$i])
    for ($j = 0; $j -lt $colCount; $j++) {
      $data[$i, $j] = $rowValues[$j]
    }
  }

  $startRow = $lastRow + 1
  $endRow = $lastRow + $rowCount
  $target = $ws.Range($ws.Cells.Item($startRow, 1), $ws.Cells.Item($endRow, $colCount))
  $target.Value2 = $data

  $wb.Save()
  [pscustomobject]@{
    status = "appended"
    backupPath = $backupPath
    startRow = $startRow
    endRow = $endRow
    rowCount = $rowCount
    sheetName = $ws.Name
  } | ConvertTo-Json -Depth 5
}
finally {
  if ($wb -ne $null) { $wb.Close($true) | Out-Null }
  if ($excel -ne $null) { $excel.Quit() | Out-Null }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
