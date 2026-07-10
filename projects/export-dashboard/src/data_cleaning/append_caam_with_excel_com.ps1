param(
  [Parameter(Mandatory=$true)][string]$MainPath,
  [Parameter(Mandatory=$true)][string]$PayloadPath
)

$ErrorActionPreference = "Stop"

$payload = Get-Content -LiteralPath $PayloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
$rows = $payload.cleaned
if (-not $rows -or $rows.Count -eq 0) {
  throw "No cleaned rows found in payload."
}
Write-Output ("Loaded rows: {0}" -f $rows.Count)

$backupPath = $MainPath -replace '\.xlsx$', ("_preCaamBackup_COM_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
Copy-Item -LiteralPath $MainPath -Destination $backupPath -Force

$excel = $null
$wb = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false
  $excel.EnableEvents = $false

  Write-Output "Opening workbook..."
  $wb = $excel.Workbooks.Open($MainPath)
  Write-Output ("Workbook opened. Sheets: {0}" -f $wb.Worksheets.Count)
  $ws = $wb.Worksheets.Item(1)
  Write-Output ("Using sheet: [{0}]" -f $ws.Name)

  $xlUp = -4162
  $lastRow = $ws.Cells.Item($ws.Rows.Count, 1).End($xlUp).Row
  Write-Output ("Last row: {0}" -f $lastRow)

  $existingApril2026 = 0
  for ($r = 2; $r -le $lastRow; $r++) {
    if ([int]$ws.Cells.Item($r, 1).Value2 -eq 2026 -and [int]$ws.Cells.Item($r, 2).Value2 -eq 4) {
      $existingApril2026++
    }
  }
  if ($existingApril2026 -gt 0) {
    throw "2026-04 already exists in 中汽协CV : $existingApril2026 rows"
  }

  $rowCount = $rows.Count
  $colCount = 10
  $data = New-Object 'object[,]' $rowCount, $colCount
  for ($i = 0; $i -lt $rowCount; $i++) {
    $rowValues = @($rows[$i])
    for ($j = 0; $j -lt $colCount; $j++) {
      $data[$i, $j] = $rowValues[$j]
    }
  }

  $startRow = $lastRow + 1
  $endRow = $lastRow + $rowCount
  Write-Output ("Writing range: {0}:{1}" -f $startRow, $endRow)
  $target = $ws.Range($ws.Cells.Item($startRow, 1), $ws.Cells.Item($endRow, $colCount))
  $target.Value2 = $data

  Write-Output "Saving workbook..."
  $wb.Save()
  Write-Output "Workbook saved."

  [pscustomobject]@{
    status = "appended"
    backupPath = $backupPath
    startRow = $startRow
    endRow = $endRow
    rowCount = $rowCount
  } | ConvertTo-Json -Depth 5
}
finally {
  if ($wb -ne $null) { $wb.Close($true) | Out-Null }
  if ($excel -ne $null) { $excel.Quit() | Out-Null }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
