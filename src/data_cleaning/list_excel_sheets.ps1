param([Parameter(Mandatory=$true)][string]$Path)

$ErrorActionPreference = "Stop"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($Path)
  try {
    for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
      $name = $wb.Worksheets.Item($i).Name
      Write-Output ("[{0}] len={1}" -f $name, $name.Length)
    }
  }
  finally {
    $wb.Close($false) | Out-Null
  }
}
finally {
  $excel.Quit() | Out-Null
}
