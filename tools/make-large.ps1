# Builds fixtures/large-50k.xlsx (+ .xls is not possible past 65,536 rows, so 60k stays xlsx-only).
param([int]$Rows = 50000, [int]$Cols = 12)
$ErrorActionPreference = "Stop"
$fixtures = Resolve-Path (Join-Path $PSScriptRoot "..\fixtures")
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
$x.ScreenUpdating = $false
try {
  $wb = $x.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $data = New-Object 'object[,]' $Rows, $Cols
  for ($r = 0; $r -lt $Rows; $r++) {
    for ($c = 0; $c -lt $Cols; $c++) {
      if ($c % 3 -eq 0) { $data[$r, $c] = "항목 $r-$c" } else { $data[$r, $c] = [double]($r * $c + 0.5) }
    }
  }
  $range = $ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item($Rows + 1, $Cols))
  $range.Value2 = $data
  $ws.Range($ws.Cells.Item(2, 2), $ws.Cells.Item($Rows + 1, 2)).NumberFormat = "#,##0.00"
  $ws.Range("A1:L1").Value2 = "헤더"
  $ws.Range("A1:L1").Font.Bold = $true
  $ws.Range("A1:L1").Interior.Color = 0xCCE5FF
  $ws.Activate(); $x.ActiveWindow.SplitRow = 1; $x.ActiveWindow.FreezePanes = $true
  $wb.SaveAs((Join-Path $fixtures "large-50k.xlsx"), 51)
  $wb.SaveAs((Join-Path $fixtures "large-50k.xls"), 56)
  $wb.Close($false)
  Write-Output "done"
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
