# Uses installed Excel as the ground truth for parser tests.
# For each input workbook it writes, into fixtures/oracle/:
#   <name>.xls            Excel 97-2003 copy (when the input is .xlsx)
#   <name>.pdf            Excel's own print output
#   <name>.oracle.json    per-sheet truth: displayed text, merges, hidden rows/cols, sizes, page breaks
# Usage: pwsh tools/excel-oracle.ps1 fixtures/sample.xlsx [more files]
param([Parameter(Mandatory = $true, ValueFromRemainingArguments = $true)][string[]]$Files)

$ErrorActionPreference = "Stop"
$outDir = Join-Path $PSScriptRoot "..\fixtures\oracle"
New-Item -ItemType Directory -Force $outDir | Out-Null
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  foreach ($file in $Files) {
    $path = (Resolve-Path $file).Path
    $base = [IO.Path]::GetFileNameWithoutExtension($path)
    $wb = $excel.Workbooks.Open($path, 0, $true)
    if ($path -match "\.xlsx$|\.xlsm$") {
      $xls = Join-Path (Resolve-Path $outDir) "$base.xls"
      $wb.SaveCopyAs((Join-Path $env:TEMP "$base-copy.xlsx"))
      $copy = $excel.Workbooks.Open((Join-Path $env:TEMP "$base-copy.xlsx"))
      $copy.SaveAs($xls, 56)
      $copy.Close($false)
    }
    $pdf = Join-Path (Resolve-Path $outDir) "$base.pdf"
    try { $wb.ExportAsFixedFormat(0, $pdf) } catch { Write-Warning "PDF export failed: $_" }
    $sheets = @()
    foreach ($ws in $wb.Worksheets) {
      $used = $ws.UsedRange
      $lastRow = $used.Row + $used.Rows.Count - 1
      $lastCol = $used.Column + $used.Columns.Count - 1
      $cells = @()
      for ($r = 1; $r -le [Math]::Min($lastRow, 400); $r++) {
        for ($c = 1; $c -le [Math]::Min($lastCol, 60); $c++) {
          $cell = $ws.Cells.Item($r, $c)
          $text = $cell.Text
          if ($text -ne "" -and -not $cell.MergeCells -or ($cell.MergeCells -and $cell.MergeArea.Cells.Item(1, 1).Address() -eq $cell.Address())) {
            if ($text -ne "") { $cells += [ordered]@{ r = $r - 1; c = $c - 1; text = $text; formula = [string]$cell.Formula; fill = [int]$cell.Interior.Color; bold = [bool]$cell.Font.Bold } }
          }
        }
      }
      $hiddenRows = @(); for ($r = 1; $r -le $lastRow; $r++) { if ($ws.Rows.Item($r).Hidden) { $hiddenRows += $r - 1 } }
      $hiddenCols = @(); for ($c = 1; $c -le $lastCol; $c++) { if ($ws.Columns.Item($c).Hidden) { $hiddenCols += $c - 1 } }
      $colPx = @(); for ($c = 1; $c -le [Math]::Min($lastCol + 2, 60); $c++) { $colPx += [Math]::Round($ws.Columns.Item($c).Width / 0.75) }
      $rowPx = @(); for ($r = 1; $r -le [Math]::Min($lastRow + 2, 400); $r++) { $rowPx += [Math]::Round($ws.Rows.Item($r).Height / 0.75) }
      # Page breaks are only computed in page break preview.
      $excel.ScreenUpdating = $false
      $hp = @(); foreach ($b in $ws.HPageBreaks) { $hp += [ordered]@{ at = $b.Location.Row - 1; manual = ($b.Type -eq -4135) } }
      $vp = @(); foreach ($b in $ws.VPageBreaks) { $vp += [ordered]@{ at = $b.Location.Column - 1; manual = ($b.Type -eq -4135) } }
      $sheets += [ordered]@{
        name = $ws.Name; visible = [int]$ws.Visible; pages = $ws.PageSetup.Pages.Count
        zoom = $ws.PageSetup.Zoom; printArea = $ws.PageSetup.PrintArea
        colPx = $colPx; rowPx = $rowPx; hiddenRows = $hiddenRows; hiddenCols = $hiddenCols
        rowBreaks = $hp; colBreaks = $vp; cells = $cells
      }
    }
    $wb.Close($false)
    $json = [ordered]@{ file = [IO.Path]::GetFileName($path); sheets = $sheets } | ConvertTo-Json -Depth 6
    [IO.File]::WriteAllText((Join-Path (Resolve-Path $outDir) "$base.oracle.json"), $json, [Text.UTF8Encoding]::new($false))
    Write-Output "oracle: $base"
  }
} finally {
  $excel.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
