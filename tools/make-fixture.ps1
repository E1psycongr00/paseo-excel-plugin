# Builds fixtures/features.xlsx with Excel itself: every feature the viewer must show, no personal data.
$ErrorActionPreference = "Stop"
$out = Join-Path (Resolve-Path (Join-Path $PSScriptRoot "..\fixtures")) "features.xlsx"
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 4) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $s1 = $wb.Worksheets.Item(1); $s1.Name = "서식"
  $s2 = $wb.Worksheets.Item(2); $s2.Name = "인쇄"
  $s3 = $wb.Worksheets.Item(3); $s3.Name = "숨김시트"
  $s4 = $wb.Worksheets.Item(4); $s4.Name = "매우숨김"

  # Formats, fonts, fills, borders, alignment
  $s1.Range("A1").Value2 = "견적서"; $s1.Range("A1:F1").Merge(); $s1.Range("A1").Font.Size = 18; $s1.Range("A1").Font.Bold = $true
  $s1.Range("A1").HorizontalAlignment = -4108; $s1.Range("A1:F1").Interior.Color = 0xF2E6DA
  $headers = "품목", "수량", "단가", "금액", "날짜", "비고"
  for ($i = 0; $i -lt 6; $i++) { $s1.Cells.Item(3, $i + 1).Value2 = $headers[$i] }
  $s1.Range("A3:F3").Font.Bold = $true; $s1.Range("A3:F3").Interior.Color = 0x7F3F1F; $s1.Range("A3:F3").Font.Color = 0xFFFFFF
  $items = @(@("볼트", 120, 350), @("너트", 80, 120.5), @("와셔", 1500, 12), @("앵커", 3, 15000))
  for ($r = 0; $r -lt $items.Count; $r++) {
    $row = 4 + $r
    $s1.Cells.Item($row, 1).Value2 = $items[$r][0]
    $s1.Cells.Item($row, 2).Value2 = [double]$items[$r][1]
    $s1.Cells.Item($row, 3).Value2 = [double]$items[$r][2]
    $s1.Cells.Item($row, 4).Formula = "=B$row*C$row"
    $s1.Cells.Item($row, 5).Value2 = [double](45600 + $r * 7)
  }
  $s1.Range("B4:B7").NumberFormat = "#,##0"
  $s1.Range("C4:C7").NumberFormat = "#,##0.00"
  $s1.Range("D4:D7").NumberFormatLocal = '#,##0;[빨강]-#,##0'
  $s1.Range("E4:E7").NumberFormat = "yyyy-mm-dd"
  $s1.Range("A8").Value2 = "합계"; $s1.Range("D8").Formula = "=SUM(D4:D7)"; $s1.Range("D8").NumberFormat = "#,##0"; $s1.Range("D8").Font.Bold = $true
  $s1.Range("C9").Value2 = [double]-1234; $s1.Range("C9").NumberFormatLocal = "#,##0;[빨강]-#,##0"
  $s1.Range("D9").Value2 = [double]0.256; $s1.Range("D9").NumberFormat = "0.0%"
  $s1.Range("A3:F8").Borders.LineStyle = 1
  $s1.Range("A8:F8").Borders.Item(8).Weight = -4138
  $s1.Range("A8:F8").Borders.Item(9).LineStyle = -4119
  $s1.Range("F4").Value2 = "이 비고는 아주 길어서 옆 칸으로 넘칩니다"
  $s1.Range("F5").Value2 = "줄바꿈이 켜진 긴 비고 문장입니다"; $s1.Range("F5").WrapText = $true
  $s1.Range("G11").Value2 = "세로"; $s1.Range("G11").Orientation = -4166
  $s1.Range("H11").Value2 = "회전45"; $s1.Range("H11").Orientation = 45
  $s1.Range("B11:C13").Merge(); $s1.Range("B11").Value2 = "세로 병합"; $s1.Range("B11").VerticalAlignment = -4108; $s1.Range("B11").HorizontalAlignment = -4108
  $s1.Range("B11:C13").Borders.Item(7).Weight = -4138; $s1.Range("B11:C13").Borders.Item(10).Weight = -4138
  $s1.Range("B11:C13").Borders.Item(8).Weight = -4138; $s1.Range("B11:C13").Borders.Item(9).Weight = -4138
  $s1.Range("E11").Value2 = "대각선"; $s1.Range("E11").Borders.Item(5).LineStyle = 1
  $s1.Range("A15").Value2 = "숨긴 행 다음"
  $s1.Rows.Item(14).Hidden = $true
  $s1.Range("I3").Value2 = "숨긴 열 다음"
  $s1.Columns.Item(8).Hidden = $false
  $s1.Columns.Item("J").Hidden = $true
  $s1.Range("J3").Value2 = "숨겨진 값"
  [void]$s1.Range("B4").AddComment("수량은 개 단위입니다")
  $s1.Hyperlinks.Add($s1.Range("A17"), "https://example.com", [Type]::Missing, [Type]::Missing, "예시 링크") | Out-Null
  $s1.Columns.Item(1).ColumnWidth = 12; $s1.Columns.Item(6).ColumnWidth = 18
  $s1.Rows.Item(1).RowHeight = 30
  $s1.Tab.Color = 0x00C0FF
  $s1.Activate(); $x.ActiveWindow.SplitRow = 3; $x.ActiveWindow.SplitColumn = 1; $x.ActiveWindow.FreezePanes = $true

  # Printing: print area, manual + automatic breaks, page break preview
  for ($r = 1; $r -le 120; $r++) { $s2.Cells.Item($r, 1).Value2 = "행 $r"; $s2.Cells.Item($r, 2).Value2 = [double]($r * 3); $s2.Cells.Item($r, 3).Formula = "=B$r*2" }
  $s2.PageSetup.PrintArea = "`$A`$1:`$E`$110"
  $s2.HPageBreaks.Add($s2.Range("A30")) | Out-Null
  $s2.PageSetup.PrintTitleRows = "`$1:`$1"
  $s2.Activate(); $x.ActiveWindow.View = 2

  $s3.Range("A1").Value2 = "숨긴 시트 내용"; $s3.Visible = 0
  $s4.Range("A1").Value2 = "매우 숨긴 시트 내용"; $s4.Visible = 2
  $s1.Activate()
  $wb.SaveAs($out, 51)
  $wb.Close($false)
  Write-Output "wrote $out"
} finally {
  $x.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
