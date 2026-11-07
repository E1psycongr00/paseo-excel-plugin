Excel viewer opens Excel workbooks in Paseo's file tab and shows them close to how Excel draws them, on desktop and on phones.

It reads .xlsx, .xlsm, .xltx, .xltm, .xls and .xlt files. The daemon parses the workbook and sends only the rows on screen to the app, so large sheets open quickly and scroll without loading everything.

What it shows:

- Every sheet as a tab, with tab colors. Hidden and very hidden sheets, rows and columns stay hidden until you turn on "숨김", which shows them shaded.
- Formula results as Excel last saved them. The viewer does not recalculate, so a file saved by a tool that writes no cached values shows blank formula cells.
- Cell formatting: number and date formats, fonts, colors, fills, borders (including double, dashed and diagonal), alignment, wrapping, indents, rotated and vertical text, merged cells, text that spills into empty neighbours, column widths, row heights, frozen panes and gridlines.
- Pictures, simple shapes and comment markers. Charts appear as labelled placeholders.
- Page break preview: the print area, manual page breaks as solid blue lines, automatic breaks as dashed lines and page numbers. Automatic breaks are computed by the viewer and can differ from Excel's by about one row or column.

Tap or click a cell to see its value, raw number, comment and link. Zoom with the toolbar buttons or Ctrl and the mouse wheel on desktop.

Files must sit inside the workspace. Password-protected workbooks, Excel 95 and older files, and .xlsb files are not supported. The viewer is read-only and never changes the file.
