**清空Excel文件中所有Sheet页的F列值**

```vb
Sub ClearColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range

    ' iterate through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' retrieve the range of column F
        lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
        Set rng = ws.Range("F1:F" & lastRow)

        ' clear the contents of column F
        rng.ClearContents
    Next ws
End Sub
```