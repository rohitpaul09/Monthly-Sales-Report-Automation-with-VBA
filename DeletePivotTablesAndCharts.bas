Attribute VB_Name = "DeletePivotTablesAndCharts"
Sub DeletePivotTablesAndCharts()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim i As Integer

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "MacroButtons" Then
            ' Delete all pivot tables in the worksheet
            For Each pt In ws.PivotTables
                pt.TableRange2.Clear
            Next pt

            ' Delete all chart objects in the worksheet
            For i = ws.ChartObjects.Count To 1 Step -1
                ws.ChartObjects(i).Delete
            Next i
        End If
    Next ws
End Sub

