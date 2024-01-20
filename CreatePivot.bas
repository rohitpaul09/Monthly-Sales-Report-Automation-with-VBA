Attribute VB_Name = "CreatePivot"
Sub CreatePivot()
    Dim ws As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the worksheet is not "MacroButtons"
        If ws.Name <> "MacroButtons" Then
            ' Loop through each pivot table in the worksheet and delete
            Dim ptToDelete As PivotTable
            For Each ptToDelete In ws.PivotTables
                ptToDelete.TableRange2.Clear
            Next ptToDelete

            ' Create a PivotCache from the worksheet's used range
            Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ws.UsedRange)

            ' Add a PivotTable to the worksheet
            Set pt = ws.PivotTables.Add(PivotCache:=pc, TableDestination:=ws.Range("I1"), TableName:="Summary" & ws.Name)

            ' Configure the PivotTable fields
            With pt
                .PivotFields("Branch").Orientation = xlRowField
                .PivotFields("Date").Orientation = xlPageField
                .PivotFields("Category").Orientation = xlColumnField
                .PivotFields("Revenue").Orientation = xlDataField
                .DataBodyRange.NumberFormat = "$0.00"
            End With
        End If
    Next ws
End Sub

