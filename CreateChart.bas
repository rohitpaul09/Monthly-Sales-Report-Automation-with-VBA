Attribute VB_Name = "CreateChart"
Sub CreateChart()
    Dim ws As Worksheet
    Dim cht As ChartObject

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "MacroButtons" Then
            ' Loop through and delete all existing chart objects in the worksheet
            Dim i As Integer
            For i = ws.ChartObjects.Count To 1 Step -1
                ws.ChartObjects(i).Delete
            Next i

            ' Add a new chart object to the worksheet
            Set cht = ws.ChartObjects.Add(Left:=545, Width:=410, Top:=200, Height:=225)

            ' Set the chart type
            cht.Chart.ChartType = xlColumnClustered

            ' Set the data source range for the chart
            cht.Chart.SetSourceData Source:=ws.Range("I9:M15")

            ' Customize chart properties
            cht.Chart.HasTitle = True
            cht.Chart.ChartTitle.Text = ws.Name
            cht.Chart.Axes(xlCategory, xlPrimary).HasTitle = False
            cht.Chart.Axes(xlValue, xlPrimary).HasTitle = False
        End If
    Next ws
End Sub

