Attribute VB_Name = "Copy_Excel_To_PPT"
Sub Copy_Excel_To_PPT()

    Dim PPT_App As Object
    Dim ppt_file As Object
    Dim my_slide As Object
    Set PPT_App = CreateObject("PowerPoint.Application")

    Set ppt_file = PPT_App.Presentations.Add

    Dim sh As Worksheet
    Dim chrt As ChartObject

    For Each sh In ThisWorkbook.Sheets
        If sh.Name <> "MacroButtons" Then
            For Each chrt In sh.ChartObjects
                ' Add a new slide with a blank layout
                Set my_slide = ppt_file.Slides.Add(ppt_file.Slides.Count + 1, ppLayoutBlank)

                ' Add and format title
                Dim titleBox As Object
                Set titleBox = my_slide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                                                          Left:=10, Top:=10, Width:=ppt_file.PageSetup.SlideWidth - 20, Height:=50)
                titleBox.TextFrame.TextRange.Text = sh.Name & " Sales Report"
                titleBox.TextFrame.TextRange.Font.Color = vbWhite
                titleBox.TextFrame.TextRange.Font.Name = "Aptos Black"
                titleBox.Fill.BackColor.RGB = vbBlack
                titleBox.TextFrame.TextRange.Font.Size = 20
                titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                
                ' Copy and paste the chart
                chrt.Copy
                my_slide.Shapes.Paste

                ' Resize and reposition the pasted chart
                With my_slide.Shapes(my_slide.Shapes.Count)
                    .LockAspectRatio = msoTrue
                    .Width = ppt_file.PageSetup.SlideWidth * 0.8 ' 80% of slide width
                    .Height = ppt_file.PageSetup.SlideHeight * 0.8 ' 80% of slide height
                    .Left = (ppt_file.PageSetup.SlideWidth - .Width) / 2 ' Center horizontally
                    .Top = titleBox.Top + titleBox.Height + 10 ' Position below the title box
                End With
            Next chrt
        End If
    Next sh

    PPT_App.Visible = True ' Ensure PowerPoint is visible after the operation

End Sub




