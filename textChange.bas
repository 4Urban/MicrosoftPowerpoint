Option Explicit

Sub TextChange()

    Dim dsgn As Design
    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim lRow As Long
    Dim lCol As Long

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Name <> "Slide Number" Then
                If shp.HasTable Then
                Set tbl = shp.Table
                    For lRow = 1 To tbl.Rows.Count
                        For lCol = 1 To tbl.Rows.Count
                            'If lCol = 2 Then
                            '    tbl.Cell(lRow, lCol).Shape.TextFrame.TextRange.Font.Italic = False
                            'End If

                            If lCol = 3 Then
                                Dim x As String
                                Dim y As String

                                x = tbl.Cell(lRow, lCol).Shape.TextFrame.TextRange.Text

                                Select Case x
                                    Case "6.25%"
                                        tbl.Cell(lRow, lCol).Shape.TextFrame.TextRange.Text = "6.67%"
                                    Case "6.67%"
                                        tbl.Cell(lRow, lCol).Shape.TextFrame.TextRange.Text = "6.25%"
                                    Case Else
                                End Select
                            End If
                        Next lCol
                    Next lRow
                End If
            End If
        Next shp
    Next sld

End Sub


Option Explicit

Sub EraseTile()

    Dim dsgn As Design
    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim lRow As Long
    Dim lCol As Long

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If Not shp.HasTable Then
                shp.Fill.Transparency = 1
            End If
        Next shp
    Next sld

End Sub

Sub ChangeTable()

    Dim dsgn As Design
    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim cll As Cell
    Dim txt As TextRange
    Dim lRow As Long
    Dim lCol As Long

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If Not shp.TextFrame.HasText And Not shp.HasTable Then
                shp.Delete

            ElseIf shp.HasTable Then
            Set tbl = shp.Table
                For lRow = 1 To tbl.Rows.Count
                    For lCol = 1 To tbl.Columns.Count
                        Set cll = tbl.Cell(lRow, lCol)
                        Set txt = cll.Shape.TextFrame.TextRange

                        If lRow = 1 Or lCol = 1 Then
                            cll.Shape.TextFrame.Ruler.Levels(1).LeftMargin = 0
                            cll.Shape.TextFrame.Ruler.Levels(1).FirstMargin = 0
                            
                            With txt.ParagraphFormat
                                .Alignment = ppAlignCenter
                                .BaseLineAlignment = ppBaselineAlignCenter
                                .LineRuleBefore = False
                                .SpaceBefore = 0
                            End With

                            With txt.Font
                                .Bold = True
                            End With

                            With cll
                                .Shape.Fill.ForeColor.RGB = RGB(205, 205, 205)
                            End With
                        ElseIf cll.Shape.Fill.ForeColor.RGB = RGB(222, 222, 222) Then
                            With txt.Font
                                .Bold = False
                                .Name = "netmarble Bold"
                                .NameFarEast = "netmarble Bold"
                                .Color.RGB = RGB(0, 0, 0)
                            End With

                            With cll
                                .Shape.Fill.ForeColor.RGB = RGB(255, 210, 17)
                            End With

                        End If
                        With cll
                            .Borders(ppBorderBottom).ForeColor.RGB = RGB(0, 0, 0)
                            .Borders(ppBorderLeft).ForeColor.RGB = RGB(0, 0, 0)
                            .Borders(ppBorderRight).ForeColor.RGB = RGB(0, 0, 0)
                            .Borders(ppBorderTop).ForeColor.RGB = RGB(0, 0, 0)
                        End With
                    Next lCol
                Next lRow
            End If
        Next shp
    Next sld

End Sub