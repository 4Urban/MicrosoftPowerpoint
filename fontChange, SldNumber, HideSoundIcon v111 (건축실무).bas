Attribute VB_Name = "Module1"
Option Explicit

Sub FontChange(txt As String)

    Dim sld As Slide
    Dim shp As Shape
    Dim tbl As Table
    Dim lRow As Long
    Dim lCol As Long

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Name <> "Slide Number" Then
                If shp.HasTextFrame Then  ' Not all shapes do
                    If shp.TextFrame.HasText Then  ' the shape may contain no text
                        With shp.TextFrame.TextRange.Font
                            '.Size = 12
                            .Name = txt
                            .NameFarEast = txt
                            '.Bold = False
                            '.Color.RGB = RGB(255, 127, 255)
                        End With
                    End If
                End If

                If shp.AutoShapeType = msoShapeMixed Then  ' Only Grouped shapes
                    If shp.TextFrame.HasText Then  ' the shape may contain no text
                        With shp.TextFrame.TextRange.Font
                            '.Size = 12
                            .Name = txt
                            .NameFarEast = txt
                            '.Bold = False
                            '.Color.RGB = RGB(255, 127, 255)
                        End With
                    End If
                End If

                If shp.HasTable Then
                Set tbl = shp.Table
                    For lRow = 1 To tbl.Rows.Count
                        For lCol = 1 To tbl.Columns.Count
                            With tbl.Cell(lRow, lCol).Shape.TextFrame.TextRange.Font
                                .Name = txt
                                .NameFarEast = txt
                            End With
                        Next lCol
                    Next lRow
                End If
            End If
        Next shp
    Next sld

End Sub

Sub NewSlideNumber(txt As String)

    Dim dsgn As Design
    Dim shp As Shape
    Dim lSlideWidth As Long, lSlideHeight As Long
    Dim lObjectWidth As Long, lObjectHeight As Long
    Dim X As Long, Y As Long

    lSlideWidth = ActivePresentation.PageSetup.SlideWidth 'get slide horizontal width
    lSlideHeight = ActivePresentation.PageSetup.SlideHeight 'get slide vertical height

    'object width and height
    'lObjectWidth = 34
    'lObjectHeight = 28

    'object width and height (Cm)
    lObjectWidth = cm2Points(1.27)
    lObjectHeight = cm2Points(0.8)

    X = lSlideWidth - lObjectWidth 'calculate horizontal position
    Y = 0 'calculate vertical position

    Dim bln As Boolean

    For Each dsgn In ActivePresentation.Designs
        For Each shp In dsgn.SlideMaster.Shapes
            If shp.Name = "Slide Number" Then
                bln = True
            End If
        Next

        If Not bln Then
            Set shp = dsgn.SlideMaster.Shapes.AddShape( _
                Type:=msoShapeRectangle, _
                Left:=X, _
                Top:=Y, _
                Width:=lObjectWidth, _
                Height:=lObjectHeight)

            With shp
                .Name = "Slide Number"
                .Fill.ForeColor.RGB = RGB(191, 191, 191)
                .Line.Visible = msoFalse
                .TextFrame.VerticalAnchor = msoAnchorMiddle

                With .TextFrame.TextRange
                    With .Font
                        .Size = 12
                        .Name = txt
                        .NameFarEast = txt
                        .Color.RGB = RGB(0, 0, 0)
                    End With
                
                    .ParagraphFormat.Alignment = ppAlignCenter

                    'text = slidenumber
                    .InsertSlideNumber
                End With
            End With
        End If
    Next

End Sub

Function cm2Points(inVal As Single)

    cm2Points = inVal * 28.346

End Function


Sub EmbedTTF()

    Dim nameOfFile As String
    nameOfFile = Application.ActivePresentation.fullName
    Application.ActivePresentation.SaveAs nameOfFile, , msoTrue

End Sub

Sub HideSoundIcon()

    Dim oSl As Slide
    Dim oSh As Shape
    
    For Each oSl In ActivePresentation.Slides
        For Each oSh In oSl.Shapes
            With oSh
                If .Type = msoMedia Then
                    If .MediaType = ppMediaTypeSound Then
                        ' It's a sound; move it
                        .Left = ActivePresentation.PageSetup.SlideWidth + 3
                    End If
                End If
            End With
        Next    ' Shape
    Next    ' Slide

End Sub

Sub Change()

    'Select the Font
    Dim fontType As Integer

    fontType = InputBox( _
        "Please enter the Number of the Font you want." & vbNewLine & _
        vbNewLine & _
        "  1: Noto Sans CJK KR Regular" & vbNewLine & _
        vbNewLine & _
        "  2: SpoqaHanSans-Regular (Recommended)" _
        , "글꼴 선택", 0, , , "", 1)

    'Select the EmbedFont
    Dim embedFont As Integer

    embedFont = InputBox( _
        "Do you want to SAVE with Embedded Fonts? Please enter the number below." & vbNewLine & _
        vbNewLine & _
        "  1: Yes" & vbNewLine & _
        vbNewLine & _
        "  2: No" _
        , "저장 여부 선택", 0, , , "", 1)

    Dim txt As String

    Select Case fontType
        Case 1
            txt = "Noto Sans CJK KR Regular"
        Case 2
            txt = "SpoqaHanSans-Regular"
    End Select

    Call FontChange(txt)
    
    Call NewSlideNumber(txt)
    
    Call HideSoundIcon

    Select Case embedFont
        Case 1
            Call EmbedTTF
    End Select



End Sub
