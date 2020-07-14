Option Explicit

Sub FontChange()

Dim sld As Slide
Dim shp As Shape
Dim tbl As Table
Dim lRow As Long
Dim lCol As Long

Dim txt As String
txt = "_고양일산 R"

For Each sld In ActivePresentation.Slides
    For Each shp In sld.Shapes
        If shp.HasTextFrame Then  ' Not all shapes do
            If shp.TextFrame.HasText Then  ' the shape may contain no text
                Dim oSize As Double
                Dim nSize As Double

                oSize = shp.TextFrame.TextRange.Font.Size
                nSize = oSize * 0.9
                
                Dim oSpacing As Double
                Dim nSpacing As Double

                oSpacing = shp.TextFrame.TextRange.Font.Spacing
                nSpacing = oSpacing * 1.1

                With shp.TextFrame.TextRange.Font
                    .Size = nSize
                    .Name = txt
                    .NameFarEast = txt
                    .Spacing = nSpacing
                    '.Bold = False
                    '.Color.RGB = RGB(255, 127, 255)
                End With

                Dim oSpaceWithin As Double
                Dim nSpaceWithin As Double

                oSpaceWithin = shp.TextFrame.TextRange.ParagraphFormat.SpaceWithin
                nSpaceWithin = oSpaceWithin * 1.2

                With shp.TextFrame.TextRange.ParagraphFormat
                    .SpaceWithin = nSpaceWithin
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
    Next shp
Next sld

End Sub
