Option Explicit

Sub FontChange()

Dim sld As Slide
Dim shp As Shape

For Each sld In ActivePresentation.Slides
    For Each shp In sld.Shapes
    If shp.HasTextFrame Then  ' Not all shapes do
    If shp.TextFrame.HasText Then  ' the shape may contain no text
        With shp.TextFrame.TextRange.Font
            '.Size = 12
            .Name = "_고양일산 R"
            .NameFarEast = "_고양일산 R"
            '.Bold = False
            '.Color.RGB = RGB(255, 127, 255)
        End With
    End If
    End If
    Next shp
Next sld
End Sub
