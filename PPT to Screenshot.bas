Sub SlideScreenshot()

    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        sld.Copy
        sld.Shapes.PasteSpecial ppPasteEnhancedMetafile
        Set shp = sld.Shapes(sld.Shapes.Count)

        With Shp
            .Height = ActivePresentation.PageSetup.SlideHeight
            .Width = ActivePresentation.PageSetup.SlideWidth
            .Left = 0
            .Top = 0
        End With
    Next sld

End Sub
