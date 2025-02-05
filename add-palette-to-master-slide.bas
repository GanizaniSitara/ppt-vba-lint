Function HexToLong(ByVal sHex As String) As Long
    sHex = Replace(sHex, "#", "")
    HexToLong = RGB(CLng("&H" & Mid(sHex, 1, 2)), _
                    CLng("&H" & Mid(sHex, 3, 2)), _
                    CLng("&H" & Mid(sHex, 5, 2)))
End Function

Sub AddSwatchesToSlideMaster()
    Dim hexColors As Variant
    hexColors = Array( _
        "#000000", "#515151", "#D9D9D9", "#E8E8C9", "#FFFFFF", _
        "#FFC9C9", "#FFB05A", "#FFCB05", "#FFFF98", "#FFFF00", _
        "#C3FB5A", "#3F7E37", "#CDF5E8", "#AFFDFD", "#006666", _
        "#007481", "#004750", "#00AEEF", "#0076B6", "#00385D", _
        "#006DE3", "#081276", "#0000FF", "#4C3D6C", "#7A0FF9", _
        "#E1C0E2", "#5C1E5B", "#752157", "#C7237A")
    
    Dim swatchW As Single, swatchH As Single, gap As Single
    swatchW = 0.25 * 28.35: swatchH = swatchW: gap = 1
    
    Dim totalSwatches As Long
    totalSwatches = UBound(hexColors) - LBound(hexColors) + 1
    Dim totalWidth As Single
    totalWidth = totalSwatches * swatchW + (totalSwatches - 1) * gap
    
    Dim margin As Single: margin = 10
    Dim slideW As Single, slideH As Single
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight
    
    Dim startX As Single, startY As Single
    startX = slideW - totalWidth - margin
    startY = slideH - swatchH - margin
    
    Dim sldMaster As Object
    Set sldMaster = ActivePresentation.SlideMaster
    
    Dim i As Long, shp As Shape
    Dim shapeNames() As String
    ReDim shapeNames(0 To totalSwatches - 1)
    
    For i = 0 To totalSwatches - 1
        Set shp = sldMaster.Shapes.AddShape(msoShapeRectangle, _
            startX + i * (swatchW + gap), startY, swatchW, swatchH)
        shp.Fill.ForeColor.RGB = HexToLong(hexColors(i))
        shp.Line.Visible = msoFalse
        shapeNames(i) = shp.Name
    Next i
    
    sldMaster.Shapes.Range(shapeNames).Group
End Sub
