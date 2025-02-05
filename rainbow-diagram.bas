Sub CreateInvertedLayersArcsAndCore_WithNamedColors()
    Dim pptPres As Presentation, pptSlide As Slide, pptShape As Shape
    Dim centerX As Single, centerY As Single
    Dim currentRadius As Single, arcWidth As Single, gap As Single
    Dim layers As Collection, segColl As Collection, seg As Object
    Dim layerColors As Collection, colLayer As Collection
    Dim layerIndex As Integer, segIndex As Integer
    Dim segmentAngle As Single, currentStart As Single
    Dim gapRadius As Single
    Dim dict As Object
    Dim colorDict As Object
    
    ' --- PARAMETERS ---
    centerX = 400
    centerY = 300
    currentRadius = 202
    arcWidth = 28
    gap = 10
    
    Set pptPres = ActivePresentation
    Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutBlank)
    
    ' --- Build the color dictionary from the provided hex values ---
    Set colorDict = CreateObject("Scripting.Dictionary")
    colorDict.Add "Black", "#000000"
    colorDict.Add "DarkGray", "#515151"
    colorDict.Add "LightGray", "#D9D9D9"
    colorDict.Add "LightBeige", "#E8E8C9"
    colorDict.Add "White", "#FFFFFF"
    colorDict.Add "LightPink", "#FFC9C9"
    colorDict.Add "Coral", "#FFB05A"
    colorDict.Add "Mustard", "#FFCB05"
    colorDict.Add "PaleYellow", "#FFFF98"
    colorDict.Add "Yellow", "#FFFF00"
    colorDict.Add "Lime", "#C3FB5A"
    colorDict.Add "ForestGreen", "#3F7E37"
    colorDict.Add "Mint", "#CDF5E8"
    colorDict.Add "Aqua", "#AFFDFD"
    colorDict.Add "DarkCyan", "#006666"
    colorDict.Add "Teal", "#007481"
    colorDict.Add "DeepTeal", "#004750"
    colorDict.Add "SkyBlue", "#00AEEF"
    colorDict.Add "DodgerBlue", "#0076B6"
    colorDict.Add "Navy", "#00385D"
    colorDict.Add "Blue", "#006DE3"
    colorDict.Add "MidnightBlue", "#081276"
    colorDict.Add "PureBlue", "#0000FF"
    colorDict.Add "SlatePurple", "#4C3D6C"
    colorDict.Add "Violet", "#7A0FF9"
    colorDict.Add "Lavender", "#E1C0E2"
    colorDict.Add "Plum", "#5C1E5B"
    colorDict.Add "Burgundy", "#752157"
    colorDict.Add "HotPink", "#C7237A"
    
    ' --- DEFINE LAYERS (in inverted order: outermost = 4 segments, then 3, then 2, then 1) ---
    Set layers = New Collection
    ' Layer 1: 4 segments (each 45°)
    Set segColl = New Collection
    Dim i As Integer
    For i = 1 To 4
         Set dict = CreateObject("Scripting.Dictionary")
         dict.Add "name", "Layer1_" & i
         dict.Add "angle", 180 / 4   ' 45°
         segColl.Add dict
    Next i
    layers.Add segColl
    
    ' Layer 2: 3 segments (each 60°)
    Set segColl = New Collection
    For i = 1 To 3
         Set dict = CreateObject("Scripting.Dictionary")
         dict.Add "name", "Layer2_" & i
         dict.Add "angle", 180 / 3   ' 60°
         segColl.Add dict
    Next i
    layers.Add segColl
    
    ' Layer 3: 2 segments (each 90°)
    Set segColl = New Collection
    For i = 1 To 2
         Set dict = CreateObject("Scripting.Dictionary")
         dict.Add "name", "Layer3_" & i
         dict.Add "angle", 180 / 2   ' 90°
         segColl.Add dict
    Next i
    layers.Add segColl
    
    ' Layer 4: 1 segment (180°)
    Set segColl = New Collection
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "name", "Layer4"
    dict.Add "angle", 180
    segColl.Add dict
    layers.Add segColl
    
    ' --- DEFINE LAYER COLORS using names from the color dictionary ---
    ' For each layer, create a collection of keys that refer to the colorDict.
    Set layerColors = New Collection
    ' Layer 1 (4 segments)
    Set colLayer = New Collection
    colLayer.Add "Teal"         ' from colorDict("Teal") -> "#007481"
    colLayer.Add "DodgerBlue"   ' "#0076B6"
    colLayer.Add "Violet"       ' "#7A0FF9"
    colLayer.Add "HotPink"      ' "#C7237A"
    layerColors.Add colLayer
    ' Layer 2 (3 segments)
    Set colLayer = New Collection
    colLayer.Add "SkyBlue"      ' "#00AEEF"
    colLayer.Add "Blue"         ' "#006DE3"
    colLayer.Add "MidnightBlue" ' "#081276"
    layerColors.Add colLayer
    ' Layer 3 (2 segments)
    Set colLayer = New Collection
    colLayer.Add "ForestGreen"  ' "#3F7E37"
    colLayer.Add "Lime"         ' "#C3FB5A"
    layerColors.Add colLayer
    ' Layer 4 (1 segment)
    Set colLayer = New Collection
    colLayer.Add "Coral"        ' "#FFB05A"
    layerColors.Add colLayer
    
    ' --- DRAW EACH LAYER WITH GAP ARCS ---
    layerIndex = 1
    For Each segColl In layers
         Dim segCount As Integer
         segCount = segColl.Count
         segmentAngle = 180 / segCount
         currentStart = 180
         Set colLayer = layerColors(layerIndex)
         segIndex = 1
         For Each dict In segColl
              Set pptShape = pptSlide.Shapes.AddShape(msoShapeBlockArc, _
                  centerX - currentRadius, centerY - currentRadius, 2 * currentRadius, 2 * currentRadius)
              pptShape.Line.Visible = msoFalse
              pptShape.Fill.ForeColor.RGB = HexToRGB(colorDict(colLayer(segIndex)))
              pptShape.Adjustments.Item(1) = currentStart
              pptShape.Adjustments.Item(2) = currentStart + segmentAngle
              pptShape.Adjustments.Item(3) = (currentRadius - arcWidth) / currentRadius
              pptShape.AlternativeText = dict("name")
              currentStart = currentStart + segmentAngle
              segIndex = segIndex + 1
         Next dict

         ' Draw gap arc for this layer.
         gapRadius = currentRadius - arcWidth
         Set pptShape = pptSlide.Shapes.AddShape(msoShapeBlockArc, _
                  centerX - gapRadius, centerY - gapRadius, 2 * gapRadius, 2 * gapRadius)
         pptShape.Line.Visible = msoFalse
         pptShape.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White gap
         pptShape.Adjustments.Item(1) = 180
         pptShape.Adjustments.Item(2) = 360
         pptShape.Adjustments.Item(3) = (gapRadius - gap) / gapRadius
         currentRadius = gapRadius - gap
         layerIndex = layerIndex + 1
    Next segColl

    ' Draw final gap between innermost layer and core.
    Set pptShape = pptSlide.Shapes.AddShape(msoShapeBlockArc, _
        centerX - currentRadius, centerY - currentRadius, 2 * currentRadius, 2 * currentRadius)
    pptShape.Line.Visible = msoFalse
    pptShape.Fill.ForeColor.RGB = RGB(255, 255, 255)
    pptShape.Adjustments.Item(1) = 180
    pptShape.Adjustments.Item(2) = 360
    pptShape.Adjustments.Item(3) = (currentRadius - gap) / currentRadius
    currentRadius = currentRadius - gap

    ' Draw core (central circle)
    Set pptShape = pptSlide.Shapes.AddShape(msoShapeOval, _
        centerX - currentRadius, centerY - currentRadius, 2 * currentRadius, 2 * currentRadius)
    pptShape.Line.Visible = msoFalse
    pptShape.Fill.ForeColor.RGB = RGB(0, 0, 0)
End Sub

Function HexToRGB(ByVal hexColor As String) As Long
    If Left(hexColor, 1) = "#" Then hexColor = Mid(hexColor, 2)
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Mid(hexColor, 1, 2))
    g = CLng("&H" & Mid(hexColor, 3, 2))
    b = CLng("&H" & Mid(hexColor, 5, 2))
    HexToRGB = RGB(r, g, b)
End Function


