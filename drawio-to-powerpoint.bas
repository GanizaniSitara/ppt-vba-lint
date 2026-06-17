Attribute VB_Name = "Module1"
Sub GenerateDiagramInCurrentSlide(xmlPath As String, _
    excludedLayerNames As Variant, backLayerNames As Variant, _
    lineShadeDiff As Long, useStandardColors As Boolean)

    ' Define standard hex colors array.
    Dim hexColors As Variant
    hexColors = Array( _
        "#FFFF98", "#C3FB5A", "#081276", "#AFFDFD", "#5C1E5B", "#000000", "#E8E8C9", _
        "#FFFF00", "#007481", "#00385D", "#0076B6", "#4C3D6C", "#E1C0E2", "#D9D9D9", _
        "#FFB05A", "#006666", "#CDF5E8", "#006DE3", "#FFC9C9", "#7A0FF9", "#515151", _
        "#FFCB05", "#004750", "#3F7E37", "#0000FF", "#C7237A", "#752157", "#FFFFFF" _
    )
    
    Dim pptPres As Presentation, pptSlide As Slide
    Set pptPres = ActivePresentation
    Set pptSlide = ActiveWindow.View.Slide

    Dim xmlDoc As New MSXML2.DOMDocument60
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    If Not xmlDoc.Load(xmlPath) Then
        Debug.Print "Error loading XML: " & xmlDoc.ParseError.reason
        Exit Sub
    End If

    ' Dictionary mapping drawio shape IDs to PPT shapes.
    Dim shapeMap As Object
    Set shapeMap = CreateObject("Scripting.Dictionary")
    Dim generatedShapes As Collection
    Set generatedShapes = New Collection
    
    Dim vertexNodes As MSXML2.IXMLDOMNodeList
    Set vertexNodes = xmlDoc.SelectNodes("//mxCell[@vertex='1']")
    Dim edgeNodes As MSXML2.IXMLDOMNodeList
    Set edgeNodes = xmlDoc.SelectNodes("//mxCell[@edge='1']")
    
    Dim shp As Shape
    Dim xPos As Single, yPos As Single, widthVal As Single, heightVal As Single
    Dim styleStr As String, fillColorStr As String, strokeColorStr As String, pos As Long
    Dim htmlDoc As New MSHTML.HTMLDocument, labelText As String
    Dim geoNode As MSXML2.IXMLDOMNode, node As MSXML2.IXMLDOMNode
    Dim shapeType As MsoAutoShapeType
    Dim currentId As String

    '--- Process Vertex Nodes ---
    For Each node In vertexNodes
        ' Exclude nodes if any ancestor's value matches an excluded layer.
        If HasAnyAncestorValue(node, excludedLayerNames, xmlDoc) Then GoTo NextVertex

        Set geoNode = node.SelectSingleNode("mxGeometry")
        If geoNode Is Nothing Then GoTo NextVertex
        
        ' Get the absolute coordinates by traversing the "parent" chain.
        Dim absX As Single, absY As Single
        Call GetAbsoluteCoordinates(xmlDoc, node, absX, absY)
        xPos = absX
        yPos = absY

        ' Use width and height from the node's own geometry.
        If Not geoNode.Attributes.getNamedItem("width") Is Nothing Then
            widthVal = Val(geoNode.Attributes.getNamedItem("width").Text)
        Else
            widthVal = 50
        End If
        If Not geoNode.Attributes.getNamedItem("height") Is Nothing Then
            heightVal = Val(geoNode.Attributes.getNamedItem("height").Text)
        Else
            heightVal = 50
        End If

        ' --- Modified Label Extraction ---
        Dim labelFromParent As String, labelFromValue As String
        labelFromParent = ""
        labelFromValue = ""
        If Not node.ParentNode Is Nothing Then
            If LCase(node.ParentNode.nodeName) = "object" Then
                If Not node.ParentNode.Attributes.getNamedItem("label") Is Nothing Then
                    labelFromParent = DecodeHtmlText(node.ParentNode.Attributes.getNamedItem("label").Text, htmlDoc)
                End If
            End If
        End If
        If Not node.Attributes.getNamedItem("value") Is Nothing Then
            labelFromValue = DecodeHtmlText(node.Attributes.getNamedItem("value").Text, htmlDoc)
        End If
        If labelFromParent <> "" And labelFromValue <> "" Then
            labelText = labelFromParent & vbCrLf & labelFromValue
        Else
            labelText = labelFromParent & labelFromValue
        End If

        If InStr(labelText, "%") > 0 Then
            labelText = FormatLabelText(node, labelText)
        End If
        If Not node.Attributes.getNamedItem("description") Is Nothing Then
            labelText = AppendNonEmptyLine(labelText, DecodeHtmlText(node.Attributes.getNamedItem("description").Text, htmlDoc))
        End If

        If Not node.Attributes.getNamedItem("style") Is Nothing Then
            styleStr = node.Attributes.getNamedItem("style").Text
        Else
            styleStr = ""
        End If

        ' Determine shape type.
        If InStr(1, styleStr, "ellipse", vbTextCompare) > 0 Then
            shapeType = msoShapeOval
        ElseIf InStr(1, styleStr, "rounded=1", vbTextCompare) > 0 Then
            shapeType = msoShapeRoundedRectangle
        Else
            shapeType = msoShapeRectangle
        End If

        Set shp = pptSlide.Shapes.AddShape(shapeType, xPos, yPos, widthVal, heightVal)
        generatedShapes.Add shp
        shp.TextFrame2.TextRange.Text = labelText
        shp.TextFrame2.WordWrap = msoFalse

        If shapeType = msoShapeOval Then
            shp.TextFrame2.TextRange.Font.Size = 4
        Else
            shp.TextFrame2.TextRange.Font.Size = 6
            shp.TextFrame2.VerticalAnchor = msoAnchorTop
            shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        End If
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)

        If HasAnyAncestorValue(node, backLayerNames, xmlDoc) Then
            shp.ZOrder msoSendToBack
        ElseIf shapeType = msoShapeRectangle Or shapeType = msoShapeRoundedRectangle Then
            shp.ZOrder msoSendToBack
        End If

        fillColorStr = GetStyleValue(styleStr, "fillColor")
        If fillColorStr <> "" Then
            If LCase(fillColorStr) = "none" Then
                shp.Fill.Visible = msoFalse
            ElseIf IsHexColor(fillColorStr) Then
                If useStandardColors Then
                    fillColorStr = GetClosestStandardColor(fillColorStr, hexColors)
                End If
                Dim fillRGB As Long
                fillRGB = HexToRGB(fillColorStr)
                If (shapeType = msoShapeRectangle Or shapeType = msoShapeRoundedRectangle) And fillRGB = RGB(0, 0, 0) Then
                    shp.Fill.Visible = msoFalse
                Else
                    shp.Fill.ForeColor.RGB = fillRGB
                End If
            End If
        End If

        If shp.Fill.Visible = msoFalse Then shp.ZOrder msoBringToFront

        strokeColorStr = GetStyleValue(styleStr, "strokeColor")
        If strokeColorStr <> "" Then
            If LCase(strokeColorStr) = "none" Then
                shp.Line.Visible = msoFalse
            ElseIf IsHexColor(strokeColorStr) Then
                Dim baseLineRGB As Long
                baseLineRGB = HexToRGB(strokeColorStr)
                shp.Line.ForeColor.RGB = DarkenColorRGB(baseLineRGB, lineShadeDiff)
            End If
        End If
        shp.Line.Weight = 0.5

        If Not node.Attributes.getNamedItem("id") Is Nothing Then
            currentId = node.Attributes.getNamedItem("id").Text
            If shapeMap.Exists(currentId) Then
                shapeMap.Remove currentId
                shapeMap.Add currentId, shp
            Else
                shapeMap.Add currentId, shp
            End If
        End If

NextVertex:
    Next node

    '--- Process Edge (Connector) Nodes ---
    Dim sourcePtNode As MSXML2.IXMLDOMNode, targetPtNode As MSXML2.IXMLDOMNode
    Dim sourceX As Single, sourceY As Single, targetX As Single, targetY As Single
    Dim edgeStyle As String
    Dim sourceId As String, targetId As String
    For Each node In edgeNodes
        If HasAnyAncestorValue(node, excludedLayerNames, xmlDoc) Then GoTo NextEdge

        sourceId = ""
        targetId = ""
        If Not node.Attributes.getNamedItem("source") Is Nothing Then
            sourceId = node.Attributes.getNamedItem("source").Text
        End If
        If Not node.Attributes.getNamedItem("target") Is Nothing Then
            targetId = node.Attributes.getNamedItem("target").Text
        End If

        Dim sourceShape As Shape, targetShape As Shape
        Set sourceShape = Nothing
        Set targetShape = Nothing
        If sourceId <> "" Then
            If shapeMap.Exists(sourceId) Then Set sourceShape = shapeMap(sourceId)
        End If
        If targetId <> "" Then
            If shapeMap.Exists(targetId) Then Set targetShape = shapeMap(targetId)
        End If

        Set geoNode = node.SelectSingleNode("mxGeometry")
        If geoNode Is Nothing And sourceShape Is Nothing And targetShape Is Nothing Then GoTo NextEdge

        Set sourcePtNode = Nothing
        Set targetPtNode = Nothing
        If Not geoNode Is Nothing Then
            Set sourcePtNode = geoNode.SelectSingleNode("mxPoint[@as='sourcePoint']")
            Set targetPtNode = geoNode.SelectSingleNode("mxPoint[@as='targetPoint']")
        End If

        If Not sourcePtNode Is Nothing Then
            sourceX = GetPointCoordinate(sourcePtNode, "x", 0)
            sourceY = GetPointCoordinate(sourcePtNode, "y", 0)
        ElseIf Not sourceShape Is Nothing Then
            sourceX = ShapeCenterX(sourceShape)
            sourceY = ShapeCenterY(sourceShape)
        Else
            sourceX = 0: sourceY = 0
        End If

        If Not targetPtNode Is Nothing Then
            targetX = GetPointCoordinate(targetPtNode, "x", 0)
            targetY = GetPointCoordinate(targetPtNode, "y", 0)
        ElseIf Not targetShape Is Nothing Then
            targetX = ShapeCenterX(targetShape)
            targetY = ShapeCenterY(targetShape)
        Else
            targetX = 0: targetY = 0
        End If

        Dim conn As Shape
        Set conn = pptSlide.Shapes.AddConnector(msoConnectorStraight, sourceX, sourceY, targetX, targetY)
        generatedShapes.Add conn
        conn.Line.ForeColor.RGB = RGB(0, 0, 0)
        conn.Line.Weight = 0.5

        If Not sourceShape Is Nothing Then
            ConnectConnectorEndpoint conn, True, sourceShape, targetX, targetY
        End If
        If Not targetShape Is Nothing Then
            ConnectConnectorEndpoint conn, False, targetShape, sourceX, sourceY
        End If

        edgeStyle = ""
        If Not node.Attributes.getNamedItem("style") Is Nothing Then
            edgeStyle = node.Attributes.getNamedItem("style").Text
            If InStr(1, edgeStyle, "dashed", vbTextCompare) > 0 Then
                conn.Line.DashStyle = msoLineDash
            ElseIf InStr(1, edgeStyle, "dotted", vbTextCompare) > 0 Then
                conn.Line.DashStyle = msoLineRoundDot
            Else
                conn.Line.DashStyle = msoLineSolid
            End If
        Else
            conn.Line.DashStyle = msoLineSolid
        End If

        strokeColorStr = GetStyleValue(edgeStyle, "strokeColor")
        If strokeColorStr <> "" Then
            If LCase(strokeColorStr) = "none" Then
                conn.Line.Visible = msoFalse
            ElseIf IsHexColor(strokeColorStr) Then
                conn.Line.ForeColor.RGB = HexToRGB(strokeColorStr)
            End If
        End If
        ApplyConnectorArrowheads conn, edgeStyle
NextEdge:
    Next node

    If generatedShapes.Count = 0 Then
        Debug.Print "No shapes generated from XML."
        Exit Sub
    End If

    '--- Determine Bounding Box for Scaling ---
    Dim shpItem As Shape
    Dim bbMinX As Single, bbMinY As Single, bbMaxX As Single, bbMaxY As Single
    bbMinX = 1E+30: bbMinY = 1E+30: bbMaxX = -1E+30: bbMaxY = -1E+30
    For Each shpItem In generatedShapes
        Dim lVal As Single, tVal As Single, rVal As Single, bVal As Single
        lVal = shpItem.Left
        tVal = shpItem.Top
        rVal = shpItem.Left + shpItem.Width
        bVal = shpItem.Top + shpItem.Height
        If lVal < bbMinX Then bbMinX = lVal
        If tVal < bbMinY Then bbMinY = tVal
        If rVal > bbMaxX Then bbMaxX = rVal
        If bVal > bbMaxY Then bbMaxY = bVal
    Next shpItem
    
    Dim diagramWidth As Single, diagramHeight As Single
    diagramWidth = bbMaxX - bbMinX
    diagramHeight = bbMaxY - bbMinY
    If diagramWidth <= 0 Or diagramHeight <= 0 Then
        Debug.Print "Generated diagram has no measurable bounds."
        Exit Sub
    End If

    Dim slideWidth As Single, slideHeight As Single
    slideWidth = pptPres.PageSetup.SlideWidth
    slideHeight = pptPres.PageSetup.SlideHeight

    '--- Scale Diagram to Fit 16:9 Slide (with Margin) ---
    Dim margin As Single: margin = 20
    Dim scaleFactorX As Single, scaleFactorY As Single, scaleFactor As Single
    scaleFactorX = (slideWidth - 2 * margin) / diagramWidth
    scaleFactorY = (slideHeight - 2 * margin) / diagramHeight
    scaleFactor = scaleFactorX
    If scaleFactorY < scaleFactor Then scaleFactor = scaleFactorY

    Dim offX As Single, offY As Single
    offX = (slideWidth - diagramWidth) / 2 - bbMinX
    offY = (slideHeight - diagramHeight) / 2 - bbMinY

    Dim scaled As Boolean: scaled = False
    If scaleFactor < 1 Then
        scaled = True
        For Each shpItem In generatedShapes
            shpItem.Left = (shpItem.Left - bbMinX) * scaleFactor + margin
            shpItem.Top = (shpItem.Top - bbMinY) * scaleFactor + margin
            shpItem.Width = shpItem.Width * scaleFactor
            shpItem.Height = shpItem.Height * scaleFactor
            If shpItem.HasTextFrame Then
                If shpItem.AutoShapeType = msoShapeOval Then
                    shpItem.TextFrame2.TextRange.Font.Size = 4
                Else
                    shpItem.TextFrame2.TextRange.Font.Size = 6
                End If
                shpItem.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                shpItem.TextFrame2.WordWrap = msoFalse
            End If
            shpItem.Line.Weight = 0.5
        Next shpItem
    Else
        For Each shpItem In generatedShapes
            shpItem.Left = shpItem.Left + offX
            shpItem.Top = shpItem.Top + offY
        Next shpItem
    End If

    If scaled Then
        Debug.Print "Diagram scaled to fit slide."
    Else
        Debug.Print "Diagram fits slide; positions adjusted."
    End If
End Sub

'------------------------------------------------------------
' This function recursively accumulates the "x" and "y" offsets
' from the current node and all its ancestors (via the "parent" attribute)
' to compute absolute coordinates.
Function GetAbsoluteCoordinates(xmlDoc As MSXML2.DOMDocument60, node As MSXML2.IXMLDOMNode, ByRef absX As Single, ByRef absY As Single)
    absX = 0: absY = 0
    Dim currentNode As MSXML2.IXMLDOMNode
    Set currentNode = node
    Dim geo As MSXML2.IXMLDOMNode
    Dim parentId As String
    Do While Not currentNode Is Nothing
        Set geo = currentNode.SelectSingleNode("mxGeometry")
        If Not geo Is Nothing Then
            If Not geo.Attributes.getNamedItem("x") Is Nothing Then
                absX = absX + Val(geo.Attributes.getNamedItem("x").Text)
            End If
            If Not geo.Attributes.getNamedItem("y") Is Nothing Then
                absY = absY + Val(geo.Attributes.getNamedItem("y").Text)
            End If
        End If
        If Not currentNode.Attributes Is Nothing Then
            If Not currentNode.Attributes.getNamedItem("parent") Is Nothing Then
                parentId = currentNode.Attributes.getNamedItem("parent").Text
                If parentId = "0" Then Exit Do
                Set currentNode = xmlDoc.SelectSingleNode("//*[@id='" & parentId & "']")
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
End Function

'------------------------------------------------------------
Function DecodeHtmlText(ByVal value As String, ByRef htmlDoc As MSHTML.HTMLDocument) As String
    On Error GoTo PlainText
    htmlDoc.body.innerHTML = value
    DecodeHtmlText = Trim(htmlDoc.body.innerText)
    Exit Function
PlainText:
    DecodeHtmlText = Trim(value)
End Function

'------------------------------------------------------------
Function AppendNonEmptyLine(ByVal baseText As String, ByVal newText As String) As String
    newText = Trim(newText)
    If newText = "" Then
        AppendNonEmptyLine = baseText
    ElseIf Trim(baseText) = "" Then
        AppendNonEmptyLine = newText
    Else
        AppendNonEmptyLine = baseText & vbCrLf & newText
    End If
End Function

'------------------------------------------------------------
Function GetLabelAttribute(ByVal node As MSXML2.IXMLDOMNode, ByVal attributeName As String) As String
    If Not node.Attributes Is Nothing Then
        If Not node.Attributes.getNamedItem(attributeName) Is Nothing Then
            GetLabelAttribute = node.Attributes.getNamedItem(attributeName).Text
            Exit Function
        End If
    End If

    If Not node.ParentNode Is Nothing Then
        If LCase(node.ParentNode.nodeName) = "object" Then
            If Not node.ParentNode.Attributes Is Nothing Then
                If Not node.ParentNode.Attributes.getNamedItem(attributeName) Is Nothing Then
                    GetLabelAttribute = node.ParentNode.Attributes.getNamedItem(attributeName).Text
                    Exit Function
                End If
            End If
        End If
    End If

    GetLabelAttribute = ""
End Function

'------------------------------------------------------------
Function FormatLabelText(ByVal node As MSXML2.IXMLDOMNode, ByVal lbl As String) As String
    Dim parts() As String, i As Long, result As String, placeholder As String, attrValue As String
    parts = Split(lbl, "%")
    result = ""
    For i = 0 To UBound(parts)
        If i Mod 2 = 0 Then
            result = result & parts(i)
        Else
            placeholder = Trim(parts(i))
            attrValue = GetLabelAttribute(node, placeholder)
            If attrValue = "" Then
                result = result & "%" & parts(i) & "%"
            Else
                result = result & attrValue
            End If
        End If
    Next i
    FormatLabelText = Trim(result)
End Function

'------------------------------------------------------------
Function GetPointCoordinate(ByVal pointNode As MSXML2.IXMLDOMNode, ByVal coordinateName As String, ByVal defaultValue As Single) As Single
    If Not pointNode.Attributes Is Nothing Then
        If Not pointNode.Attributes.getNamedItem(coordinateName) Is Nothing Then
            GetPointCoordinate = Val(pointNode.Attributes.getNamedItem(coordinateName).Text)
            Exit Function
        End If
    End If
    GetPointCoordinate = defaultValue
End Function

'------------------------------------------------------------
Function ShapeCenterX(ByVal shp As Shape) As Single
    ShapeCenterX = shp.Left + shp.Width / 2
End Function

'------------------------------------------------------------
Function ShapeCenterY(ByVal shp As Shape) As Single
    ShapeCenterY = shp.Top + shp.Height / 2
End Function

'------------------------------------------------------------
Function PreferredConnectionSite(ByVal shp As Shape, ByVal towardX As Single, ByVal towardY As Single) As Long
    On Error Resume Next
    Dim siteCount As Long
    siteCount = shp.ConnectionSiteCount
    If Err.Number <> 0 Or siteCount < 1 Then
        Err.Clear
        PreferredConnectionSite = 1
        Exit Function
    End If
    On Error GoTo 0

    If siteCount < 4 Then
        PreferredConnectionSite = 1
        Exit Function
    End If

    Dim dx As Single, dy As Single
    dx = towardX - ShapeCenterX(shp)
    dy = towardY - ShapeCenterY(shp)
    If Abs(dx) > Abs(dy) Then
        If dx >= 0 Then
            PreferredConnectionSite = 2
        Else
            PreferredConnectionSite = 4
        End If
    Else
        If dy >= 0 Then
            PreferredConnectionSite = 3
        Else
            PreferredConnectionSite = 1
        End If
    End If

    If PreferredConnectionSite > siteCount Then PreferredConnectionSite = 1
End Function

'------------------------------------------------------------
Sub ConnectConnectorEndpoint(ByVal conn As Shape, ByVal connectBegin As Boolean, ByVal targetShape As Shape, ByVal towardX As Single, ByVal towardY As Single)
    On Error GoTo Done
    Dim siteIndex As Long
    siteIndex = PreferredConnectionSite(targetShape, towardX, towardY)
    If connectBegin Then
        conn.ConnectorFormat.BeginConnect targetShape, siteIndex
    Else
        conn.ConnectorFormat.EndConnect targetShape, siteIndex
    End If
    conn.RerouteConnections
Done:
End Sub

'------------------------------------------------------------
Function GetStyleValue(ByVal styleStr As String, ByVal styleName As String) As String
    Dim parts As Variant, item As Variant, eqPos As Long
    Dim key As String, value As String
    parts = Split(styleStr, ";")
    For Each item In parts
        eqPos = InStr(1, CStr(item), "=", vbBinaryCompare)
        If eqPos > 0 Then
            key = Trim(Left(CStr(item), eqPos - 1))
            value = Trim(Mid(CStr(item), eqPos + 1))
            If StrComp(key, styleName, vbTextCompare) = 0 Then
                GetStyleValue = value
                Exit Function
            End If
        End If
    Next item
    GetStyleValue = ""
End Function

'------------------------------------------------------------
Function IsHexColor(ByVal colorValue As String) As Boolean
    Dim i As Long, ch As String
    colorValue = Replace(Trim(colorValue), "#", "")
    If Len(colorValue) <> 6 Then
        IsHexColor = False
        Exit Function
    End If
    For i = 1 To 6
        ch = Mid(colorValue, i, 1)
        If InStr(1, "0123456789ABCDEFabcdef", ch, vbBinaryCompare) = 0 Then
            IsHexColor = False
            Exit Function
        End If
    Next i
    IsHexColor = True
End Function

'------------------------------------------------------------
Sub ApplyConnectorArrowheads(ByVal conn As Shape, ByVal styleStr As String)
    Dim startArrow As String, endArrow As String
    startArrow = LCase(GetStyleValue(styleStr, "startArrow"))
    endArrow = LCase(GetStyleValue(styleStr, "endArrow"))

    If startArrow <> "" And startArrow <> "none" Then
        conn.Line.BeginArrowheadStyle = msoArrowheadTriangle
    Else
        conn.Line.BeginArrowheadStyle = msoArrowheadNone
    End If

    If endArrow <> "" And endArrow <> "none" Then
        conn.Line.EndArrowheadStyle = msoArrowheadTriangle
    Else
        conn.Line.EndArrowheadStyle = msoArrowheadNone
    End If
End Sub

'------------------------------------------------------------
Function HasAnyAncestorValue(ByVal node As MSXML2.IXMLDOMNode, valueArray As Variant, _
    ByVal xmlDoc As MSXML2.DOMDocument60) As Boolean
    Dim parentId As String, parentNode As MSXML2.IXMLDOMNode, item As Variant
    HasAnyAncestorValue = False
    Do While Not node Is Nothing
        If Not node.Attributes Is Nothing Then
            If Not node.Attributes.getNamedItem("parent") Is Nothing Then
                parentId = node.Attributes.getNamedItem("parent").Text
                Set parentNode = xmlDoc.SelectSingleNode("//*[@id='" & parentId & "']")
                If Not parentNode Is Nothing Then
                    If Not parentNode.Attributes.getNamedItem("value") Is Nothing Then
                        For Each item In valueArray
                            If LCase(Trim(parentNode.Attributes.getNamedItem("value").Text)) = LCase(Trim(item)) Then
                                HasAnyAncestorValue = True
                                Exit Function
                            End If
                        Next item
                    End If
                End If
                Set node = parentNode
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
End Function

'------------------------------------------------------------
Function GetClosestStandardColor(currentColor As String, colorArray As Variant) As String
    Dim currentRGB As Long, currentR As Long, currentG As Long, currentB As Long
    currentRGB = HexToRGB(currentColor)
    currentR = currentRGB And &HFF
    currentG = (currentRGB \ &H100) And &HFF
    currentB = (currentRGB \ &H10000) And &HFF
    
    Dim bestColor As String, bestDiff As Double, diff As Double
    bestDiff = 1E+30
    Dim i As Long, candidateColor As String, candidateRGB As Long
    Dim candR As Long, candG As Long, candB As Long
    For i = LBound(colorArray) To UBound(colorArray)
        candidateColor = colorArray(i)
        candidateRGB = HexToRGB(candidateColor)
        candR = candidateRGB And &HFF
        candG = (candidateRGB \ &H100) And &HFF
        candB = (candidateRGB \ &H10000) And &HFF
        diff = Sqr((currentR - candR) ^ 2 + (currentG - candG) ^ 2 + (currentB - candB) ^ 2)
        If diff < bestDiff Then
            bestDiff = diff
            bestColor = candidateColor
        End If
    Next i
    GetClosestStandardColor = bestColor
End Function

'------------------------------------------------------------
Function DarkenColorRGB(baseRGB As Long, diff As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = baseRGB And &HFF
    g = (baseRGB \ &H100) And &HFF
    b = (baseRGB \ &H10000) And &HFF
    If r - diff < 0 Then r = 0 Else r = r - diff
    If g - diff < 0 Then g = 0 Else g = g - diff
    If b - diff < 0 Then b = 0 Else b = b - diff
    DarkenColorRGB = RGB(r, g, b)
End Function

'------------------------------------------------------------
Function HexToRGB(hexStr As String) As Long
    On Error GoTo ErrorHandler
    hexStr = Trim(hexStr)
    hexStr = Replace(hexStr, "#", "")
    If Len(hexStr) < 6 Then
        HexToRGB = RGB(0, 0, 0)
        Exit Function
    End If
    Dim r As Long, g As Long, b As Long
    r = CLng("&H" & Left(hexStr, 2))
    g = CLng("&H" & Mid(hexStr, 3, 2))
    b = CLng("&H" & Right(hexStr, 2))
    HexToRGB = RGB(r, g, b)
    Exit Function
ErrorHandler:
    HexToRGB = RGB(0, 0, 0)
End Function

'------------------------------------------------------------
Sub DemoRun()
    Dim excludedLayerNames As Variant, backLayerNames As Variant
    excludedLayerNames = Array("excludedLayerName")
    backLayerNames = Array("moveToBackLayerName")
    ' the flag True at the end is whether we should move any Drawio colours towards
    ' our preferred colour scheme
    GenerateDiagramInCurrentSlide "C:\Documents\drawio.xml", _
        excludedLayerNames, backLayerNames, 20, True
End Sub
