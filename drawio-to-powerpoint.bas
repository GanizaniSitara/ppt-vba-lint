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
                    labelFromParent = Trim(node.ParentNode.Attributes.getNamedItem("label").Text)
                End If
            End If
        End If
        If Not node.Attributes.getNamedItem("value") Is Nothing Then
            htmlDoc.body.innerHTML = node.Attributes.getNamedItem("value").Text
            labelFromValue = Trim(htmlDoc.body.innerText)
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
            labelText = labelText & vbCrLf & Trim(node.Attributes.getNamedItem("description").Text)
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

        pos = InStr(1, styleStr, "fillColor=")
        If pos > 0 Then
            fillColorStr = Mid(styleStr, pos + 10, 7)
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

        If shp.Fill.Visible = msoFalse Then shp.ZOrder msoBringToFront

        pos = InStr(1, styleStr, "strokeColor=")
        If pos > 0 Then
            strokeColorStr = Mid(styleStr, pos + 12, 7)
            Dim baseLineRGB As Long
            baseLineRGB = HexToRGB(strokeColorStr)
            shp.Line.ForeColor.RGB = DarkenColorRGB(baseLineRGB, lineShadeDiff)
        End If
        shp.Line.Weight = 0.5

        If Not node.Attributes.getNamedItem("id") Is Nothing Then
            currentId = node.Attributes.getNamedItem("id").Text
            shapeMap.Add currentId, shp
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

        Set geoNode = node.SelectSingleNode("mxGeometry")
        If geoNode Is Nothing Then GoTo NextEdge

        Set sourcePtNode = geoNode.SelectSingleNode("mxPoint[@as='sourcePoint']")
        Set targetPtNode = geoNode.SelectSingleNode("mxPoint[@as='targetPoint']")
        If Not sourcePtNode Is Nothing Then
            If Not sourcePtNode.Attributes.getNamedItem("x") Is Nothing Then
                sourceX = Val(sourcePtNode.Attributes.getNamedItem("x").Text)
            Else
                sourceX = 0
            End If
            If Not sourcePtNode.Attributes.getNamedItem("y") Is Nothing Then
                sourceY = Val(sourcePtNode.Attributes.getNamedItem("y").Text)
            Else
                sourceY = 0
            End If
        Else
            sourceX = 0: sourceY = 0
        End If
        If Not targetPtNode Is Nothing Then
            If Not targetPtNode.Attributes.getNamedItem("x") Is Nothing Then
                targetX = Val(targetPtNode.Attributes.getNamedItem("x").Text)
            Else
                targetX = 0
            End If
            If Not targetPtNode.Attributes.getNamedItem("y") Is Nothing Then
                targetY = Val(targetPtNode.Attributes.getNamedItem("y").Text)
            Else
                targetY = 0
            End If
        Else
            targetX = 0: targetY = 0
        End If

        Dim conn As Shape
        Set conn = pptSlide.Shapes.AddConnector(msoConnectorStraight, sourceX, sourceY, targetX, targetY)
        conn.Line.ForeColor.RGB = RGB(0, 0, 0)
        conn.Line.Weight = 0.5
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
NextEdge:
    Next node

    '--- Determine Bounding Box for Scaling ---
    Dim shpItem As Shape
    Dim bbMinX As Single, bbMinY As Single, bbMaxX As Single, bbMaxY As Single
    bbMinX = 1E+30: bbMinY = 1E+30: bbMaxX = -1E+30: bbMaxY = -1E+30
    For Each shpItem In pptSlide.Shapes
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
        For Each shpItem In pptSlide.Shapes
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
        For Each shpItem In pptSlide.Shapes
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
Function FormatLabelText(ByVal node As MSXML2.IXMLDOMNode, ByVal lbl As String) As String
    Dim parts() As String, i As Long, result As String, placeholder As String, attrValue As String
    parts = Split(lbl, "%")
    result = ""
    For i = 1 To UBound(parts) Step 2
        placeholder = Trim(parts(i))
        If Not node.Attributes.getNamedItem(placeholder) Is Nothing Then
            attrValue = node.Attributes.getNamedItem(placeholder).Text
        Else
            attrValue = placeholder
        End If
        result = result & attrValue & vbCrLf
    Next i
    If Len(result) >= 2 Then
        result = Left(result, Len(result) - 2)
    End If
    FormatLabelText = result
End Function

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
