Sub GenerateDiagramInCurrentSlide(xmlPath As String)
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

    ' Process vertices and edges separately.
    Dim vertexNodes As MSXML2.IXMLDOMNodeList
    Set vertexNodes = xmlDoc.SelectNodes("//mxCell[@vertex='1']")
    
    Dim edgeNodes As MSXML2.IXMLDOMNodeList
    Set edgeNodes = xmlDoc.SelectNodes("//mxCell[@edge='1']")
    
    Dim shp As Shape
    Dim xPos As Single, yPos As Single, widthVal As Single, heightVal As Single
    Dim styleStr As String, fillColor As String, strokeColor As String, pos As Long
    Dim htmlDoc As New MSHTML.HTMLDocument, labelText As String
    Dim geoNode As MSXML2.IXMLDOMNode, node As MSXML2.IXMLDOMNode
    Dim shapeType As MsoAutoShapeType
    
    ' --- Process vertex nodes ---
    For Each node In vertexNodes
        Set geoNode = node.SelectSingleNode("mxGeometry")
        If geoNode Is Nothing Then GoTo NextVertex
        
        ' Get geometry attributes safely.
        If Not geoNode.Attributes.getNamedItem("x") Is Nothing Then
            xPos = Val(geoNode.Attributes.getNamedItem("x").Text)
        Else
            xPos = 0
        End If
        If Not geoNode.Attributes.getNamedItem("y") Is Nothing Then
            yPos = Val(geoNode.Attributes.getNamedItem("y").Text)
        Else
            yPos = 0
        End If
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
        
        ' Get label text from "value" attribute.
        labelText = ""
        If Not node.Attributes.getNamedItem("value") Is Nothing Then
            htmlDoc.body.innerHTML = node.Attributes.getNamedItem("value").Text
            labelText = Trim(htmlDoc.body.innerText)
        End If
        
        ' Get style string.
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
        
        ' For rounded shapes, if no text from value attribute, try "label" attribute.
        If labelText = "" And shapeType = msoShapeRoundedRectangle Then
            If Not node.Attributes.getNamedItem("label") Is Nothing Then
                htmlDoc.body.innerHTML = node.Attributes.getNamedItem("label").Text
                labelText = Trim(htmlDoc.body.innerText)
            End If
        End If
        
        Set shp = pptSlide.Shapes.AddShape(shapeType, xPos, yPos, widthVal, heightVal)
        shp.TextFrame2.TextRange.Text = labelText
        
        ' Uncheck text wrapping.
        shp.TextFrame2.WordWrap = msoFalse
        
        ' Set text formatting based on shape type:
        If shapeType = msoShapeOval Then
            shp.TextFrame2.TextRange.Font.Size = 4
        Else
            shp.TextFrame2.TextRange.Font.Size = 6
        End If
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
        ' Send rectangle or rounded rectangle shapes to back.
        If shapeType = msoShapeRectangle Or shapeType = msoShapeRoundedRectangle Then
            shp.ZOrder msoSendToBack
        End If
        
        ' Set fill color if available.
        pos = InStr(1, styleStr, "fillColor=")
        If pos > 0 Then
            fillColor = Mid(styleStr, pos + 10, 7)
            Dim fillRGB As Long
            fillRGB = HexToRGB(fillColor)
            ' For rectangle shapes, if fill is black, make fill invisible.
            If (shapeType = msoShapeRectangle Or shapeType = msoShapeRoundedRectangle) And fillRGB = RGB(0, 0, 0) Then
                shp.Fill.Visible = msoFalse
            Else
                shp.Fill.ForeColor.RGB = fillRGB
            End If
        End If
        
        ' Set stroke (outline) color if available.
        pos = InStr(1, styleStr, "strokeColor=")
        If pos > 0 Then
            strokeColor = Mid(styleStr, pos + 12, 7)
            shp.Line.ForeColor.RGB = HexToRGB(strokeColor)
        End If
        
        ' Set outline weight to 0.5pt.
        shp.Line.Weight = 0.5
        
NextVertex:
    Next node

    ' --- Process edge (connector) nodes ---
    Dim sourcePtNode As MSXML2.IXMLDOMNode, targetPtNode As MSXML2.IXMLDOMNode
    Dim sourceX As Single, sourceY As Single, targetX As Single, targetY As Single
    Dim edgeStyle As String
    For Each node In edgeNodes
        Set geoNode = node.SelectSingleNode("mxGeometry")
        If geoNode Is Nothing Then GoTo NextEdge
        
        Set sourcePtNode = geoNode.SelectSingleNode("mxPoint[@as='sourcePoint']")
        Set targetPtNode = geoNode.SelectSingleNode("mxPoint[@as='targetPoint']")
        If sourcePtNode Is Nothing Or targetPtNode Is Nothing Then GoTo NextEdge
        
        sourceX = 0: sourceY = 0: targetX = 0: targetY = 0
        If Not sourcePtNode.Attributes.getNamedItem("x") Is Nothing Then
            sourceX = Val(sourcePtNode.Attributes.getNamedItem("x").Text)
        End If
        If Not sourcePtNode.Attributes.getNamedItem("y") Is Nothing Then
            sourceY = Val(sourcePtNode.Attributes.getNamedItem("y").Text)
        End If
        If Not targetPtNode.Attributes.getNamedItem("x") Is Nothing Then
            targetX = Val(targetPtNode.Attributes.getNamedItem("x").Text)
        End If
        If Not targetPtNode.Attributes.getNamedItem("y") Is Nothing Then
            targetY = Val(targetPtNode.Attributes.getNamedItem("y").Text)
        End If
        
        Dim conn As Shape
        Set conn = pptSlide.Shapes.AddConnector(msoConnectorStraight, sourceX, sourceY, targetX, targetY)
        
        ' Set connector style if defined.
        If Not node.Attributes.getNamedItem("style") Is Nothing Then
            edgeStyle = node.Attributes.getNamedItem("style").Text
            pos = InStr(1, edgeStyle, "strokeColor=")
            If pos > 0 Then
                strokeColor = Mid(edgeStyle, pos + 12, 7)
                conn.Line.ForeColor.RGB = HexToRGB(strokeColor)
            End If
        End If
        
        conn.Line.Weight = 0.5
NextEdge:
    Next node

    ' --- Determine bounding box of all shapes ---
    Dim shpItem As Shape
    Dim minX As Single, minY As Single, maxX As Single, maxY As Single
    minX = 1E+30: minY = 1E+30: maxX = -1E+30: maxY = -1E+30
    For Each shpItem In pptSlide.Shapes
        Dim leftVal As Single, topVal As Single, rightVal As Single, bottomVal As Single
        leftVal = shpItem.Left
        topVal = shpItem.Top
        rightVal = shpItem.Left + shpItem.Width
        bottomVal = shpItem.Top + shpItem.Height
        If leftVal < minX Then minX = leftVal
        If topVal < minY Then minY = topVal
        If rightVal > maxX Then maxX = rightVal
        If bottomVal > maxY Then maxY = bottomVal
    Next shpItem

    Dim diagramWidth As Single, diagramHeight As Single
    diagramWidth = maxX - minX
    diagramHeight = maxY - minY

    Dim slideWidth As Single, slideHeight As Single
    slideWidth = pptPres.PageSetup.SlideWidth
    slideHeight = pptPres.PageSetup.SlideHeight

    ' --- Scale to fit 16:9 slide (with margin) ---
    Dim margin As Single: margin = 20 ' points
    Dim scaleFactorX As Single, scaleFactorY As Single, scaleFactor As Single
    scaleFactorX = (slideWidth - 2 * margin) / diagramWidth
    scaleFactorY = (slideHeight - 2 * margin) / diagramHeight
    scaleFactor = scaleFactorX
    If scaleFactorY < scaleFactor Then scaleFactor = scaleFactorY
    
    Dim scaled As Boolean: scaled = False
    ' Only scale if diagram is larger than available area.
    If scaleFactor < 1 Then
        scaled = True
        For Each shpItem In pptSlide.Shapes
            shpItem.Left = margin + (shpItem.Left - minX) * scaleFactor
            shpItem.Top = margin + (shpItem.Top - minY) * scaleFactor
            shpItem.Width = shpItem.Width * scaleFactor
            shpItem.Height = shpItem.Height * scaleFactor
            ' Ensure text remains at proper size after scaling.
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
        ' Center the diagram if no scaling is needed.
        Dim offsetX As Single, offsetY As Single
        offsetX = (slideWidth - diagramWidth) / 2 - minX
        offsetY = (slideHeight - diagramHeight) / 2 - minY
        For Each shpItem In pptSlide.Shapes
            shpItem.Left = shpItem.Left + offsetX
            shpItem.Top = shpItem.Top + offsetY
        Next shpItem
    End If

    If scaled Then
        Debug.Print "Diagram scaled to fit slide."
    Else
        Debug.Print "Diagram fits slide without scaling; diagram centered."
    End If
End Sub

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

Sub DemoRun()
    Dim ignoreList As Variant    
    ' Drawio hast to be CONTENT ONLY xml as you'd take it out of the "Edit Diagram" menu in drawio.	
    GenerateDiagramInCurrentSlide "C:\Documents\drawio.filw.xml"
End Sub
