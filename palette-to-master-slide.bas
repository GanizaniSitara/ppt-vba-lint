Function HexToLong(ByVal sHex As String) As Long
    sHex = Replace(sHex, "#", "")
    HexToLong = RGB(CLng("&H" & Mid(sHex, 1, 2)), _
                    CLng("&H" & Mid(sHex, 3, 2)), _
                    CLng("&H" & Mid(sHex, 5, 2)))
End Function

Function LoadConfig(configPath As String) As Object
    Dim fso As Object, ts As Object, line As String
    Dim currentSection As String
    Dim config As Object, sectionDict As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(configPath) Then
        Set LoadConfig = Nothing
        Exit Function
    End If
    
    Set config = CreateObject("Scripting.Dictionary")
    Set ts = fso.OpenTextFile(configPath, 1)
    currentSection = ""
    
    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        
        If line <> "" And Left(line, 1) <> ";" Then
            If Left(line, 1) = "[" And Right(line, 1) = "]" Then
                currentSection = LCase(Mid(line, 2, Len(line) - 2))
                If Not config.Exists(currentSection) Then
                    Set sectionDict = CreateObject("Scripting.Dictionary")
                    config.Add currentSection, sectionDict
                End If
            ElseIf currentSection <> "" Then
                Dim eqPos As Long, key As String, value As String
                eqPos = InStr(line, "=")
                If eqPos > 0 Then
                    key = LCase(Trim(Left(line, eqPos - 1)))
                    value = Trim(Split(Mid(line, eqPos + 1), "|")(0)) ' Extract only hex color
                    config(currentSection)(key) = value
                End If
            End If
        End If
    Loop
    ts.Close
    
    Set LoadConfig = config
End Function

Sub AddSwatchesToSlideMaster()
    Dim configPath As String
    configPath = "C:\temp\config.ini"
    
    Dim config As Object
    Set config = LoadConfig(configPath)
    If config Is Nothing Then Exit Sub
    
    Dim hexColors As Object
    Set hexColors = config("colours")
    
    Dim hexArray() As String
    Dim i As Long
    ReDim hexArray(0 To hexColors.Count - 1)
    
    Dim key As Variant
    i = 0
    For Each key In hexColors
        hexArray(i) = hexColors(key)
        i = i + 1
    Next key
    
    Dim swatchW As Single, swatchH As Single, gap As Single
    swatchW = 0.25 * 28.35: swatchH = swatchW: gap = 1
    
    Dim totalSwatches As Long
    totalSwatches = UBound(hexArray) - LBound(hexArray) + 1
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
    
    Dim shp As Shape
    Dim shapeNames() As String
    ReDim shapeNames(0 To totalSwatches - 1)
    
    For i = 0 To totalSwatches - 1
        Set shp = sldMaster.Shapes.AddShape(msoShapeRectangle, _
            startX + i * (swatchW + gap), startY, swatchW, swatchH)
        shp.Fill.ForeColor.RGB = HexToLong(CStr(hexArray(i)))
        shp.Line.Visible = msoFalse
        shapeNames(i) = shp.Name
    Next i
    
    sldMaster.Shapes.Range(shapeNames).Group
End Sub

