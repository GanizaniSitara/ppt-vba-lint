Option Explicit
Dim configSettings As Object

Sub ColorAuditor()
    Dim pptSlide As slide
    Dim pptShape As shape
    Dim approvedColors As Collection
    Dim approvedFonts As Collection
    Dim logFile As Object
    Dim timeStamp As String
    Dim logFileName As String
    Dim slideIndex As Integer

    ' Load configuration from C:\temp\config.ini
    Set configSettings = LoadConfig("C:\temp\config.ini")
    If configSettings Is Nothing Then
        MsgBox "Configuration file not found. Exiting.", vbCritical
        Exit Sub
    End If
    If Not configSettings.Exists("fonts") Or Not configSettings.Exists("colours") Then
        MsgBox "Required [fonts] or [colours] section missing in config.ini. Exiting.", vbCritical
        Exit Sub
    End If

    timeStamp = Format(Now, "yyyymmdd_hhmmss")
    logFileName = "C:\temp\ColourAuditor_" & timeStamp & ".log"
    Set logFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(logFileName, True)

    Set approvedColors = GetApprovedColors()
    Set approvedFonts = GetApprovedFonts()

    slideIndex = 1
    For Each pptSlide In ActivePresentation.Slides
        logFile.WriteLine "Processing Slide " & slideIndex & ": " & GetSlideTitle(pptSlide)
        For Each pptShape In pptSlide.Shapes
            ProcessShape pptShape, approvedColors, approvedFonts, logFile, slideIndex
        Next pptShape
        logFile.WriteLine "Finished processing Slide " & slideIndex & ": " & GetSlideTitle(pptSlide)
        slideIndex = slideIndex + 1
    Next pptSlide

    logFile.WriteLine "Audit completed."
    logFile.Close

    Debug.Print "Audit completed. Output written to: " & logFileName
    OpenLogFile logFileName
End Sub

Sub ProcessShape(pptShape As shape, approvedColors As Collection, approvedFonts As Collection, logFile As Object, slideIndex As Integer)
    Dim fillColor As Long, lineColor As Long
    Dim hexColor As String, rgbDesc As String, colorName As String
    Dim isApproved As Boolean
    Dim pptTable As Table
    Dim pptRow As Integer, pptCol As Integer

    If pptShape.Type = msoGroup Then
        Dim groupedShape As shape
        For Each groupedShape In pptShape.GroupItems
            ProcessShape groupedShape, approvedColors, approvedFonts, logFile, slideIndex
        Next groupedShape
    ElseIf pptShape.Type = msoTable Then
        Set pptTable = pptShape.Table
        For pptRow = 1 To pptTable.Rows.Count
            For pptCol = 1 To pptTable.Columns.Count
                With pptTable.Cell(pptRow, pptCol).Shape
                    If .Fill.ForeColor.Type = msoColorTypeRGB Then
                        fillColor = .Fill.ForeColor.RGB
                        hexColor = RGBToHex(fillColor)
                        rgbDesc = RGBToDescription(fillColor)
                        colorName = GetColorName(hexColor, approvedColors)
                        isApproved = (colorName <> "Unknown")
                        If Not isApproved Then
                            logFile.WriteLine "WARN: Slide " & slideIndex & " | Table Cell(" & pptRow & "," & pptCol & ") | " & hexColor & " | " & colorName & " | " & rgbDesc
                        End If
                    End If
                End With
            Next pptCol
        Next pptRow
    Else
        If pptShape.Fill.ForeColor.Type = msoColorTypeRGB Then
            fillColor = pptShape.Fill.ForeColor.RGB
            hexColor = RGBToHex(fillColor)
            rgbDesc = RGBToDescription(fillColor)
            colorName = GetColorName(hexColor, approvedColors)
            isApproved = (colorName <> "Unknown")
            If Not isApproved Then
                logFile.WriteLine "WARN: Slide " & slideIndex & " | " & pptShape.Name & " | " & hexColor & " | " & colorName & " | " & rgbDesc
            End If
        End If

        If pptShape.Line.ForeColor.Type = msoColorTypeRGB Then
            lineColor = pptShape.Line.ForeColor.RGB
            hexColor = RGBToHex(lineColor)
            rgbDesc = RGBToDescription(lineColor)
            colorName = GetColorName(hexColor, approvedColors)
            isApproved = (colorName <> "Unknown")
            If Not isApproved Then
                logFile.WriteLine "WARN: Slide " & slideIndex & " | " & pptShape.Name & " (Line) | " & hexColor & " | " & colorName & " | " & rgbDesc
            End If
        End If

        If pptShape.HasTextFrame Then
            If pptShape.TextFrame.HasText Then
                Dim fontName As String
                fontName = pptShape.TextFrame.TextRange.Font.Name
                If Not IsFontApproved(fontName, approvedFonts) Then
                    logFile.WriteLine "WARN: Slide " & slideIndex & " | " & pptShape.Name & " | Non-compliant font: " & fontName
                End If
            End If
        End If
    End If
End Sub

Function GetApprovedColors() As Collection
    Dim colors As New Collection
    Dim secColours As Object
    Set secColours = configSettings("colours")
    Dim key As Variant
    For Each key In secColours.Keys
        colors.Add secColours(key)
    Next key
    Set GetApprovedColors = colors
End Function

Function GetApprovedFonts() As Collection
    Dim fonts As New Collection
    Dim secFonts As Object
    Set secFonts = configSettings("fonts")
    If secFonts.Exists("normal") Then fonts.Add secFonts("normal")
    If secFonts.Exists("light") Then fonts.Add secFonts("light")
    If secFonts.Exists("medium") Then fonts.Add secFonts("medium")
    Set GetApprovedFonts = fonts
End Function

Function IsFontApproved(fontName As String, approvedFonts As Collection) As Boolean
    Dim approvedFont As Variant
    IsFontApproved = False
    For Each approvedFont In approvedFonts
        If fontName = approvedFont Then
            IsFontApproved = True
            Exit Function
        End If
    Next approvedFont
End Function

Function GetSlideTitle(slide As slide) As String
    Dim minX As Double, minY As Double
    Dim titleShape As shape, slideTitle As String
    minX = slide.Master.Width
    minY = slide.Master.Height
    slideTitle = "Untitled"
    For Each titleShape In slide.Shapes
        If titleShape.HasTextFrame Then
            If titleShape.TextFrame.HasText Then
                If titleShape.Left <= minX And titleShape.Top <= minY Then
                    minX = titleShape.Left
                    minY = titleShape.Top
                    slideTitle = titleShape.TextFrame.TextRange.Text
                End If
            End If
        End If
    Next titleShape
    GetSlideTitle = slideTitle
End Function

Function OpenLogFile(logFileName As String)
    Shell "notepad.exe " & logFileName, vbNormalFocus
End Function

Function RGBToHex(rgbVal As Long) As String
    Dim r As Long, g As Long, b As Long
    r = rgbVal Mod 256
    g = (rgbVal \ 256) Mod 256
    b = (rgbVal \ 65536) Mod 256
    RGBToHex = "#" & Right("0" & Hex(r), 2) & Right("0" & Hex(g), 2) & Right("0" & Hex(b), 2)
End Function

Function RGBToDescription(rgbVal As Long) As String
    RGBToDescription = "R: " & (rgbVal Mod 256) & " G: " & ((rgbVal \ 256) Mod 256) & " B: " & ((rgbVal \ 65536) Mod 256)
End Function

Function GetColorName(hexColor As String, approvedColors As Collection) As String
    Dim colEntry As Variant, parts As Variant
    For Each colEntry In approvedColors
        parts = Split(colEntry, "|")
        If UCase(parts(0)) = UCase(hexColor) Then
            GetColorName = parts(1)
            Exit Function
        End If
    Next colEntry
    GetColorName = "Unknown"
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
        ' Ignore blank lines and lines starting with a semicolon (;)
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
                    value = Trim(Mid(line, eqPos + 1))
                    config(currentSection)(key) = value
                End If
            End If
        End If
    Loop
    ts.Close
    Set LoadConfig = config
End Function
