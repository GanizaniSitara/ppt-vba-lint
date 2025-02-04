Sub ColorAuditor()
    Dim pptSlide As slide
    Dim pptShape As shape
    Dim fillColor As Long
    Dim lineColor As Long
    Dim hexColor As String
    Dim rgbDesc As String
    Dim colorName As String
    Dim isApproved As Boolean
    Dim approvedColors As Collection
    Dim approvedFonts As Collection
    Dim pptTable As Table
    Dim pptRow As Integer
    Dim pptCol As Integer
    Dim logFile As Object
    Dim timeStamp As String
    Dim logFileName As String

    ' Create a timestamped log file
    timeStamp = Format(Now, "yyyymmdd_hhmmss")
    logFileName = "C:\\temp\\ColourAuditor_" & timeStamp & ".log"
    Set logFile = CreateObject("Scripting.FileSystemObject").CreateTextFile(logFileName, True)

    ' Get the approved colors and fonts
    Set approvedColors = GetApprovedColors()
    Set approvedFonts = GetApprovedFonts()

    ' Loop through all slides sequentially
    Dim slideIndex As Integer
    slideIndex = 1
    For Each pptSlide In ActivePresentation.Slides
        ' Print the title and index of the slide before processing its shapes
        logFile.WriteLine "Processing Slide " & slideIndex & ": " & GetSlideTitle(pptSlide)

        ' Loop through all shapes in the slide
        For Each pptShape In pptSlide.Shapes
            ProcessShape pptShape, approvedColors, approvedFonts, logFile, slideIndex
        Next pptShape

        ' Ensure slide processing completes before moving to the next slide
        logFile.WriteLine "Finished processing Slide " & slideIndex & ": " & GetSlideTitle(pptSlide)
        slideIndex = slideIndex + 1
    Next pptSlide

    logFile.WriteLine "Audit completed."
    logFile.Close

    Debug.Print "Audit completed. Output written to: " & logFileName

    ' Open the log file in Notepad
    OpenLogFile logFileName
End Sub

Sub ProcessShape(pptShape As shape, approvedColors As Collection, approvedFonts As Collection, logFile As Object, slideIndex As Integer)
    Dim fillColor As Long
    Dim lineColor As Long
    Dim hexColor As String
    Dim rgbDesc As String
    Dim colorName As String
    Dim isApproved As Boolean
    Dim pptTable As Table
    Dim pptRow As Integer
    Dim pptCol As Integer

    ' Handle grouped shapes
    If pptShape.Type = msoGroup Then
        Dim groupedShape As shape
        For Each groupedShape In pptShape.GroupItems
            ProcessShape groupedShape, approvedColors, approvedFonts, logFile, slideIndex
        Next groupedShape
    ElseIf pptShape.Type = msoTable Then
        ' Handle tables separately
        Set pptTable = pptShape.Table
        For pptRow = 1 To pptTable.Rows.Count
            For pptCol = 1 To pptTable.Columns.Count
                With pptTable.Cell(pptRow, pptCol).Shape
                    ' Check fill color for each cell
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
        ' Check for fill color if the shape has a fill
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

        ' Check for line color if the shape has a line
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

        ' Check for font compliance in text frames
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
    colors.Add "#FFFF98|Light yellow"
    colors.Add "#C3F5BA|Lime"
    colors.Add "#001276|Bright blue"
    colors.Add "#AFFDFD|Bright mint"
    colors.Add "#5C1E5B|Bright purple"
    colors.Add "#000000|Black"
    colors.Add "#E8E8C9|Stone"
    colors.Add "#FFFF00|Bright yellow"
    colors.Add "#007481|Light teal"
    colors.Add "#00385D|Dark blue"
    colors.Add "#0076B6|Light blue"
    colors.Add "#4C3D6C|Dark purple"
    colors.Add "#E1C0E2|Light purple"
    colors.Add "#D9D9D9|Light grey"
    colors.Add "#FFE05A|Light orange"
    colors.Add "#006666|Teal"
    colors.Add "#CDF5E8|Mint"
    colors.Add "#006DE3|Active blue"
    colors.Add "#FFC9C9|Light claret"
    colors.Add "#7A0FF9|Electric violet"
    colors.Add "#515151|Dark grey"
    colors.Add "#FFCB05|Orange"
    colors.Add "#004750|Dark Teal"
    colors.Add "#3F7F37|Green"
    colors.Add "#0000FF|Electric blue"
    colors.Add "#C7273A|Bright claret"
    colors.Add "#752157|Dark claret"
    colors.Add "#FFFFFF|White"
    colors.Add "#00AEEF|Cyan"
    Set GetApprovedColors = colors
End Function

Function GetApprovedFonts() As Collection
    Dim fonts As New Collection
    fonts.Add "Barclays Effra"
    fonts.Add "Barclays Effra Light"
    fonts.Add "Barclays Effra Medium"
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
    Dim minX As Double
    Dim minY As Double
    Dim titleShape As shape
    Dim slideTitle As String

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
