Sub ExportSlideRegion()
    Dim sld As Slide
    Dim exportPath As String, tempImgPath As String, newImgPath As String
    Dim tmpShape As Shape, picShape As Shape
    Dim slideW As Single, slideH As Single
    Dim desiredX As Single, desiredY As Single, desiredW As Single, desiredH As Single

    exportPath = ActivePresentation.Path
    If exportPath = "" Then exportPath = "C:\Temp"

    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight

    ' Define region: top right corner region
    desiredW = 400   ' desired width of exported region
    desiredH = 200   ' desired height of exported region
    desiredX = slideW - desiredW
    desiredY = 0

    For Each sld In ActivePresentation.Slides
        tempImgPath = exportPath & "\temp_slide.png"
        sld.Export tempImgPath, "PNG", slideW, slideH

        ' Insert full slide as picture
        Set tmpShape = sld.Shapes.AddPicture(FileName:=tempImgPath, _
            LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
            Left:=0, Top:=0, Width:=slideW, Height:=slideH)

        With tmpShape.PictureFormat
            .CropLeft = desiredX
            .CropTop = desiredY
            .CropRight = slideW - (desiredX + desiredW)
            .CropBottom = slideH - (desiredY + desiredH)
        End With

        ' Copy and paste as picture to ensure proper export of cropped region
        tmpShape.Copy
        Set picShape = sld.Shapes.PasteSpecial(DataType:=ppPastePNG)(1)

        newImgPath = exportPath & "\Slide" & sld.SlideIndex & "_region.png"
        picShape.Export newImgPath, ppShapeFormatPNG

        picShape.Delete
        tmpShape.Delete
        Kill tempImgPath
    Next sld
End Sub