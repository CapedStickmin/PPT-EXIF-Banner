Option Explicit

' --- 全局变量与常量 ---
Private exifCollection As Object
' !!! 重要：请根据您的实际情况修改下面这两行路径 !!!
' Logo可以使用网络链接，但为了更快的速度和离线使用，建议下载到本地并使用本地路径。
Const NIKON_LOGO_PATH As String = "https://i1.wp.com/naturebyandreas.se/wp-content/uploads/2015/10/nikon-logo.jpg"
Const CSV_FILE_PATH As String = "E:\成片\exif_data.csv"

' =========================================================================
' --- 主宏程序 ---
' =========================================================================
Public Sub CreateBannersOnNewSlides_Batch()
    On Error GoTo ErrorHandler

    Dim fileDialog As Object
    Dim imagePath As Variant
    Dim imageName As String
    Dim oSlide As Slide
    Dim oPicture As Shape
    Dim exifData As Variant
    Dim processedCount As Long
    Dim failedFiles As String

    ' --- 步骤 1: 一次性加载全部EXIF数据 ---
    If LoadExifData() = False Then Exit Sub

    ' --- 步骤 2: 打开文件选择对话框，并允许多选 ---
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "请选择一个或多个照片文件 (可按住 Ctrl 或 Shift)"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Images", "*.JPG;*.JPEG"
        
        If .Show <> -1 Then
            MsgBox "操作已取消。", vbInformation
            Exit Sub
        End If
    End With

    ' --- 步骤 3: 循环处理每一个选中的文件 ---
    processedCount = 0
    failedFiles = ""
    
    For Each imagePath In fileDialog.SelectedItems
        imageName = Mid(CStr(imagePath), InStrRev(CStr(imagePath), "\") + 1)

        If exifCollection.Exists(imageName) Then
            exifData = exifCollection(imageName)

            ' 为每张图片在末尾创建一张新的空白幻灯片
            Set oSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)
            Set oPicture = oSlide.Shapes.AddPicture(fileName:=CStr(imagePath), LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=0, Top:=0)
            
            PositionAndCreateLayout oSlide, oPicture, exifData

            processedCount = processedCount + 1
        Else
            ' 记录找不到数据的文件名
            failedFiles = failedFiles & vbCrLf & " - " & imageName
        End If
    Next imagePath

    ' --- 步骤 4: 显示最终的处理报告 ---
    Dim finalMsg As String
    finalMsg = "批量处理完成！" & vbCrLf & vbCrLf & "成功为 " & processedCount & " 张图片创建了新页面。"
    
    If failedFiles <> "" Then
        finalMsg = finalMsg & vbCrLf & vbCrLf & "下列文件的数据未在CSV中找到，已跳过：" & failedFiles
    End If
    
    MsgBox finalMsg, vbInformation, "处理报告"

    ' --- 清理对象 ---
    Set fileDialog = Nothing
    Set oPicture = Nothing
    Set oSlide = Nothing
    
    Exit Sub

ErrorHandler:
    MsgBox "宏在运行过程中发生严重错误: " & Err.Description, vbCritical, "Macro Error"
End Sub


' =========================================================================
' --- 子程序 ---
' =========================================================================
Private Sub PositionAndCreateLayout(ByRef sld As Slide, ByRef pic As Shape, ByVal data As Variant)
    ' --- 【横幅宽度调节】---
    Const FIXED_BANNER_WIDTH As Single = 600 ' <--- 在这里修改为你想要的宽度

    Dim sldWidth As Single, sldHeight As Single
    Dim safeWidth As Single, safeHeight As Single
    Dim group As Shape
    Dim oBanner As Shape, oLogo As Shape, oTextBoxInfo As Shape, oTextBoxParams As Shape
    Dim infoText As String, paramsText As String
    Dim bannerHeight As Single, totalVisualHeight As Single, bannerTop As Single, bannerLeft As Single

    sldWidth = ActivePresentation.PageSetup.slideWidth
    sldHeight = ActivePresentation.PageSetup.slideHeight
    
    ' --- 【安全区调节】---
    ' 0.87 表示图片和横幅的总高度最多占幻灯片高度的87%，但是注意留白包括横幅范围。
    safeHeight = sldHeight * 0.87
    safeWidth = sldWidth * 0.95

    With pic
        .LockAspectRatio = msoTrue
        If (.Width / .Height) >= (safeWidth / safeHeight) Then
            .Width = safeWidth
        Else
            .Height = safeHeight
        End If
        
        With .Shadow
            .Type = msoShadow26
            .Visible = msoTrue
            .Blur = 12
            .Transparency = 0.5
            .OffsetX = 0
            .OffsetY = 6
        End With
    End With

    infoText = CStr(data(0)) & vbCrLf & CStr(data(1))
    paramsText = CStr(data(2)) & "    F" & CStr(data(3)) & "    " & CStr(data(4)) & "s    " & CStr(data(5))
    
    bannerHeight = 40
    totalVisualHeight = pic.Height + 10 + bannerHeight
    
    pic.Left = (sldWidth - pic.Width) / 2
    bannerLeft = (sldWidth - FIXED_BANNER_WIDTH) / 2
    
    pic.Top = (sldHeight - totalVisualHeight) / 2
    
    bannerTop = pic.Top + pic.Height + 10
    
    Set oBanner = sld.Shapes.AddShape(msoShapeRectangle, bannerLeft, bannerTop, FIXED_BANNER_WIDTH, bannerHeight)
    With oBanner
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0.4
        .line.Visible = msoFalse
    End With
    
    Set oLogo = sld.Shapes.AddPicture(NIKON_LOGO_PATH, msoFalse, msoTrue, 0, 0)
    With oLogo
        .LockAspectRatio = msoTrue
        .Height = 30
        .Left = oBanner.Left + 15
        .Top = oBanner.Top + (oBanner.Height - .Height) / 2
    End With
    
    Set oTextBoxInfo = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, oLogo.Left + oLogo.Width + 10, oBanner.Top, oBanner.Width * 0.5, bannerHeight)
    With oTextBoxInfo.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = infoText
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 11
        .TextRange.Font.Fill.ForeColor.RGB = RGB(220, 220, 220)
        .TextRange.Font.Bold = msoTrue
    End With
    
    Set oTextBoxParams = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, oBanner.Left + oBanner.Width * 0.5, oBanner.Top, oBanner.Width * 0.5 - 15, bannerHeight)
    With oTextBoxParams.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = paramsText
        .TextRange.ParagraphFormat.Alignment = msoAlignRight
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 14
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.Font.Bold = msoTrue
    End With
    
    Set group = sld.Shapes.Range(Array(pic.Name, oBanner.Name, oLogo.Name, oTextBoxInfo.Name, oTextBoxParams.Name)).group

    Set oBanner = Nothing
    Set oLogo = Nothing
    Set oTextBoxInfo = Nothing
    Set oTextBoxParams = Nothing
    Set group = Nothing
End Sub


' =========================================================================
' --- 数据加载函数 ---
' =========================================================================
Private Function LoadExifData() As Boolean
    On Error GoTo LoadError
    
    If Not exifCollection Is Nothing Then LoadExifData = True: Exit Function
    
    If Dir(CSV_FILE_PATH) = "" Then
        MsgBox "错误：找不到数据文件！" & vbCrLf & CSV_FILE_PATH, vbCritical, "数据文件丢失"
        LoadExifData = False: Exit Function
    End If
    
    Dim adoStream As Object, line As String, columns() As String, key As String, value As Variant
    Set adoStream = CreateObject("ADODB.Stream")
    With adoStream: .Type = 2: .Charset = "UTF-8": .Open: .LoadFromFile CSV_FILE_PATH: End With
    
    Set exifCollection = CreateObject("Scripting.Dictionary")
    exifCollection.CompareMode = vbTextCompare
    
    If Not adoStream.EOS Then line = adoStream.ReadText(-2) ' Skip header
    
    Do Until adoStream.EOS
        line = adoStream.ReadText(-2)
        columns = Split(line, ",")
        If UBound(columns) >= 7 Then
            key = Trim(Replace(columns(1), """", ""))
            value = Array(CStr(columns(2)), CStr(columns(3)), CStr(columns(4)), CStr(columns(5)), CStr(columns(6)), "ISO " & CStr(columns(7)))
            If Not exifCollection.Exists(key) Then exifCollection.Add key, value
        End If
    Loop
    
    adoStream.Close: Set adoStream = Nothing
    LoadExifData = True
    Exit Function
    
LoadError:
    MsgBox "加载EXIF数据时发生错误: " & Err.Description, vbCritical, "Load Data Error"
    LoadExifData = False
    If Not adoStream Is Nothing Then
        If adoStream.State = 1 Then adoStream.Close
        Set adoStream = Nothing
    End If
End Function```