VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form6 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "PG Draw - Process Image"
   ClientHeight    =   4515
   ClientLeft      =   3090
   ClientTop       =   2385
   ClientWidth     =   5130
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   301
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   342
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4470
      Left            =   0
      ScaleHeight     =   298
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6045
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   1.17485e-38
   End
   Begin VB.Menu menucommand 
      Caption         =   "Command"
      Begin VB.Menu CommandOk 
         Caption         =   "OK"
      End
      Begin VB.Menu CommandCancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu ProcessMenu 
      Caption         =   "Process Image"
      Begin VB.Menu ProcessSmooth 
         Caption         =   "Smooth"
      End
      Begin VB.Menu ProcessSharpen 
         Caption         =   "Sharpen"
      End
      Begin VB.Menu ProcessEmboss 
         Caption         =   "Emboss"
      End
      Begin VB.Menu ProcessDiffuse 
         Caption         =   "Diffuse"
      End
      Begin VB.Menu ProcessPixelize 
         Caption         =   "Pixelize"
      End
      Begin VB.Menu ProcessSolarize 
         Caption         =   "Solarize"
      End
      Begin VB.Menu ProcessGrayscale 
         Caption         =   "GrayScale"
      End
      Begin VB.Menu ProcessBright 
         Caption         =   "Brighten"
      End
      Begin VB.Menu contrast 
         Caption         =   "Contrast"
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu ProcessCustom 
         Caption         =   "Custom Filter"
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim t1, HBmp, hDestDC As Variant
Dim i As Long, j As Long
Dim red As Integer, blue As Integer, green As Integer


Private Sub CommandCancel_Click()
    Unload Me
End Sub

Private Sub CommandOk_Click()
    Form1.Picture2.Picture = Picture1.Image
    Unload Me
End Sub


Private Sub contrast_Click()
Form_con_param.Show
End Sub

Private Sub Form_Load()
    On Error GoTo BadImageType
    Dim pixel As Long
    
    Picture1.Picture = Form1.Picture2.Image
    Form6.Refresh

    x = Picture1.ScaleWidth
    y = Picture1.ScaleHeight
    If x > 800 Or y > 800 Then
        MsgBox "Image too large to process. Please try loading a smaller image."
        Unload Me
    
End If

Form6.Width = Form6.ScaleX(Picture1.Width + 6, vbPixels, vbTwips)
Form6.Height = Form6.ScaleY(Picture1.Height + 43, vbPixels, vbTwips)
Form6.BorderStyle = 1
Form6.Refresh

Form4.Show
Form4.Caption = "Reading pixels"
Form4.Refresh

    For i = 0 To y - 1
        For j = 0 To x - 1
            pixel = GetPixel(Form6.Picture1.hdc, j, i)
            red = pixel Mod 256
            green = ((pixel And &HFF00) / 256&) Mod 256&
            blue = (pixel And &HFF0000) / 65536
            ImagePixels(0, i, j) = red
            ImagePixels(1, i, j) = green
            ImagePixels(2, i, j) = blue
        Next
        Form4.ProgressBar1.value = i * 100 / (y - 1)
        DoEvents
    Next
    Form4.Hide
    Exit Sub
    
BadImageType:
    MsgBox Err.Description
Exit Sub

End Sub


Private Sub ProcessBright_Click()
    



    Dim max As Integer
    Dim add As Integer
    Form4.Show
    Form4.Caption = "Brightening image ..."
    Form4.Refresh

    HBmp = CreateCompatibleBitmap(Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Picture1.hdc)
    SelectObject hDestDC, HBmp
    
    t1 = Timer
    For i = 1 To y - 2
        For j = 1 To x - 2
            
            If ImagePixels(0, i, j) >= ImagePixels(1, i, j) Then
            max = ImagePixels(0, i, j)
            If max < ImagePixels(2, i, j) Then
            max = ImagePixels(2, i, j)
            End If
            Else
            max = ImagePixels(1, i, j)
            If max < ImagePixels(2, i, j) Then
            max = ImagePixels(2, i, j)
            End If
            End If
            
            If max < 235 Then
            
            
            red = ImagePixels(0, i, j) + 20
            green = ImagePixels(1, i, j) + 20
            blue = ImagePixels(2, i, j) + 20
            Else
            add = 255 - max
            red = ImagePixels(0, i, j) + add
            green = ImagePixels(1, i, j) + add
            blue = ImagePixels(2, i, j) + add
            End If
            SetPixelV hDestDC, j, i, RGB(red, green, blue)
        Next
        Form4.ProgressBar1.value = i * 100 / (y - 1)
        DoEvents
    Next
    Form4.Hide
    BitBlt Picture1.hdc, 1, 1, Picture1.ScaleWidth - 2, Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Picture1.Refresh
    Call DeleteDC(hDestDC)
    Call DeleteObject(HBmp)

' UNCOMMENT NEXT LINE TO TIME OPERATION
    MsgBox "Processing completed in " & Format(Timer - t1, "##.000")

End Sub

Private Sub ProcessCustom_Click()
Dim RedSum As Integer, GreenSum As Integer, BlueSum As Integer


Dim fi As Integer, fj As Integer
Dim Offset As Integer
Dim Weight As Single

    Form5.Show 1    ' wait for user to define filter
    If FilterCancel = True Then Exit Sub
    t1 = Timer
    If FilterNorm = 0 Then FilterNorm = 1
    
    If Form5.Option1.value Then
        Offset = 1
    Else
        Offset = 2
    End If
    DoEvents
    
    Form4.Show
    Form4.Caption = "Processing pixels..."
    Form4.Refresh
    
    HBmp = CreateCompatibleBitmap(Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Picture1.hdc)
    SelectObject hDestDC, HBmp
    
    For i = Offset To y - Offset - 1
        For j = Offset To x - Offset - 1
            RedSum = 0: GreenSum = 0: BlueSum = 0
            For fi = -Offset To Offset
                For fj = -Offset To Offset
                    Weight = CustomFilter(fi + 2, fj + 2)
                    RedSum = RedSum + ImagePixels(0, i + fi, j + fj) * Weight
                    GreenSum = GreenSum + ImagePixels(1, i + fi, j + fj) * Weight
                    BlueSum = BlueSum + ImagePixels(2, i + fi, j + fj) * Weight
                Next
            Next
            red = Abs(RedSum / FilterNorm + FilterBias)
            green = Abs(GreenSum / FilterNorm + FilterBias)
            blue = Abs(BlueSum / FilterNorm + FilterBias)
            SetPixelV hDestDC, j, i, RGB(red, green, blue)
        Next
        Form4.ProgressBar1.value = i * 100 / (y - 1)
        DoEvents
    Next
    Form4.Hide
    BitBlt Picture1.hdc, 1, 1, Picture1.ScaleWidth - 2, Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Picture1.Refresh
    Call DeleteDC(hDestDC)
    Call DeleteObject(HBmp)

' UNCOMMENT NEXT LINE TO TIME OPERATION
    MsgBox "Processing completed in " & Format(Timer - t1, "##.000")
End Sub

Private Sub ProcessDiffuse_Click()


Dim Rx As Integer, Ry As Integer
   
    t1 = Timer
    Form4.Show
    Form4.Caption = "Diffusing image ..."
    Form4.Refresh

    HBmp = CreateCompatibleBitmap(Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Picture1.hdc)
    SelectObject hDestDC, HBmp
    For i = 2 To y - 3
        For j = 2 To x - 3
            Rx = Rnd * 4 - 2
            Ry = Rnd * 4 - 2
            red = ImagePixels(0, i + Rx, j + Ry)
            green = ImagePixels(1, i + Rx, j + Ry)
            blue = ImagePixels(2, i + Rx, j + Ry)
            SetPixelV hDestDC, j, i, RGB(red, green, blue)
        Next
        Form4.ProgressBar1.value = i * 100 / (y - 1)
        DoEvents
    Next
    Form4.Hide
    BitBlt Picture1.hdc, 1, 1, Picture1.ScaleWidth - 2, Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Picture1.Refresh
    Call DeleteDC(hDestDC)
    Call DeleteObject(HBmp)

' UNCOMMENT NEXT LINE TO TIME OPERATION
    MsgBox "Processing completed in " & Format(Timer - t1, "##.000")
End Sub

Private Sub ProcessEmboss_Click()
Dim Dx As Integer, Dy As Integer


    Dx = 1
    Dy = 1
    
    Form4.Show
    Form4.Caption = "Embossing image ..."
    Form4.Refresh

    HBmp = CreateCompatibleBitmap(Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Picture1.hdc)
    SelectObject hDestDC, HBmp
    
    t1 = Timer
    For i = 1 To y - 2
        For j = 1 To x - 2
            red = Abs(ImagePixels(0, i, j) - ImagePixels(0, i + Dx, j + Dy) + 128)
            green = Abs(ImagePixels(1, i, j) - ImagePixels(1, i + Dx, j + Dy) + 128)
            blue = Abs(ImagePixels(2, i, j) - ImagePixels(2, i + Dx, j + Dy) + 128)
            SetPixelV hDestDC, j, i, RGB(red, green, blue)
        Next
        Form4.ProgressBar1.value = i * 100 / (y - 1)
        DoEvents
    Next
    Form4.Hide
    BitBlt Picture1.hdc, 1, 1, Picture1.ScaleWidth - 2, Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Picture1.Refresh
    Call DeleteDC(hDestDC)
    Call DeleteObject(HBmp)

' UNCOMMENT NEXT LINE TO TIME OPERATION
    MsgBox "Processing completed in " & Format(Timer - t1, "##.000")
    
End Sub

Private Sub ProcessGrayscale_Click()

Dim graytone As Integer

    
    Form4.Show
    Form4.Caption = "GrayScaling image ..."
    Form4.Refresh

    HBmp = CreateCompatibleBitmap(Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Picture1.hdc)
    SelectObject hDestDC, HBmp
    
    t1 = Timer
    For i = 1 To y - 2
        For j = 1 To x - 2
            graytone = (ImagePixels(0, i, j) + ImagePixels(1, i, j) + ImagePixels(2, i, j)) / 3
            SetPixelV hDestDC, j, i, RGB(graytone, graytone, graytone)
        Next
        Form4.ProgressBar1.value = i * 100 / (y - 1)
        DoEvents
    Next
    Form4.Hide
    BitBlt Picture1.hdc, 1, 1, Picture1.ScaleWidth - 2, Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Picture1.Refresh
    Call DeleteDC(hDestDC)
    Call DeleteObject(HBmp)

' UNCOMMENT NEXT LINE TO TIME OPERATION
    MsgBox "Processing completed in " & Format(Timer - t1, "##.000")

End Sub

Private Sub ProcessPixelize_Click()

Dim Dx As Integer, Dy As Integer

Dim xpixel As Integer, ypixel As Integer
Dim R
    t1 = Timer
    Picture1.FillStyle = vbSolid
    For i = 1 To y / 3
        For j = 1 To x / 3
            ypixel = Rnd * x + 4 - 2
            xpixel = Rnd * y + 4 - 2
            R = Int(Rnd() * 3) + 2
            red = ImagePixels(0, xpixel, ypixel)
            green = ImagePixels(1, xpixel, ypixel)
            blue = ImagePixels(2, xpixel, ypixel)
            Picture1.FillColor = RGB(red, green, blue)
            Picture1.Circle (ypixel, xpixel), R, RGB(red, green, blue)
        Next
        Picture1.Refresh
    Next
    Picture1.FillStyle = vbTransparent
' UNCOMMENT NEXT LINE TO TIME OPERATION
    MsgBox "Processing completed in " & Format(Timer - t1, "##.000")

End Sub

Private Sub ProcessSharpen_Click()

Dim Dx As Integer, Dy As Integer



    Dx = 1: Dy = 1
    t1 = Timer
    Form4.Show
    Form4.Caption = "Sharpening image ..."
    Form4.Refresh
    
    HBmp = CreateCompatibleBitmap(Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Picture1.hdc)
    SelectObject hDestDC, HBmp
    
    For i = 1 To y - 2
        For j = 1 To x - 2
            red = ImagePixels(0, i, j) + 0.5 * (ImagePixels(0, i, j) - ImagePixels(0, i - Dx, j - Dy))
            green = ImagePixels(1, i, j) + 0.5 * (ImagePixels(1, i, j) - ImagePixels(1, i - Dx, j - Dy))
            blue = ImagePixels(2, i, j) + 0.5 * (ImagePixels(2, i, j) - ImagePixels(2, i - Dx, j - Dy))
            If red > 255 Then red = 255
            If red < 0 Then red = 0
            If green > 255 Then green = 255
            If green < 0 Then green = 0
            If blue > 255 Then blue = 255
            If blue < 0 Then blue = 0
            SetPixelV hDestDC, j, i, RGB(red, green, blue)
        Next
        Form4.ProgressBar1.value = i * 100 / (y - 1)
        DoEvents
    Next
    Form4.Hide
    BitBlt Picture1.hdc, 1, 1, Picture1.ScaleWidth - 2, Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Picture1.Refresh
    Call DeleteDC(hDestDC)
    Call DeleteObject(HBmp)
    
' UNCOMMENT NEXT LINE TO TIME OPERATION
    MsgBox "Processing completed in " & Format(Timer - t1, "##.000")
End Sub



Private Sub ProcessSmooth_Click()



    Form4.Show
    Form4.Caption = "Smoothing image ..."
    Form4.Refresh
    t1 = Timer
    HBmp = CreateCompatibleBitmap(Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Picture1.hdc)
    SelectObject hDestDC, HBmp

    For i = 1 To y - 2
        For j = 1 To x - 2
            red = ImagePixels(0, i - 1, j - 1) + ImagePixels(0, i - 1, j) + ImagePixels(0, i - 1, j + 1) + _
            ImagePixels(0, i, j - 1) + ImagePixels(0, i, j) + ImagePixels(0, i, j + 1) + _
            ImagePixels(0, i + 1, j - 1) + ImagePixels(0, i + 1, j) + ImagePixels(0, i + 1, j + 1)
            
            green = ImagePixels(1, i - 1, j - 1) + ImagePixels(1, i - 1, j) + ImagePixels(1, i - 1, j + 1) + _
            ImagePixels(1, i, j - 1) + ImagePixels(1, i, j) + ImagePixels(1, i, j + 1) + _
            ImagePixels(1, i + 1, j - 1) + ImagePixels(1, i + 1, j) + ImagePixels(1, i + 1, j + 1)
            
            blue = ImagePixels(2, i - 1, j - 1) + ImagePixels(2, i - 1, j) + ImagePixels(2, i - 1, j + 1) + _
            ImagePixels(2, i, j - 1) + ImagePixels(2, i, j) + ImagePixels(2, i, j + 1) + _
            ImagePixels(2, i + 1, j - 1) + ImagePixels(2, i + 1, j) + ImagePixels(2, i + 1, j + 1)
            
            SetPixelV hDestDC, j, i, RGB(red / 9, green / 9, blue / 9)
        Next
        Form4.ProgressBar1.value = i * 100 / (y - 1)
        DoEvents
    Next
    Form4.Hide
    BitBlt Picture1.hdc, 1, 1, Picture1.ScaleWidth - 2, Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Picture1.Refresh
    Call DeleteDC(hDestDC)
    Call DeleteObject(HBmp)
    
' UNCOMMENT NEXT LINE TO TIME OPERATION
    MsgBox "Processing completed in " & Format(Timer - t1, "##.000")
End Sub

Private Sub ProcessSolarize_Click()



    Form4.Show
    Form4.Caption = "Solarizing image ..."
    Form4.Refresh
    t1 = Timer
    HBmp = CreateCompatibleBitmap(Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Picture1.hdc)
    SelectObject hDestDC, HBmp

    For i = 1 To y - 2
        For j = 1 To x - 2
            red = ImagePixels(0, i, j)
            green = ImagePixels(1, i, j)
            blue = ImagePixels(2, i, j)
            If ((red < 128) Or (red > 255)) Then red = 255 - red
            If ((green < 128) Or (green > 255)) Then green = 255 - green
            If ((blue < 128) Or (blue > 255)) Then blue = 255 - blue
            SetPixelV hDestDC, j, i, RGB(red, green, blue)
        Next
        Form4.ProgressBar1.value = i * 100 / (y - 1)
        DoEvents
    Next
    Form4.Hide
    BitBlt Picture1.hdc, 1, 1, Picture1.ScaleWidth - 2, Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Picture1.Refresh
    Call DeleteDC(hDestDC)
    Call DeleteObject(HBmp)
    
' UNCOMMENT NEXT LINE TO TIME OPERATION
    MsgBox "Processing completed in " & Format(Timer - t1, "##.000")

End Sub

Sub contrast_image()

Dim max As Integer
    Dim add As Integer
    Dim graytone, value As Integer
    
    Form4.Show
    Form4.Caption = "Contrasting image ..."
    Form4.Refresh

    HBmp = CreateCompatibleBitmap(Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight)
    hDestDC = CreateCompatibleDC(Picture1.hdc)
    SelectObject hDestDC, HBmp
    
    t1 = Timer
    
    For i = 1 To y - 2
        For j = 1 To x - 2
            graytone = (ImagePixels(0, i, j) + ImagePixels(1, i, j) + ImagePixels(2, i, j)) / 3
            ImagePixels(0, i, j) = graytone
        Next
        DoEvents
    Next
    
   
    
    For i = 1 To y - 2
        For j = 1 To x - 2
            If ImagePixels(0, i, j) < cx0 Then
                value = CDbl(ImagePixels(0, i, j)) * cm1
            ElseIf ImagePixels(0, i, j) < cx1 Then
                value = CDbl(ImagePixels(0, i, j)) * cm2 + cc2
            Else
                value = CDbl(ImagePixels(0, i, j)) * cm3 + cc3
            End If
            If value > 255 Then value = 255
            If value < 0 Then value = 0
            SetPixelV hDestDC, j, i, RGB(value, value, value)
        Next
        Form4.ProgressBar1.value = i * 100 / (y - 1)
        DoEvents
    Next
    
    Form4.Hide
    BitBlt Picture1.hdc, 1, 1, Picture1.ScaleWidth - 2, Picture1.ScaleHeight - 2, hDestDC, 1, 1, &HCC0020
    Picture1.Refresh
    Call DeleteDC(hDestDC)
    Call DeleteObject(HBmp)

' UNCOMMENT NEXT LINE TO TIME OPERATION
    MsgBox "Processing completed in " & Format(Timer - t1, "##.000")

End Sub
