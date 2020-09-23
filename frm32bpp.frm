VERSION 5.00
Begin VB.Form frm32bpp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Redraw Using Auto-Alpha Detection"
      Enabled         =   0   'False
      Height          =   645
      Left            =   4350
      TabIndex        =   12
      Top             =   2835
      Width           =   2265
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   4425
      Width           =   7425
   End
   Begin VB.CheckBox chkImage 
      Caption         =   "Test Image 2"
      Height          =   360
      Index           =   1
      Left            =   1830
      TabIndex        =   10
      Top             =   3555
      Width           =   1530
   End
   Begin VB.CheckBox chkImage 
      Caption         =   "Test Image 1"
      Height          =   360
      Index           =   0
      Left            =   150
      TabIndex        =   9
      Top             =   3555
      Value           =   1  'Checked
      Width           =   1530
   End
   Begin VB.CommandButton cmdRender 
      Caption         =   "GDI+  paRGB"
      Height          =   540
      Index           =   3
      Left            =   5610
      TabIndex        =   7
      Top             =   2265
      Width           =   1800
   End
   Begin VB.CommandButton cmdRender 
      Caption         =   "GDI+  aRGB"
      Height          =   540
      Index           =   2
      Left            =   3765
      TabIndex        =   6
      Top             =   2265
      Width           =   1800
   End
   Begin VB.CommandButton cmdRender 
      Caption         =   "AlphaBlend (SrcAlpha)"
      Height          =   540
      Index           =   1
      Left            =   5610
      TabIndex        =   5
      Top             =   1710
      Width           =   1800
   End
   Begin VB.CommandButton cmdRender 
      Caption         =   "VB's PaintPicture"
      Height          =   540
      Index           =   0
      Left            =   3765
      TabIndex        =   4
      Top             =   1710
      Width           =   1800
   End
   Begin VB.OptionButton Option1 
      Caption         =   "32bpp ARGB (RGB is preMultiplied)"
      Height          =   360
      Index           =   2
      Left            =   3765
      TabIndex        =   3
      Top             =   1215
      Width           =   3690
   End
   Begin VB.OptionButton Option1 
      Caption         =   "32bpp ARGB (Not preMultiplied RGB)"
      Height          =   360
      Index           =   1
      Left            =   3765
      TabIndex        =   2
      Top             =   735
      Width           =   3690
   End
   Begin VB.OptionButton Option1 
      Caption         =   "32bpp RGB (No Alpha channel used)"
      Height          =   360
      Index           =   0
      Left            =   3735
      TabIndex        =   1
      Top             =   240
      Width           =   3690
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0FF&
      Height          =   3390
      Left            =   60
      ScaleHeight     =   3330
      ScaleWidth      =   3570
      TabIndex        =   0
      Top             =   120
      Width           =   3630
   End
   Begin VB.Label Label1 
      Height          =   630
      Left            =   105
      TabIndex        =   8
      Top             =   3990
      Width           =   7545
   End
End
Attribute VB_Name = "frm32bpp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' For everyone but Win95 users?  If you want GDI+, download it from Microsoft:
' http://www.microsoft.com/downloads/details.aspx?familyid=6a63ab9c-df12-4d41-933c-be590feaa05a&displaylang=en

Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY1D        ' used as DMA overlay on a DIB
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound As SafeArrayBound
End Type
Private Type BITMAP             ' used to get stdPicture attributes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type BITMAPINFOHEADER   ' used to create a DIB
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiPalette(0 To 255) As Long
End Type
Private Type BLENDFUNCTION      ' used for AlphaBlend API
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type
Private Const AC_SRC_OVER = &H0
Private Const AC_SRC_ALPHA = &H1

Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal lBlendFunction As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

' used for workaround of VB not exposing IStream interface
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long

' GDI+ calls
Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, image As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipDrawImage Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal X As Single, ByVal Y As Single) As Long
Private Const PixelFormat32bppARGB As Long = &H26200A
Private Const PixelFormat32bppPARGB As Long = &HE200B
Private Const PixelFormat32bppRGB As Long = &H22009

Private gdiToken As Long            ' gdi+ token (must be released)
Private picTest As StdPicture



Private Function IsAlpha(ByVal imageHandle As Long, ByRef isPARGB As Boolean) As Boolean

    ' Returns whether or not a 32bpp bitmap should be handled as alpha or not
    ' Additionally, it can identify if the image is not pre-multiplied in most cases:

    ' [IN] imageHandle :: stdPicture or DIB handle
    ' [OUT] isPARGB :: (see DISCLAIMER below)
    '                  when true, image may have been pre-multiplied
    '                  when false, the image is definitely not pre -multiplied
    ' [OUT] Return value. (see DISCLAIMER below)
    '       True indicates the image is using alpha channels
    '       False indicates the image is not using alpha channels

    ' Terms:
    '       RGB = alphas are not used (either all=0 or all=255) or image is not 32bpp
    '       ARGB = non-premultiplied RGB values w/alpha used
    '       pARGB = premultiplied RGB values w/Alpha used

    ' DISCLAIMERS:

    '   PreMultiplied Assumption: There is no fool-proof way to determine if an image
    '   is premultiplied. Some images can be determined to be not premultiplied,
    '   but no image can be determined to be premutliplied without prior knowledge.
    '   Examples of RGB bytes and an Alpha byte
    '       1. Pixel 255,255,255,A:50. Since any RGB byte is > Alpha byte, cannot
    '          be premultiplied. Formula:  RGB byte = (Alpha byte * RGB byte)\255
    '       2. Pixel 200,150,50,A:210. May be premultiplied, may not be either
    '       3. Pixel 255,255,255,A:255. May be either & doesn't matter since both
    '          premultiplied formula and non-premultiplied results in same overall Color
    '   If image is not RGB & cannot be determined 100% as ARGB, then pARGB is True
    '   Through testing, it appears that AlphaBlend and GDI+ will handle example 2
    '   just fine even if the image isn't pARGB and we say it is. That is a nice finding

    '   Alpha Assumption: There is also no fool proof way to determine if a image with
    '   all alpha values set to zero, whether the alpha values are not used or the image
    '   is 100% transparent ARGB. The assumption is that an image would not be created
    '   and stored that is completely transparent. Fading routines could Change alpha
    '   values but odds are that the image would not be stored as transparent. Therefore,
    '   when every alpha value is zero, this routine's return value is False

    ' So what does this code do for you? It will make an educated guess as to which
    ' alpha properties apply to the passed 32bpp image. This will help you determine
    ' which flags are needed to draw the image.
    '   For AlphaBlend: this is the BLENDFUNCTION.AlphaFormat flag
    '   For GDI+: this is one of the PixelFormat32bppXXXX flags

    ' Possible Results of the For:Next loop below:
    ' ZeroCount = max :: either 100% transparent ARGB or RGB (will assume RGB, not ARGB)
    '   max ZeroCount is width*height
    '   ^^ FYI: if lPARGB=0 also, then the image is black
    ' lPARGB = 2 :: definitely ARGB
    ' lPARGB = 1 :: meaning is reliant on ZeroCount
    '               if ZeroCount<max then it is definitely ARGB
    '               if ZeroCount=max then is not black, is also ARGB and/or transparent
    ' lPARGB = 0 :: all RGB bytes are <= their alpha byte. Therefore, it can mean
    '               image either ARGB or pARGB (no way of knowing for sure).
    '               Prior knowledge is absolutely required to be 100% positive.
    '               This is the catch. If your user cannot tell you the image is
    '               premultiplied or not, then an algorithmic "best" guess is to
    '               use pARGB since AlphaBlend & GDI+ do not seem to mind and
    '               draw the image correctly. At least on the ones I have tested
    ' Last  Note :: should all alpha values = 255, it is always safe to assume
    '               all 3 conditions: RGB, ARGB, pARGB

    Dim lPos As Long
    Dim X As Long, Y As Long, R As Long

    Dim lPARGB As Long, bIsAlpha As Boolean, zeroCount As Long

    Dim tBMP As BITMAP, tSA As SAFEARRAY1D
    Dim dibBytes() As Byte

    ' validdate image is a bitmap
    If GetGDIObject(imageHandle, Len(tBMP), tBMP) = 0 Then
        isPARGB = False
        Exit Function
    ElseIf tBMP.bmBits = 0 Or tBMP.bmBitsPixel <> 32 Then
    ' validate image is 32 bits and we can access the DIB
        isPARGB = False
        Exit Function
    End If

    With tSA ' prepare array overlay
        .cbElements = 1 ' byte elements
        .cDims = 1      ' single dim array
        .pvData = tBMP.bmBits
        .rgSABound.cElements = tBMP.bmWidthBytes * tBMP.bmHeight
    End With
    ' overlay now
    CopyMemory ByVal VarPtrArray(dibBytes()), VarPtr(tSA), 4&

    ' loop thru the bytes, trying to determine if image is NOT premultiplied
    For Y = 0 To Abs(tBMP.bmHeight) - 1
        lPos = Y * tBMP.bmWidthBytes
        For X = lPos + 3 To lPos + tBMP.bmWidthBytes - 1 Step 4
            Select Case dibBytes(X)
            Case 0
                If lPARGB = 0 Then
                    ' zero alpha, if any of the RGB bytes are non-zero, then this is not pre-multiplied
                    If Not dibBytes(X - 1) = 0 Then
                        lPARGB = 1 ' not premultiplied
                    ElseIf Not dibBytes(X - 2) = 0 Then
                        lPARGB = 1
                    ElseIf Not dibBytes(X - 3) = 0 Then
                        lPARGB = 1
                    End If
                    ' but don't exit loop until we know if any alphas are non-zero
                End If
                zeroCount = zeroCount + 1 ' helps in decision factor at end of loop
            Case 255
                ' no way to indicate if premultiplied or not, unless...
                If lPARGB = 1 Then
                    lPARGB = 2    ' not pre-multiplied because of the zero check above
                    Exit For
                End If
            Case Else
                ' if any Exit For's below get triggered, not pre-multiplied
                If lPARGB = 1 Then
                    lPARGB = 2: Exit For
                ElseIf dibBytes(X - 3) > dibBytes(X) Then
                    lPARGB = 2: Exit For
                ElseIf dibBytes(X - 2) > dibBytes(X) Then
                    lPARGB = 2: Exit For
                ElseIf dibBytes(X - 1) > dibBytes(X) Then
                    lPARGB = 2: Exit For
                End If
            End Select
        Next
        If lPARGB = 2 Then Exit For
    Next
    ' remove overlay
    CopyMemory ByVal VarPtrArray(dibBytes()), 0&, 4&
    
    If zeroCount = tBMP.bmWidth * tBMP.bmHeight Then ' every alpha value was zero
        isPARGB = False: bIsAlpha = False ' assume RGB, else 100% transparent ARGB
        ' also if lPARGB=0, then image is completely black
    Else
        Select Case lPARGB
            Case 2: isPARGB = False: bIsAlpha = True ' 100% positive ARGB
            Case 1: isPARGB = False: bIsAlpha = True ' now 100% positive ARGB
            Case 0: isPARGB = True: bIsAlpha = True
            ' ^^ set isPARGB so you can decide whether to premultiply or not
        End Select
    End If

    IsAlpha = bIsAlpha

End Function

Private Function GetPreMultipliedImage(ByVal imageHandle As Long) As Long

    ' DESTRUCTION OF ORIGINAL PIXELS WILL ALWAYS OCCUR WHEN PREMULTIPLYING BYTES
    ' Therefore, if original pixels are needed, you should copy the image and
    ' work on that copy, not the original. This premultiplies a "copy"

    Dim lPos As Long, hDib As Long, hDibPtr As Long
    Dim X As Long, Y As Long, R As Long

    Dim tBMP As BITMAP, tSA As SAFEARRAY1D, dSA As SAFEARRAY1D
    Dim tBMPI As BITMAPINFO
    Dim dibBytes() As Byte, srcBytes() As Byte

    ' validdate image is a bitmap
    If GetGDIObject(imageHandle, Len(tBMP), tBMP) = 0 Then Exit Function
    If tBMP.bmBits = 0 Or tBMP.bmBitsPixel <> 32 Then Exit Function

    With tBMPI.bmiHeader
        .biBitCount = 32
        .biHeight = tBMP.bmHeight
        .biWidth = tBMP.bmWidth
        .biPlanes = 1
        .biSize = 40 ' len(tBMPI.bmiHeader)
    End With
    
    hDib = CreateDIBSection(Me.hDC, tBMPI, 0, hDibPtr, 0, 0)
    If hDib = 0 Then Exit Function
    
    With tSA ' prepare array overlay
        .cbElements = 1 ' byte elements
        .cDims = 1      ' single dim array
        .pvData = tBMP.bmBits
        .rgSABound.cElements = tBMP.bmWidthBytes * tBMP.bmHeight
    End With
    CopyMemory ByVal VarPtrArray(srcBytes()), VarPtr(tSA), 4&
    
    dSA = tSA
    dSA.pvData = hDibPtr
    CopyMemory ByVal VarPtrArray(dibBytes()), VarPtr(dSA), 4&
    
    For Y = 0 To Abs(tBMP.bmHeight) - 1
         lPos = Y * tBMP.bmWidthBytes
         For X = lPos + 3 To lPos + tBMP.bmWidthBytes - 1 Step 4
             If srcBytes(X) = 0 Then
                 CopyMemory dibBytes(X - 3), 0&, &H4
             ElseIf dibBytes(X) = 255 Then
                 CopyMemory dibBytes(X - 3), srcBytes(X - 3), &H4
             Else
                 For R = X - 3 To X - 1
                     dibBytes(R) = ((0& + srcBytes(R)) * srcBytes(X)) \ &HFF
                 Next
                 dibBytes(X) = srcBytes(X)
             End If
         Next
     Next
    ' remove overlays
    CopyMemory ByVal VarPtrArray(srcBytes()), 0&, 4&
    CopyMemory ByVal VarPtrArray(dibBytes()), 0&, 4&
    
    ' return DIB handle -- don't forget to destroy it
    GetPreMultipliedImage = hDib

End Function



Private Sub cmdRender_Click(Index As Integer)
    
    If picTest Is Nothing Then Exit Sub
    
    Dim bAlpha As Boolean, bPARGB As Boolean ' < used when auto-detecting alpha
    
    Picture1.Cls
    Command1.Enabled = (Abs(Index) <> 0)
    
    Select Case Abs(Index) ' negative Index indicates Auto-Detect button clicked
    
    Case 0: ' vb's paintpicture (bitblt, gpiLoadImageFromFile will do same thing)
        Picture1.PaintPicture picTest, 0, 0
        
    Case 1: ' alphablend. With .AlphaFormat=0, same results as BitBlt
        Dim bf As BLENDFUNCTION, lBlend As Long, tDC As Long, tOldBmp As Long
        Dim imgCx As Long, imgCy As Long, hPreMul As Long, imgHandle As Long
        
        ' may be premultiplying a copy of the image for AlphaBlend's use.
        imgHandle = picTest.handle
        
        With bf ' fill in the blend function
            If Index < 0 Then   ' auto-detect
                If IsAlpha(picTest.handle, bPARGB) = True Then
                    If bPARGB = False Then  ' not premultiplied
                        ' AlphaBlend requires premultiplied, make it so
                        hPreMul = GetPreMultipliedImage(picTest.handle)
                        If hPreMul = 0 Then
                            ' shouldn't happen unless low on resources?
                            MsgBox "Failed to create necessary image", vbExclamation + vbOKCancel
                            Exit Sub
                        End If
                        imgHandle = hPreMul ' use this handle for AlphaBlend
                    End If
                    .AlphaFormat = AC_SRC_ALPHA ' image is premultiplied
                End If
            Else        ' not auto-detecting
                .AlphaFormat = AC_SRC_ALPHA ' assume image is premultiplied; can go wrong either way
            End If
            .BlendOp = AC_SRC_OVER
            .SourceConstantAlpha = 255
        End With
        CopyMemory lBlend, bf, 4&
        
        ' get image width/height
        imgCx = ScaleX(picTest.Width, vbHimetric, vbPixels)
        imgCy = ScaleY(picTest.Height, vbHimetric, vbPixels)
        ' select image into a DC
        tDC = CreateCompatibleDC(Me.hDC)
        tOldBmp = SelectObject(tDC, imgHandle)
        
        On Error Resume Next
        AlphaBlend Picture1.hDC, 0, 0, imgCx, imgCy, tDC, 0, 0, imgCx, imgCy, lBlend
        
        SelectObject tDC, tOldBmp
        DeleteDC tDC
        If hPreMul <> 0 Then DeleteObject hPreMul
        
        If Err Then ' Win95 system? Shouldn't get an error w/AlphaBlend on any other O/S
            ' prevent this button from being used again
            Command1.Enabled = False
            cmdRender(Abs(Index)).Enabled = False
            Index = 0 ' so routines don't try to set focus on this btn
            MsgBox Err.Description, vbCritical + vbOKOnly, "Disabling this button"
            Err.Clear
        End If
        
    Case 2, 3: ' GDI+ without pre-multiplied RGB flag
        Dim tImage As Long, tGraphics As Long ' < gdi+ handles
        Dim lPixelFormat As Long, tBMP As BITMAP
        
        If Index < 0 Then   ' auto-detect
            bAlpha = IsAlpha(picTest.handle, bPARGB)
            If bAlpha = False Then  ' no alpha
                lPixelFormat = PixelFormat32bppRGB
            ElseIf bPARGB = True Then ' premultiplied
                lPixelFormat = PixelFormat32bppPARGB
            Else                        ' not premultiplied
                lPixelFormat = PixelFormat32bppARGB
            End If
        Else                ' not auto-detecting
            If Index = 2 Then
                lPixelFormat = PixelFormat32bppARGB
            Else
                lPixelFormat = PixelFormat32bppPARGB
            End If
        End If
        
        ' need image dimensions
        GetGDIObject picTest.handle, Len(tBMP), tBMP
        
        With tBMP
            ' having problems finding other GDI+ functions that will load a 32bpp
            ' image correctly. This works, but one thing must be kept in mind...
            ' When using the Scan0 call, do not get rid of the source bytes until
            ' you dispose of the image -- else crash!
            If GdipCreateBitmapFromScan0(.bmWidth, .bmHeight, .bmWidthBytes, lPixelFormat, ByVal .bmBits, tImage) = 0 Then
                ' wrap GDI+ around our destination DC
                If GdipCreateFromHDC(Picture1.hDC, tGraphics) = 0 Then
                    ' dib is bottom up, scan0 does top down, so flip it
                    GdipImageRotateFlip tImage, 6 ' flip vertically
                    GdipDrawImage tGraphics, tImage, 0!, 0!
                    GdipDeleteGraphics tImage
                End If
                GdipDisposeImage tImage
            End If
        End With
    
    End Select
    
    ShowTip Index
    cmdRender(0).Tag = Abs(Index) ' allows other routines to reset focus to this btn
    
    Picture1.Refresh
    
End Sub

Private Sub Command1_Click()
    ' auto detect
    If Val(cmdRender(0).Tag) = 0 Then
        ' this was disabled, but you get the idea...
        MsgBox "PaintPicture, BitBlt, stdPicture.Render on alpha channel images is futile.", vbInformation + vbOKOnly
    Else
        ' re-paint using auto-detection; negative index is flag for cmdRender
        Call cmdRender_Click(-cmdRender(0).Tag)
        If Me.Visible Then cmdRender(Val(cmdRender(0).Tag)).SetFocus
    End If
End Sub

Private Sub Form_Load()

    Picture1.AutoRedraw = True
    Label1.Caption = "FYI:  BitBlt, AlphaBlend w/o SrcAlpha flag, GDI+ gdipLoadImageFromFile all produce the same effects as VB's PaintPicture.  Those functions are not compatible with alpha channel bitmaps"
    
    ' see if we can use GDI+
    Dim gStartUp As GdiplusStartupInput
    gStartUp.GdiplusVersion = 1
    On Error Resume Next
    If Not GdiplusStartup(gdiToken, gStartUp) = 0 Then
        ' nope,disable those buttons
        cmdRender(2).Enabled = False
        cmdRender(3).Enabled = False
        gdiToken = 0
    End If
    If Err Then Err.Clear
    
    Option1(1).Value = True ' start with a test image
        
    ' add some initial notes/tips
    Text1 = "Two images to play with. The 1st one shows what can go wrong if you are using 32 bit alpha images without " & _
        "knowing if the alpha channel is premultiplied or not." & vbNewLine & _
        " The 2nd one is a rare type of alpha image. It is one where no individual RGB byte is > than " & _
        "the alpha byte within the entire image. Therefore, AlphaBlend & GDI+ draws it just fine whether the image's RGB is pre-multiplied or not."
        
End Sub


Private Sub Form_Terminate()
    ' from experience, shut down GDI+ in Terminate, not Unload
    If Not gdiToken = 0 Then GdiplusShutdown gdiToken
End Sub


Private Sub chkImage_Click(Index As Integer)
    If chkImage(Index) = 1 Then
        
        chkImage(Abs(Index - 1)).Value = 0  ' uncheck the other checkbox
        
        ' selecting a different image, redraw using last clicked cmdRender
        Select Case True
        Case Option1(0).Value
            Call Option1_Click(0)
        Case Option1(1).Value
            Call Option1_Click(1)
        Case Option1(2).Value
            Call Option1_Click(2)
        End Select
    
    End If
End Sub


Private Sub Option1_Click(Index As Integer)
    
    ' select different RGB, alpha format
    Set picTest = LoadPicture("") ' clear previous image, if any
    
    Dim tArray() As Byte
    ' LoadResPicture does not like 32bpp images, so we will create a bitmap from an array
    If chkImage(0) = 1 Then
        tArray() = LoadResData(101, "Custom")
    Else
        tArray() = LoadResData(102, "Custom")
    End If
    Set picTest = CreateStdPicFromArray(tArray)
    
    ' The images in RES file are ARGB (alpha with no premultiplied RGB bytes)
    ' now validate the image and change alpha format as needed
    Select Case Index
        Case 0: ChangeAlphaContent -1 ' remove alpha & validate
        Case 1: ChangeAlphaContent 0  ' validate only
        Case 2: ChangeAlphaContent 1  ' premultiply RGB & validate
    End Select
    
    Picture1.Cls
    Call cmdRender_Click(Val(cmdRender(0).Tag))
    If Me.Visible Then cmdRender(Val(cmdRender(0).Tag)).SetFocus
    
End Sub

Private Sub ChangeAlphaContent(alphaAdjuster As Integer)

    ' validates passed handle is 32bpp and removes alpha values or premultiplies RGB
    Dim tSA As SAFEARRAY1D, tBMP As BITMAP
    Dim dibBytes() As Byte
    Dim X As Long, Y As Long
    Dim bAbort As Boolean
    
    If picTest Is Nothing Then
        bAbort = True
    ElseIf GetGDIObject(picTest.handle, Len(tBMP), tBMP) = 0 Then
        bAbort = True
    ElseIf tBMP.bmBits = 0 Or tBMP.bmBitsPixel <> 32 Then
        bAbort = True
    End If
    
    If bAbort Then
        MsgBox "Invalid 32bit bitmap image", vbExclamation + vbOKOnly
        Set picTest = Nothing
        Exit Sub
    ElseIf alphaAdjuster = 0 Then ' just validating, not modifying
        Exit Sub
    End If
    
    With tSA
        .cbElements = 1 ' byte elements
        .cDims = 1      ' single dim array
        .pvData = tBMP.bmBits
        .rgSABound.cElements = tBMP.bmWidth * tBMP.bmHeight * 4&
    End With
    ' overlay now
    CopyMemory ByVal VarPtrArray(dibBytes()), VarPtr(tSA), 4&
    
    If alphaAdjuster = -1 Then
        ' remove the alpha value
        For X = 3 To tSA.rgSABound.cElements - 1 Step 4
            dibBytes(X) = 0
        Next
    Else
        ' premultiply RGBs
        For X = 3 To tSA.rgSABound.cElements - 1 Step 4
            For Y = X - 3 To X - 1
                dibBytes(Y) = ((0& + dibBytes(X)) * dibBytes(Y)) \ &HFF
            Next
        Next
    End If
    ' remove overlay
    CopyMemory ByVal VarPtrArray(dibBytes()), 0&, 4&

End Sub


Private Sub ShowTip(Index As Integer)

    Select Case Index
    Case 0: ' PaintPicture
        If Option1(0).Value Then ' no alpha values
            Text1.Text = "The image's background is black. A 32bit image with no alpha values is just like a 24bit image."
        Else ' aRGB & paRGB
            Text1.Text = "The black background should be transparent and show shadows, but it doesn't because standard GDI APIs can't do alpha values."
        End If
    Case 1: ' AlphaBlend
        If Option1(0).Value = True Then ' no alpha values
            Text1.Text = "The image's background is black. A 32bit image with no alpha values is just like a 24bit image. You would think the AlphaBlend would read the image as 100% transparent, on Win9x it is invisilbe, but on Win2K it isn't. Regardless, this obviously is not what we want. Alphablend should not be used on non-premultiplied RGB pixels."
        ElseIf Option1(1).Value = True Then ' not premultiplied
            If chkImage(1) = 0 Then ' jade image
                Text1.Text = "This is almost correct. This is not a pre-multiplied RGB image, so AlphaBlend is more or less guessing. Again, don't use AlphaBlend on non-premultiplied RGB pixels."
            Else                    ' toucan
                Text1.Text = "Although not premultiplied, this image's pixels are set in just a way that the image can be interpretted both as pre-multiplied and not premultiplied, and is drawn fine."
            End If
        Else        ' premultiplied
            Text1.Text = "This is correct. RGB bytes are pre-multiplied, so AlphaBlend can handle the blending just fine."
        End If
    Case 2, 3:
        If Option1(0).Value = True Then ' no alpha values
            Text1.Text = "GDI+ being told this is an alpha image, will interpret the image as 100% transparent because all alpha values are zero. This what I expected from AlphaBlend too, but it does things differently."
        ElseIf Option1(1).Value = True Then ' not pre-multiplied
            If Index = 2 Then
                Text1.Text = "GDI+ is using the correct flag so it knows how to draw the non-premultiplied RGB bytes."
            Else
                If chkImage(1) = 0 Then
                    Text1.Text = "GDI+ is using the wrong flag. Therefore, it is trying to interpret the RGB bytes as premultiplied."
                Else
                    Text1.Text = "GDI+ is using the wrong flag. But because this image's pixels are set in just a way that the image can be interpretted both as pre-multiplied and not premultiplied, the image is drawn fine."
                End If
            End If
        Else            ' pre-multiplied
            If Index = 3 Then
                Text1.Text = "GDI+ is using the correct flag so it knows how to draw the non-premultiplied RGB bytes."
            Else
                If chkImage(1) = 0 Then
                    Text1.Text = "GDI+ is using the wrong flag. Therefore, it is trying to interpret the RGB bytes as premultiplied."
                Else
                    Text1.Text = "GDI+ is using the wrong flag. But because this image's pixels are set in just a way that the image can be interpretted both as pre-multiplied and not premultiplied, the image is drawn fine."
                End If
            End If
        End If
    End Select
            
End Sub


Private Function CreateStdPicFromArray(bytContent() As Byte, Optional byteOffset As Long = 0) As IPicture
    
    ' function creates a IStream-compatible IUnknown
    On Error GoTo HandleError
        
        Dim o_lngLowerBound As Long
        Dim o_lngByteCount  As Long
        Dim o_hMem As Long
        Dim o_lpMem  As Long
        
        Dim aGUID(0 To 4) As Long, tStream As IUnknown
        ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
        aGUID(0) = &H7BF80980
        aGUID(1) = &H101ABF32
        aGUID(2) = &HAA00BB8B
        aGUID(3) = &HAB0C3000
        
        o_lngByteCount = UBound(bytContent) - byteOffset + 1
        o_hMem = GlobalAlloc(&H2, o_lngByteCount)
        If o_hMem <> 0 Then
            o_lpMem = GlobalLock(o_hMem)
            If o_lpMem <> 0 Then
                CopyMemory ByVal o_lpMem, bytContent(byteOffset), o_lngByteCount
                Call GlobalUnlock(o_hMem)
                If CreateStreamOnHGlobal(o_hMem, 1, tStream) = 0 Then
                    Call OleLoadPicture(ByVal ObjPtr(tStream), 0, 0, aGUID(0), CreateStdPicFromArray)
                End If
            End If
        End If
    
HandleError:
End Function

