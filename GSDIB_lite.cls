VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSDIB_lite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'########################################################
'#                                                      #
'# this is a little sample of the GSDIB Graphics Engine #
'#                                                      #
'#   i posted a long time ago a project based on this   #
'#    engine, and this time i decided to show a tiny    #
'#     part of it, and its possible implementations     #
'#                                                      #
'#  The current use of it is under my personal image    #
'#  image viewer, that will probably not come to light  #
'#  until the end of this year...                       #
'#  I'm also using the engine in a pac-man clone game   #
'#  but i have discontinued it a long time ago.         #
'#  Most recent and shocking use of it can be in my new #
'#  updated NVIDIA tweaking tool: NVtweaker             #
'#  I have, however, discontinued its development due   #
'#  to the lack of time. I have however created a full  #
'#  virtual interface based on the GSDIB engine, so that#
'#  the program has full skinning capabilities.         #
'#                                                      #
'#  The engine is on sale, but don't come to me with a  #
'#  number smaller than $2000.                          #
'#  The full engine includes:                           #
'#      - primitives drawing                            #
'#      - text drawing (with gradients and textures)    #
'#      - flip, rotate, resize, scroll, mirror          #
'#      - greyscaling, monochroming (2 dither patterns) #
'#      - blur, sharpen, diffuse                        #
'#      - 3x3 convolution matrix effects                #
'#      - Alpha Blending completely written in assembly #
'#        code for maximum speed. It comes with several #
'#        flavors for the blending to maximize the      #
'#        engine capabilities.                          #
'#      - Brightness, Contrast, Intensity, Gamma        #
'#      - RGB channels management (kill, swap, shift)   #
'#      - Light Mapping and Bump Mapping written in asm #
'#      - Minimum and Maximum noise filters             #
'#      - Scale2x resampling                            #
'#      - Oil Painting, Scanlines, Zooming effect...    #
'#      - Loading of BMP, TGA (uncompressed), GIF, JPG  #
'#        PNG (non interlaced), PPM, PGM, PBM and XPM   #
'#        image formats.                                #
'#                    AND MORE!!!                       #
'#                                                      #
'#    Please note that i have been working on it for    #
'#  19 months, and i put a lot of effort on it, that's  #
'#  why i won't be giving it for free.                  #
'#                                                      #
'#  DON'T FLAME ME, this is NOT ADVERTISING, it is to   #
'#  prevent people for asking for the full source.      #
'#   In fact, im not really planning to sell it, as it  #
'#   still needs a lot of work, and there are a lot of  #
'#   filters i never added due to my great laziness.    #
'#  And as you have seen, lots of my programs depend on #
'#  the use of it, so i will keep it as my own engine.  #
'#                                                      #
'#  Reuse and redistribution of this file is permitted  #
'#  as long as this sign remains intact & proper credit #
'#  is given.                                           #
'#                                                      #
'#  Copyright 2001-2003 by gonchuki                     #
'#  gonchuki@yahoo.es                                   #
'#  http://gonchuki.8m.com                              #
'########################################################

Private MyhDC As Long, lpRGB As Long, sRB As Long
Private bmH As BITMAPINFOHEADER
Private hBMO As Long, hDIB As Long

Private sa As SAFEARRAY2D

Public Sub Create(ByVal W As Long, ByVal H As Long)
'Creates or resizes the current DIB
    Me.Destroy
    With bmH
        .biSize = Len(bmH)
        .biWidth = W
        .biHeight = H
        .biPlanes = 1
        .biBitCount = 32
            sRB = ((W * 4) And &HFFFFFFFC)
        .biSizeImage = (sRB * H)
    End With
    MyhDC = CreateCompatibleDC(0)
    hDIB = CreateDIBSection(MyhDC, bmH, 0, lpRGB, 0, 0)
    If hDIB Then
        hBMO = SelectObject(MyhDC, hDIB)
        sa.Bounds(0).cElements = bmH.biSizeImage
        Cls
    Else
        Err.Raise 825, , "DIB failed to create." & vbCrLf & "GDI Resources may be low." & vbCrLf & vbCrLf & "Please restart your system and try again."
    End If
    SetStretchBltMode MyhDC, 3
End Sub

Public Sub Clone(ByVal Pic As StdPicture, Optional ByVal FitSize As Boolean = True)
'Clones a Picture object, you may use LoadPicture to
'directly create the DIB without calling Create()

    Dim bmp As BITMAP
    
    GetObjectA Pic.handle, Len(bmp), bmp
    If FitSize Then Create bmp.bmWidth, bmp.bmHeight
    If hDIB Then
        Dim hDCt As Long, hBMOt As Long
        hDCt = CreateCompatibleDC(MyhDC)
        hBMOt = SelectObject(hDCt, Pic.handle)
        StretchBlt MyhDC, 0, 0, bmH.biWidth, bmH.biHeight, hDCt, 0, 0, bmp.bmWidth, bmp.bmHeight, vbSrcCopy
        SelectObject hDCt, hBMOt
        DeleteDC hDCt
    Else
        Err.Raise -1, , "DIB has to be created first"
    End If
End Sub

Public Sub CloneHDC(ByVal theHDC As Long, Optional ByVal dwRop As Long = SRCCOPY, Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0)
'clones an entire hDC or a selected area
'useful for image cropping/copying
    BitBlt MyhDC, 0, 0, bmH.biWidth, bmH.biHeight, theHDC, X, Y, dwRop
End Sub

Public Sub CloneDIB(ByVal theDIB As GSDIB_lite)
    Me.Create theDIB.Width, theDIB.Height
    BitBlt MyhDC, 0, 0, bmH.biWidth, bmH.biHeight, theDIB.hDC, 0, 0, SRCCOPY
End Sub

Public Sub PaintTo(ByVal toDC As Long, Optional ByVal atX As Long = 0, Optional ByVal atY As Long = 0, Optional ByVal dwRop As Long = SRCCOPY)
'paints the entire picture to selected hDC
    BitBlt toDC, atX, atY, bmH.biWidth, bmH.biHeight, MyhDC, 0, 0, dwRop
End Sub

Public Sub Cls()
    If hDIB Then ZeroMemory ByVal lpRGB, bmH.biSizeImage
End Sub

Public Sub Destroy()
If MyhDC Then
    Cls
    If hBMO Then DeleteObject SelectObject(MyhDC, hBMO): hBMO = 0
    DeleteDC MyhDC: MyhDC = 0
End If
End Sub

Private Sub Class_Initialize()
    
    sa.cDims = 1
    sa.cbElements = 1
    
End Sub

Private Sub Class_Terminate()
    Me.Destroy
End Sub

Public Property Get RGBsize() As Long
    RGBsize = bmH.biSizeImage
End Property

Friend Property Get RGBpointer() As Long
    RGBpointer = lpRGB
End Property

Public Property Get Width() As Long
    Width = bmH.biWidth
End Property

Public Property Get Height() As Long
    Height = bmH.biHeight
End Property

Friend Sub MapArray(ByRef sB() As Byte)
    sa.pvData = lpRGB
    CopyMemory ByVal VarPtrArray(sB), VarPtr(sa), 4
End Sub

Friend Sub MapArrayLong(ByRef sB() As Long)
    sa.pvData = lpRGB
    CopyMemory ByVal VarPtrArray(sB), VarPtr(sa), 4
End Sub

Friend Sub UnMapArray(ByRef sB() As Byte)
    CopyMemory ByVal VarPtrArray(sB), 0&, 4
End Sub

Friend Sub UnMapArrayLong(ByRef sB() As Long)
    CopyMemory ByVal VarPtrArray(sB), 0&, 4
End Sub

Public Property Get BytesPerScanLine() As Long
    'Scans must align on dword boundaries:
    BytesPerScanLine = (bmH.biWidth * 4) And &HFFFFFFFC
End Property

Public Property Get DIBits() As Byte()
    Dim sB() As Byte, Bits() As Byte
    Call Me.MapArray(sB)
    ReDim Bits(0 To bmH.biSizeImage - 1)
    CopyMemory Bits(0), sB(0), bmH.biSizeImage
    DIBits = Bits
End Property

Public Property Let DIBits(ByRef newBits() As Byte)
    Dim sB() As Byte
    Me.MapArray sB
    CopyMemory sB(0), newBits(0), bmH.biSizeImage
End Property

Public Property Get hDC() As Long
    hDC = MyhDC
End Property

