VERSION 5.00
Begin VB.Form frmFireworks 
   Caption         =   "Fireworks +"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   Icon            =   "frmFireworks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmFireworks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

':: THE FRAME RATE LIMITER IS OPTIONAL
#Const WITH_LIMITER = 1

':: used to override VB method of "Me.Caption"... surely faster, so less overhead
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

':: size of the display area
Private wnd_cx As Long, wnd_cy As Long
':: our DIB object and scanline width
Private mDIB As New GSDIB_lite, sRB As Long

':: these are explained in "launchRocket" sub
Private bit_px() As Double, bit_py() As Double
Private bit_vx() As Double, bit_vy() As Double
Private bit_sx() As Long, bit_sy() As Long
Private bit_l() As Long, bit_f() As Long
Private bit_p() As Long, bit_c() As Long

':: mixing rates for the fading
':: RXV is 253 - RV since we are adding some hardcoded
'   fading to the black tone. Try setting it to 256 and
'   see how the screen remains colored for a bit of time
Private Const RU As Long = 64, RV As Long = 64, RB As Long = 128
Private Const RXU As Long = 256 - RU, RXV As Long = 253 - RV, RXB As Long = 256 - RB

':: maximum number of 'alive' particles we can have at a given time
Private Const Bits As Long = 15000
':: maximum density of our fireworks
Private Const BIT_MAX As Long = 300
':: how smaller can be our fireworks, try with 290 and see
'   some tiny fireworks in action...
Private Const BIT_VARIANCE As Long = 220

':: FPS counter
Private curTim As Long, FPS As Long
':: FPS limiter
Private idleTicks As Long
':: timer to control automatic launching of fireworks
Private tmrLaunch As Long

':: the rect we use to force the update
Private rc As RECT

Private Sub Form_Load()
    ':: the rect we use to ask for a repaint
    rc.Right = 1: rc.Bottom = 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ':: launch a rocket where the user clicked
    launchRocket X, Y
End Sub

Private Sub Form_Paint()
    ':: render next frame
    OnRender
    
    ':: queue our next frame
    ':: this trick acts as if we have a second thread
    '   since it acts in parallel with our app code.
    '   the trick is we tell windows we need to update
    '   our display area, so it sends a WM_PAINT message
    '   as soon as it is idle
    InvalidateRect hwnd, rc, False
End Sub

Private Sub Form_Resize()
    ':: added resizing ability to allow better performance
    '   on slow systems... and also some insane FPS!
    '   i get 260 FPS @ 200x200 on my Duron 650 + TNT2 pro
    
    wnd_cx = Me.ScaleWidth
    wnd_cy = Me.ScaleHeight
    
    ':: create the drawing surface
    mDIB.Create wnd_cx, wnd_cy
    sRB = mDIB.BytesPerScanLine
    
    Call InitFireworks
    
End Sub

Private Sub Form_Terminate()
    ':: clean-up everything
    mDIB.Destroy
    Set mDIB = Nothing
    
    Erase bit_px, bit_py
    Erase bit_vx, bit_vy
    Erase bit_sx, bit_sy
    Erase bit_l, bit_f, bit_p, bit_c

    Set frmFireworks = Nothing
End Sub

Private Sub InitFireworks()
    ':: allocate enough memory to hold our particles
    ReDim bit_px(Bits)
    ReDim bit_py(Bits)
    ReDim bit_vx(Bits)
    ReDim bit_vy(Bits)
    ReDim bit_sx(Bits)
    ReDim bit_sy(Bits)
    ReDim bit_l(Bits)
    ReDim bit_f(Bits)
    ReDim bit_p(Bits)
    ReDim bit_c(Bits)
    
    tmrLaunch = GetTickCount + 500
End Sub

Private Sub OnRender()
Dim i As Long, j As Long
Dim r1 As Long, g1 As Long, b1 As Long
Dim r2 As Long, g2 As Long, b2 As Long

Dim js As Long

Dim s As Long
Dim foo As Long

Dim sB() As Byte, Wi As Long
Dim sBL() As Long
        
        
#If WITH_LIMITER Then
    ':: FPS limiter, currently set at around 70-73 FPS
    '   final speed depends on the precision of your PC
    '   and windows version... just didn't want to go
    '   with the hi-res timer...
    Do Until GetTickCount >= idleTicks
    Loop
    idleTicks = GetTickCount + 14
#End If
    
    ':: FPS counter [fixed, it should count the frames better]
    If GetTickCount >= curTim Then
        SetWindowText hwnd, "Fireworks +  [" & FPS & " FPS]"
        FPS = 0
        curTim = GetTickCount + 1000
    End If
    FPS = FPS + 1

    ':: map our arrays to the DIB bits
    mDIB.MapArray sB
    mDIB.MapArrayLong sBL
    ':: we are using here the classic SafeArray trick, while
    '   using two mapped arrays to the same memory location.
    '   this has its advatages that we need a simple assignment
    '   to change a given pixel (using the long array), while
    '   we can access the RGB pixel components individually
    '   with the byte array.

    ':: minor speedup by saving this value
    Wi = sRB - 4

    ':: try to launch a rocket if 500 ms have passed
    ':: the Rnd Mod 3 causes some rockets not to launch
    '   so they launch with a bit of randomness
    If GetTickCount > tmrLaunch Then
        If Rnd Mod 3 Then
            launchRocket (Rnd * wnd_cx), (Rnd * (wnd_cy - 8)) + 5
        End If
        tmrLaunch = GetTickCount + 500
    End If
    
'    // fade code
    
    ':: the fading is really a bluring function, that only
    '   takes some pixels into account. It also gives different
    '   weight values to the pixels, so it is also in part a
    '   blending function. For gods simplicity we will refer
    '   to it just as the bluring function.
    ':: my current implementation is a mod from the original.
    '   as given by the previous author we had this format for
    '   the bluring:
    
    '  - x +
    '  - x +
    '  - - -
    
    ' where - is an ignored pixel
    '       x is an important bluring pixel
    '       + is a less important bluring pixel
    
    ':: This kind of bluring had a visual effect as if it was
    '   some wind blowing from the right, but it certainly not
    '   looked uniform: the rocket flyed vertically and only
    '   the smoke was affected by the 'wind'
    '   Also the smoke trails for the particles faded to the
    '   left, even if the particles were unaffected by the
    '   misterious wind.
    
    ':: my actual mod does the following with the pixels:
    '  - + -
    '  * x *
    '  - - -
    
    ' where '*' are third priority pixels for the bluring.
    
    ':: This new bluring process, along with the hardcoded
    '   fading (see the declares section), make for a much
    '   more realistic effect than the original one.
    
    For s = 0 To (sRB * (wnd_cy - 2)) Step sRB
        ':: clear the right and left-most pixel to prevent a fall-through from the sides
        sBL(j \ 4) = 0
        sBL((j + Wi) \ 4) = 0
        
        For j = s + 4 To s + Wi Step 4
            ':: we save this value to increase speed
            js = j + sRB
            
            ':: mix the pixels with the pre-selected rate
            r1 = (sB(js) * RU + sB(j) * RXU)
            g1 = (sB(js + 1) * RU + sB(j + 1) * RXU)
            b1 = (sB(js + 2) * RU + sB(j + 2) * RXU)
            
            r2 = (sB(j - 4) * RB + sB(j + 4) * RXB)
            g2 = (sB(j - 3) * RB + sB(j + 5) * RXB)
            b2 = (sB(j - 2) * RB + sB(j + 6) * RXB)
            
            ':: i optimized the code a lot here by simplifying
            '   the math in the calculations. Basically i am
            '   doing another blur/blend
            ':: see the original post for a more complex (but
            '   a little easier to follow) way of doing this.
            sBL(j \ 4) = (((b2 * RV + b1 * RXV)) And &HFF0000) + _
                         (((g2 * RV + g1 * RXV) \ 256&) And &HFF00&) + _
                         (r2 * RV + r1 * RXV) \ 65536
        Next
    Next

'   // render code
    For foo = 0 To Bits - 1
        Select Case bit_f(foo)
            Case 1
                ':: gravity acts on this particle
                bit_vy(foo) = bit_vy(foo) + (Rnd * 0.02)
                ':: move it and make it fall
                bit_px(foo) = bit_px(foo) + bit_vx(foo)
                bit_py(foo) = bit_py(foo) + bit_vy(foo)
                ':: decrement the particle's life
                bit_l(foo) = bit_l(foo) - 1
                ':: check if the particle has passed-away (dead ^_^) or
                '   it felt outside the drawing region
                If (bit_l(foo) = 0) Or (bit_px(foo) < 0) Or (bit_py(foo) < 1) Or (bit_px(foo) > (wnd_cx - 1)) Or (bit_py(foo) > (wnd_cy - 2)) Then
                    ':: clear it and set it as available
                    bit_c(foo) = 0
                    bit_f(foo) = 0
                Else
                    ':: check if it's a colored or white pixel
                    If (bit_p(foo) = 0) Then
                        ':: this adds some blink to the white particles
                        If (Round(Rnd) = 0) Then
                             sBL(CLng(bit_px(foo)) + ((wnd_cy - CLng(bit_py(foo)) - 1) * wnd_cx)) = -1
                        End If
                    Else
                        ':: draw the colored particle
                        sBL(CLng(bit_px(foo)) + ((wnd_cy - CLng(bit_py(foo)) - 1) * wnd_cx)) = bit_c(foo)
                    End If
                End If
            Case 2
                ':: elevate the rocket
                bit_sy(foo) = bit_sy(foo) - 5
                ':: check if we are to explode
                If (bit_sy(foo) <= bit_py(foo)) Then bit_f(foo) = 1
                ':: we draw some of the pixels with a bit of
                '   randomness to better emulate the smoke
                If Round(Rnd * 10#) = 0 Then
                    ':: add some waving to the rocket smoke trail
                    i = Round(Rnd * 2#)
                    j = Round(Rnd * 4#)
                    sBL(CLng(bit_sx(foo) + i) + ((wnd_cy - CLng(bit_sy(foo) + j) - 1) * wnd_cx)) = -1
                End If
        End Select
    Next
    
    ':: get the arrays unmapped, this avoids crashes in WinNT
    mDIB.UnMapArray sB
    mDIB.UnMapArrayLong sBL
    
    ':: put it to screen
    mDIB.PaintTo frmFireworks.hDC
    
End Sub

Private Sub launchRocket(ByVal X As Long, ByVal Y As Long)
Dim r As Long, g As Long, b As Long
Dim i As Long, pColor As Long, bitCount As Long
Dim D As Double, d1 As Double
Dim nBitMax As Long
    
    ':: do not allow a rocket at the top of the screen
    If Y <= 1 Then Y = 2
    
    ':: improve randomness
    Randomize Rnd
    
    ':: add some randomness to the size of the fireworks
    nBitMax = (BIT_MAX - BIT_VARIANCE) + (Rnd * BIT_VARIANCE)
        
    ':: pick a random color
    ':: we provide a minimum color, so we do not get a
    '   too dark color... fireworks have vivid colors!
    r = (Rnd * 226) + 30
    g = (Rnd * 240) + 16
    b = (Rnd * 220) + 36
    pColor = (r * 65536) + (g * 256&) + b

    For i = 0 To Bits - 1
        ':: check if it is an unassigned slot
        If bit_f(i) = 0 Then
            ':: locate the particle
            bit_px(i) = X
            bit_py(i) = Y
            ':: check an expansion rate for the explosion
            D = Rnd * 6.28
            d1 = Rnd
            bit_vx(i) = Sin(D) * d1
            bit_vy(i) = Cos(D) * d1
            ':: pick a random life for the particle
            bit_l(i) = (Rnd * 100#) + 100
            ':: pick a particle color, may be the chosen one, or white
            '   This effect adds some sparkles to the firework explosion
            bit_p(i) = Round(Rnd * 2#)
            ':: save the previously choosen color
            bit_c(i) = pColor
            ':: position of the rocket on screen (before it explodes)
            bit_sx(i) = X
            bit_sy(i) = wnd_cy - 5
            ':: set this as an ascending rocket
            bit_f(i) = 2
            ':: increment number of allocated particles,
            '   and exit if we reached the max for this one.
            bitCount = bitCount + 1
            If (bitCount = nBitMax) Then Exit For
        End If
    Next
End Sub

