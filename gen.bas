Attribute VB_Name = "mod_gen"
Option Explicit

Public Type ptDouble
    x As Double
    y As Double
End Type

Public Type lineDbl
    ptStart As ptDouble
    ptEnd As ptDouble
End Type

Public Type Rect
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Type RGBA
    b As Byte
    g As Byte
    r As Byte
    a As Byte
End Type

Public Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Public Type BITMAPINFOHEADER   '40 bytes
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

Public Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBA
End Type

Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0&
Public Const LR_LOADFROMFILE = &H10
Public Const IMAGE_BITMAP = 0&
Public Const gdPi   As Double = 3.14159265358979    'Pi
Public Const rads As Double = gdPi / 180


Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public PicInfo As BITMAP
Public DIBInfo As BITMAPINFO
Public light() As RGBA
Public lightMap() As Byte
Public shadowMap() As Byte
Public alpha() As RGBA
Public img() As RGBA
Public texture() As RGBA
Public displace() As RGBA
Public LightColors(0 To 255) As RGBA
Public bUseColor As Boolean
Public bAutoUpdate As Boolean

Public Function BlendColors(color1 As RGBA, color2 As RGBA, ByVal lSteps As Long) As RGBA()
    'Creates an array of colors blending from Color1 to Color2 in lSteps number of steps.
    'Returns the  BlendColors() array.
    Dim lIdx                As Long
    Dim fRedStp             As Single
    Dim fGrnStp             As Single
    Dim fBluStp             As Single
    Dim ReturnColors()    As RGBA

    'Stop possible error
    If lSteps < 2 Then lSteps = 2
    'Find the amount of change for each color element per color change.
    fRedStp = Div(CSng(color2.r) - color1.r, CSng(lSteps))
    fGrnStp = Div(CSng(color2.g) - color1.g, CSng(lSteps))
    fBluStp = Div(CSng(color2.b) - color1.b, CSng(lSteps))
    
    'Create the colors
    ReDim ReturnColors(lSteps - 1) As RGBA
    ReturnColors(0) = color1            'First Color
    ReturnColors(lSteps - 1) = color2   'Last Color
    For lIdx = 1 To lSteps - 2          'All Colors between
        ReturnColors(lIdx).r = CLng(color1.r) + (fRedStp * CSng(lIdx))
        ReturnColors(lIdx).g = CLng(color1.g) + (fGrnStp * CSng(lIdx))
        ReturnColors(lIdx).b = CLng(color1.b) + (fBluStp * CSng(lIdx))
    Next lIdx
    'Return colors in array
    BlendColors = ReturnColors
End Function

Public Function Div(ByVal dNumer As Double, ByVal dDenom As Double) As Double
' Divides 2 numbers avoiding a "Division by zero" error.

    If dDenom <> 0 Then
        Div = dNumer / dDenom
    Else
        Div = 0
    End If
End Function

Public Function LineAngleRadians(Line1 As lineDbl) As Double
'Calculates the angle(in radians) of a line from ptStart to ptEnd.
    Dim dDeltaX As Double
    Dim dDeltaY As Double
    Dim dAngle  As Double

    dDeltaX = Line1.ptEnd.x - Line1.ptStart.x
    dDeltaY = Line1.ptEnd.y - Line1.ptStart.y
    If dDeltaX = 0 Then      'Vertical
        If dDeltaY < 0 Then
            dAngle = gdPi / 2
        Else
            dAngle = gdPi * 1.5
        End If
    ElseIf dDeltaY = 0 Then  'Horizontal
        If dDeltaX >= 0 Then
            dAngle = 0
        Else
            dAngle = gdPi
        End If
    Else    'Angled
        'Note: ++ = positive X, positive Y; +- = positive X, negative Y; etc.
        'On a true coordinate plane, Y increases as it move upward.
        'In VB coordinates, Y is reversed. It increases as it moves downward.
        'Calc for true Upper Right Quadrant (++) (For VB this is +-)
        dAngle = Atn(Abs(dDeltaY / dDeltaX))        'VB Upper Right (+-)
        'Correct for other 3 quadrants in VB coordinates (Reversed Y)
        If dDeltaX >= 0 And dDeltaY >= 0 Then       'VB Lower Right (++)
            dAngle = (gdPi * 2) - dAngle
        ElseIf dDeltaX < 0 And dDeltaY >= 0 Then    'VB Lower Left (-+)
            dAngle = gdPi + dAngle
        ElseIf dDeltaX < 0 And dDeltaY < 0 Then     'VB Upper Left (--)
            dAngle = gdPi - dAngle
        End If
    End If
    LineAngleRadians = dAngle
End Function

Public Function RadiansToDegrees(ByVal dRadians As Double) As Double
'Converts Radians to Degrees.

    RadiansToDegrees = dRadians * (180# / gdPi)
End Function

Public Function LineAngleDegrees(Line1 As lineDbl) As Double
'Returns the angle of a line in degrees (see LineAngleRadians).

    LineAngleDegrees = RadiansToDegrees(LineAngleRadians(Line1))
End Function

Public Function Distance(ptStart As ptDouble, ptEnd As ptDouble) As Double
'Calculates the distance between 2 points.

    'Standard hypotenuse equation (c = Sqr(a^2 + b^2))
    Distance = Sqr(((ptEnd.x - ptStart.x) ^ 2) + ((ptEnd.y - ptStart.y) ^ 2))
End Function

Public Function PointOnLine(ptStart As ptDouble, ptEnd As ptDouble, ByVal dDistance As Double) As ptDouble
'Returns a point on a line at dDistance from ptStart.
'This point need not be between ptStart and ptEnd.
    Dim dDX     As Single
    Dim dDY     As Single
    Dim dLen    As Single
    Dim dPct    As Single
    
    If dDistance > 1000000 Then
        dDistance = 1000000
    End If
        
    dLen = Distance(ptStart, ptEnd)
    
    If dLen > 0 Then
        dDX = ptEnd.x - ptStart.x
        dDY = ptEnd.y - ptStart.y
        dPct = Div(dDistance, dLen)
        PointOnLine.x = ptStart.x + (dDX * dPct)
        PointOnLine.y = ptStart.y + (dDY * dPct)
    Else
        PointOnLine.x = ptStart.x
        PointOnLine.y = ptStart.y
    End If
    
End Function

Public Sub SetColorMap(pb As PictureBox)
    Dim xx As Integer
    Dim yy As Integer
    Dim tmp()  As RGBA
    Dim w As Integer
    Dim h As Integer
    
    w = UBound(lightMap, 1)
    h = UBound(lightMap, 2)
    ReDim tmp(w, h) As RGBA
    For xx = 0 To w
        For yy = 0 To h
            tmp(xx, yy) = LightColors(255 - lightMap(w - xx, h - yy))
        Next yy
    Next xx
    SetPic pb, tmp
    pb.Refresh
    ReDim tmp(0, 0)
    Erase tmp
End Sub

Public Sub ResetColorLight(pb As PictureBox)
    Dim xx As Integer
    Dim yy As Integer
    Dim tmp()  As RGBA
    Dim w As Integer
    Dim h As Integer
    
    w = UBound(lightMap, 1)
    h = UBound(lightMap, 2)
    ReDim tmp(w, h) As RGBA
    For xx = 0 To w
        For yy = 0 To h
            tmp(xx, yy) = LightColors(255 - lightMap(w - xx, h - yy))
        Next yy
    Next xx
    SetPic pb, tmp
    pb.Refresh
    ReDim tmp(0, 0)
    Erase tmp
End Sub

' *========================================================================*
' Thanks goes to Mike D Sutton of EDIAS Software for his help in
' understanding some of the mechanics required for this routine,
' even though my interpretation probably takes too many liberties away
' from his sound guidance.
' Also thanks to the many demo bump mapping source codes out their in
' WWW land which also helped me get this to where it is so far :-)
' *========================================================================*
Public Sub ApplyLightMap(ByVal LtX As Integer, ByVal LtY As Integer)
    Dim x As Integer, y As Integer
    Dim r As Long, g As Long, b As Long
    Dim nx As Double, ny As Double
    Dim LX As Long, LY As Long
    Dim w As Integer, h As Integer
    Dim xx As Integer, yy As Integer
    
    w = UBound(img, 1)
    h = UBound(img, 2)
    
    For x = LBound(img, 1) + 1 To UBound(img, 1) - 1
        For y = LBound(img, 2) + 1 To UBound(img, 2) - 1
            ' calculate some lighting normals ...
            nx = CInt(alpha(x + 1, y).r) - alpha(x - 1, y).r
            ny = CInt(alpha(x, y + 1).r) - alpha(x, y - 1).r
            ' get distance from light to point ...
            LX = x - LtX
            LY = y - LtY
            ' adjust the normals and level out the contrast ...
            nx = nx - LX
            ny = ny - LY
            nx = nx + 128
            ny = ny + 128
            ' catch any out of bounds problems ...
            If nx <= 0 Then nx = 1
            If nx >= UBound(img, 1) Then nx = UBound(img, 1) - 1 '0
            If ny <= 0 Then ny = 1
            If ny >= UBound(img, 2) Then ny = UBound(img, 2) - 1 ' 0
            ' and map result onto picture ...
            img(x, y).r = lightMap(nx, ny)
            img(x, y).g = lightMap(nx, ny)
            img(x, y).b = lightMap(nx, ny)
        Next y
    Next x
End Sub

Public Sub BuildLightMap(ByVal w As Integer, ByVal h As Integer)
    Dim xx As Integer
    Dim yy As Integer
    Dim n As Single
    Dim l As Long
    Dim pic() As RGBA
    Dim ratio As Double
    Dim count As Double
    
    ReDim lightMap(w, h) As Byte
    For xx = 1 To w - 1 'LBound(lightMap, 1) To UBound(lightMap, 1) ' - 1
        For yy = 1 To h - 1 'LBound(lightMap, 2) To UBound(lightMap, 2) ' - 1
            ' we invert the mapping cos for some strange reason it
            ' ends up opposite other wise ...
            lightMap(xx, yy) = light(w - xx, h - yy).r
        Next yy
    Next xx
End Sub

Public Function LoadPic(pic As StdPicture) As RGBA()
  Dim hdcNew As Long
  Dim oldhand As Long
  Dim BytesPerScanLine As Long
  Dim PadBytesPerScanLine As Long
  Dim ret As Long
  Dim b() As RGBA
    
  hdcNew = CreateCompatibleDC(0&)
  oldhand = SelectObject(hdcNew, pic.Handle)
  Call GetObject(pic.Handle, Len(PicInfo), PicInfo)
  With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = PicInfo.bmWidth
    .biHeight = -PicInfo.bmHeight     'bottom up scan line is now inverted
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    PadBytesPerScanLine = _
       BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With
  'redimension  (BGR+pad,x,y)
  ReDim b(1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As RGBA
  'get bytes
  ret = GetDIBits(hdcNew, pic.Handle, 0, PicInfo.bmHeight, b(1, 1), DIBInfo, DIB_RGB_COLORS)
  SelectObject hdcNew, oldhand
  DeleteDC hdcNew
  
  LoadPic = b
End Function

Public Sub SetPic(pb As Picture, bits() As RGBA)
    Dim BytesPerScanLine As Long
    Dim PadBytesPerScanLine As Long
    
  Dim hdcNew As Long
  Dim oldhand As Long

  hdcNew = CreateCompatibleDC(0&)
  oldhand = SelectObject(hdcNew, pb.Handle)
  
  Call GetObject(pb.Handle, Len(PicInfo), PicInfo)
  With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = PicInfo.bmWidth
    .biHeight = -PicInfo.bmHeight
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    PadBytesPerScanLine = _
       BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With
  SetDIBits hdcNew, pb.Handle, 0, PicInfo.bmHeight, bits(1, 1), DIBInfo, DIB_RGB_COLORS

  SelectObject hdcNew, oldhand
  DeleteDC hdcNew
End Sub

' *========================================================================*
' Thanks goes to Mike D Sutton of EDIAS Software for his help in
' understanding some of the mechanics required for this routine,
' even though my interpretation probably takes too many liberties away
' from his sound guidance :-)
' *========================================================================*
Public Function AlphaDisplace(pic As StdPicture, alpha As StdPicture, Optional ByVal maxOffset As Integer = 15) As RGBA() 'StdPicture
    Dim x As Integer, y As Integer, ofsx As Integer, ofsy As Integer
    Dim xdisp As Integer, ydisp As Integer, w As Integer, h As Integer
    Dim bits() As RGBA, abits() As RGBA, retbits() As RGBA
    Dim scaleFactor As Double
 
    scaleFactor = maxOffset / 255
    bits = LoadPic(pic)
    abits = LoadPic(alpha)
    w = UBound(abits, 1)
    h = UBound(abits, 2)
    ReDim retbits(LBound(bits, 1) To UBound(bits, 1), LBound(bits, 2) To UBound(bits, 2)) As RGBA
    
'On Error Resume Next
    For y = LBound(abits, 2) To UBound(abits, 2)
        For x = LBound(abits, 1) To UBound(abits, 1)
            xdisp = (abits(x, y).r * scaleFactor) + x
            ydisp = (abits(x, y).r * scaleFactor) + y
            If xdisp > w Then xdisp = w - xdisp
            If xdisp < 0 Then xdisp = 1 'w + xdisp '0
            If ydisp > h Then ydisp = h - ydisp
            If ydisp < 0 Then ydisp = 1 'h + ydisp '0
        '    bits(X, Y).r = abits(xdisp, ydisp).r
        '    bits(X, Y).g = abits(xdisp, ydisp).g
        '    bits(X, Y).b = abits(xdisp, ydisp).b
            retbits(x, y).r = bits(xdisp, ydisp).r
            retbits(x, y).g = bits(xdisp, ydisp).g
            retbits(x, y).b = bits(xdisp, ydisp).b
        Next x
    Next y
'On Error GoTo 0
    
    ReDim bits(0, 0)
    ReDim abits(0, 0)
    Erase bits
    Erase abits
    AlphaDisplace = retbits
End Function

' *========================================================================*
' Thanks goes to Steve Mc Mahon of VBAccelerator for this excellent helper
' routine :-)
' *========================================================================*
Public Sub SetHLS(ByVal h As Integer, ByVal l As Integer, ByVal s As Integer, r As Long, g As Long, b As Long)
    Dim MyR     As Single, MyG    As Single, MyB    As Single
    Dim MyH     As Single, MyL    As Single, MyS    As Single
    Dim Min     As Single, Max    As Single, Delta  As Single
    
    MyH = (h / 60) - 1: MyL = l / 100: MyS = s / 100
    If MyS = 0 Then
        MyR = MyL: MyG = MyL: MyB = MyL
    Else
        If MyL <= 0.5 Then
            Min = MyL * (1 - MyS)
        Else
            Min = MyL - MyS * (1 - MyL)
        End If
        Max = 2 * MyL - Min
        Delta = Max - Min
        
        Select Case MyH
        Case Is < 1
            MyR = Max
            If MyH < 0 Then
                MyG = Min
                MyB = MyG - MyH * Delta
            Else
                MyB = Min
                MyG = MyH * Delta + MyB
            End If
        Case Is < 3
            MyG = Max
            If MyH < 2 Then
                MyB = Min
                MyR = MyB - (MyH - 2) * Delta
            Else
                MyR = Min
                MyB = (MyH - 2) * Delta + MyR
            End If
        Case Else
            MyB = Max
            If MyH < 4 Then
                MyR = Min
                MyG = MyR - (MyH - 4) * Delta
            Else
                MyG = Min
                MyR = (MyH - 4) * Delta + MyG
            End If
        End Select
    End If
    
    r = MyR * 255: g = MyG * 255: b = MyB * 255
End Sub

' *========================================================================*
' Thanks goes to Steve Mc Mahon of VBAccelerator for this excellent helper
' routine :-)
' *========================================================================*
Public Sub GetRGB(ByVal col As Long, ByRef r As Long, ByRef g As Long, ByRef b As Long)
  r = col Mod 256
  g = ((col And &HFF00&) \ 256&) Mod 256&
  b = (col And &HFF0000) \ 65536
End Sub

' *========================================================================*
' Thanks goes to Steve Mc Mahon of VBAccelerator for this excellent helper
' routine :-)
' *========================================================================*
Private Function Maximum(rr As Single, rG As Single, rB As Single) As Single
   If (rr > rG) Then
      If (rr > rB) Then
         Maximum = rr
      Else
         Maximum = rB
      End If
   Else
      If (rB > rG) Then
         Maximum = rB
      Else
         Maximum = rG
      End If
   End If
End Function

' *========================================================================*
' Thanks goes to Steve Mc Mahon of VBAccelerator for this excellent helper
' routine :-)
' *========================================================================*
Private Function Minimum(rr As Single, rG As Single, rB As Single) As Single
   If (rr < rG) Then
      If (rr < rB) Then
         Minimum = rr
      Else
         Minimum = rB
      End If
   Else
      If (rB < rG) Then
         Minimum = rB
      Else
         Minimum = rG
      End If
   End If
End Function

' *========================================================================*
' Thanks goes to Steve Mc Mahon of VBAccelerator for this excellent helper
' routine :-)
' *========================================================================*
Public Sub RGBToHLS( _
      ByVal r As Long, ByVal g As Long, ByVal b As Long, _
      h As Single, s As Single, l As Single _
   )
Dim Max As Single
Dim Min As Single
Dim Delta As Single
Dim rr As Single, rG As Single, rB As Single

   rr = r / 255: rG = g / 255: rB = b / 255

'{Given: rgb each in [0,1].
' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
        Max = Maximum(rr, rG, rB)
        Min = Minimum(rr, rG, rB)
        l = (Max + Min) / 2    '{This is the lightness}
        '{Next calculate saturation}
        If Max = Min Then
            'begin {Acrhomatic case}
            s = 0
            h = 0
           'end {Acrhomatic case}
        Else
           'begin {Chromatic case}
                '{First calculate the saturation.}
           If l <= 0.5 Then
               s = (Max - Min) / (Max + Min)
           Else
               s = (Max - Min) / (2 - Max - Min)
            End If
            '{Next calculate the hue.}
            Delta = Max - Min
           If rr = Max Then
                h = (rG - rB) / Delta    '{Resulting color is between yellow and magenta}
           ElseIf rG = Max Then
                h = 2 + (rB - rr) / Delta '{Resulting color is between cyan and yellow}
           ElseIf rB = Max Then
                h = 4 + (rr - rG) / Delta '{Resulting color is between magenta and cyan}
            End If
            'h = h * 60
           'If h < 0# Then
           '     h = h + 360            '{Make degrees be nonnegative}
           'End If
        'end {Chromatic Case}
      End If
'end {RGB_to_HLS}
End Sub

' ===================================================================================
' Function : PtInEllipse
' ===================================================================================
' returns True if pt is within the ellipse bounded by rcBounds ...
' ===================================================================================
Public Function PtInEllipse(pt As ptDouble, rcBounds As Rect) As Boolean
    Dim x As Double
    Dim y As Double
    Dim a As Double
    Dim b As Double
    
'   Determine radii ...
    a = (rcBounds.right - rcBounds.left) / 2
    b = (rcBounds.bottom - rcBounds.top) / 2
'   Determine x, y ...
    x = pt.x - (rcBounds.left + rcBounds.right) / 2
    y = pt.y - (rcBounds.top + rcBounds.bottom) / 2
'   Apply ellipse formula ...
    PtInEllipse = ((x * x) / (a * a) + (y * y) / (b * b) <= 1)
End Function


