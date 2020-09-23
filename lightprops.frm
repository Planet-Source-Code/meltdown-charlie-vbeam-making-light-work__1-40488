VERSION 5.00
Begin VB.Form fLightingProps 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lighting Properties"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8130
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   382
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7740
      Left            =   9045
      Picture         =   "lightprops.frx":0000
      ScaleHeight     =   516
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   516
      TabIndex        =   22
      Top             =   5685
      Visible         =   0   'False
      Width           =   7740
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   255
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   20
      Top             =   3240
      Width           =   3750
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         DrawMode        =   6  'Mask Pen Not
         X1              =   -7
         X2              =   366
         Y1              =   48
         Y2              =   48
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         DrawMode        =   6  'Mask Pen Not
         X1              =   142
         X2              =   142
         Y1              =   -2
         Y2              =   145
      End
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2310
      Left            =   330
      Picture         =   "lightprops.frx":C0042
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   19
      Top             =   6585
      Visible         =   0   'False
      Width           =   3810
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2310
      Left            =   4200
      Picture         =   "lightprops.frx":DB924
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   18
      Top             =   6615
      Visible         =   0   'False
      Width           =   3810
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2310
      Left            =   4155
      Picture         =   "lightprops.frx":F7206
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   17
      Top             =   3240
      Width           =   3810
   End
   Begin VB.HScrollBar scrollReflect 
      Height          =   210
      Left            =   1500
      Max             =   255
      TabIndex        =   15
      Top             =   2670
      Value           =   35
      Width           =   1545
   End
   Begin VB.HScrollBar scrollBlue 
      Height          =   210
      Left            =   5610
      Max             =   255
      TabIndex        =   11
      Top             =   2070
      Value           =   238
      Width           =   1125
   End
   Begin VB.HScrollBar scrollGreen 
      Height          =   210
      Left            =   3495
      Max             =   255
      TabIndex        =   10
      Top             =   2070
      Value           =   235
      Width           =   1155
   End
   Begin VB.HScrollBar scrollRed 
      Height          =   210
      Left            =   1305
      Max             =   255
      TabIndex        =   6
      Top             =   2070
      Value           =   235
      Width           =   1170
   End
   Begin VB.PictureBox shpColour 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   5940
      ScaleHeight     =   870
      ScaleWidth      =   810
      TabIndex        =   1
      Top             =   990
      Width           =   840
   End
   Begin VB.PictureBox pbColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   375
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   367
      TabIndex        =   0
      Top             =   990
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      Height          =   240
      Index           =   3
      Left            =   1485
      Top             =   2655
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   240
      Index           =   2
      Left            =   5595
      Top             =   2055
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      Height          =   240
      Index           =   1
      Left            =   3480
      Top             =   2055
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      Height          =   240
      Index           =   0
      Left            =   1290
      Top             =   2055
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Light Position"
      Height          =   195
      Index           =   1
      Left            =   285
      TabIndex        =   21
      Top             =   2985
      Width           =   945
   End
   Begin VB.Image imgAmbient 
      Height          =   315
      Left            =   3645
      Top             =   450
      Width           =   1620
   End
   Begin VB.Image imgDiffuse 
      Height          =   315
      Left            =   1995
      Top             =   450
      Width           =   1590
   End
   Begin VB.Image imgSpecular 
      Height          =   345
      Left            =   285
      Top             =   450
      Width           =   1605
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "35"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3150
      TabIndex        =   16
      Top             =   2655
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Specular Value"
      Height          =   195
      Left            =   330
      TabIndex        =   14
      Top             =   2670
      Width           =   1080
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      Height          =   180
      Index           =   2
      Left            =   4770
      TabIndex        =   13
      Top             =   2100
      Width           =   360
   End
   Begin VB.Label labBlue 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "235"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5160
      TabIndex        =   12
      Top             =   2055
      Width           =   405
   End
   Begin VB.Label labGreen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "235"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3030
      TabIndex        =   9
      Top             =   2055
      Width           =   405
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
      Height          =   180
      Index           =   1
      Left            =   2565
      TabIndex        =   8
      Top             =   2100
      Width           =   450
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
      Height          =   180
      Index           =   0
      Left            =   435
      TabIndex        =   7
      Top             =   2100
      Width           =   375
   End
   Begin VB.Label labRed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "235"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   855
      TabIndex        =   5
      Top             =   2055
      Width           =   405
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   16
      X1              =   463
      X2              =   458
      Y1              =   161
      Y2              =   167
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   15
      X1              =   16
      X2              =   20
      Y1              =   159
      Y2              =   167
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   14
      Visible         =   0   'False
      X1              =   21
      X2              =   17
      Y1              =   51
      Y2              =   57
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   13
      X1              =   246
      X2              =   242
      Y1              =   27
      Y2              =   33
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   12
      X1              =   134
      X2              =   130
      Y1              =   27
      Y2              =   33
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   10
      X1              =   20
      X2              =   16
      Y1              =   27
      Y2              =   33
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   8
      X1              =   459
      X2              =   463
      Y1              =   51
      Y2              =   56
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   7
      X1              =   124
      X2              =   128
      Y1              =   27
      Y2              =   32
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   5
      X1              =   236
      X2              =   240
      Y1              =   27
      Y2              =   32
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   4
      X1              =   348
      X2              =   352
      Y1              =   27
      Y2              =   32
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   241
      X2              =   241
      Y1              =   34
      Y2              =   50
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   129
      X2              =   129
      Y1              =   34
      Y2              =   51
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   23
      X1              =   16
      X2              =   16
      Y1              =   51
      Y2              =   159
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   22
      X1              =   16
      X2              =   16
      Y1              =   35
      Y2              =   50
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   21
      X1              =   123
      X2              =   21
      Y1              =   27
      Y2              =   27
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   20
      X1              =   233
      X2              =   135
      Y1              =   27
      Y2              =   27
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   19
      X1              =   203
      X2              =   174
      Y1              =   51
      Y2              =   51
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   18
      X1              =   457
      X2              =   130
      Y1              =   51
      Y2              =   51
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   17
      X1              =   348
      X2              =   247
      Y1              =   27
      Y2              =   27
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   11
      X1              =   352
      X2              =   352
      Y1              =   32
      Y2              =   51
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AmbientColour"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   495
      Width           =   1500
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   9
      X1              =   241
      X2              =   241
      Y1              =   34
      Y2              =   51
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Diffuse Colour"
      Height          =   255
      Left            =   2025
      TabIndex        =   3
      Top             =   495
      Width           =   1500
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   128
      X2              =   128
      Y1              =   34
      Y2              =   51
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Specular Colour"
      Height          =   255
      Left            =   285
      TabIndex        =   2
      Top             =   510
      Width           =   1590
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   2
      X1              =   456
      X2              =   21
      Y1              =   166
      Y2              =   167
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   1
      X1              =   463
      X2              =   463
      Y1              =   57
      Y2              =   161
   End
End
Attribute VB_Name = "fLightingProps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bDraggingAngle As Boolean
Private iColorToChange As Integer
Private SpecularColor As Long
Private DiffuseColor As Long
Private AmbientColor As Long
Private Reflectance As Integer
Private LightX As Integer
Private LightY As Integer
Private mSetLight As Boolean

Private Sub UpdateLight()
    Dim x As Integer, y As Integer
    Dim c As Byte
    
    ' reload all of our pixel maps ...
    img = LoadPic(Picture1)
    light = LoadPic(Picture3)
    alpha = LoadPic(Picture2)
    ' build the light map for illumination values ...
    BuildLightMap UBound(img, 1), UBound(img, 2)
    ' apply the light to the image using alpha channel for bumpiness ...
    ApplyLightMap LightX, LightY 'UBound(img, 1) \ 2, -(UBound(img, 2) \ 2)
    ' map the colors to our lighting values ...
    For x = LBound(img, 1) To UBound(img, 1)
        For y = LBound(img, 2) To UBound(img, 2)
           c = img(x, y).r
           img(x, y) = LightColors(c)
        Next y
    Next x
    ' and set the results back to the picture box ...
    SetPic Picture1, img
    Picture1.Refresh
    ReDim img(0, 0)
    ReDim alpha(0, 0)
    ReDim light(0, 0)
    Erase img
    Erase light
    Erase alpha
End Sub

Private Sub UpdateColors()
    Dim i As Integer
    Dim r As Long, g As Long, b As Long
    Dim aRed As Long, aGreen As Long, aBlue As Long
    Dim dRed As Long, dGreen As Long, dBlue As Long
    Dim sRed As Long, sGreen As Long, sBlue As Long
    Dim nx As Single, ny2 As Single, ca As Single, cap As Single
    
    nx = gdPi / 2
    ny2 = nx / 256
    GetRGB AmbientColor, aRed, aGreen, aBlue
    GetRGB DiffuseColor, dRed, dGreen, dBlue
    GetRGB SpecularColor, sRed, sGreen, sBlue
    For i = 0 To 255
        ca = Cos(nx)
        cap = ca ^ Reflectance
        nx = nx - ny2
        r = Fix(aRed + (dRed * ca) + (sRed * cap))
        If r > 255 Then r = 255
        LightColors(i).r = r
        g = Fix(aGreen + (dGreen * ca) + (sGreen * cap))
        If g > 255 Then g = 255
        LightColors(i).g = g
        b = Fix(aBlue + (dBlue * ca) + (sBlue * cap))
        If b > 255 Then b = 255
        LightColors(i).b = b
    Next i
End Sub

Private Sub Form_Load()
    Dim lH As Long
    Dim iLoop As Integer
    Dim r As Long
    Dim g As Long
    Dim b As Long
    Dim n As Integer
    Dim incr As Single
    Dim colrs() As RGBA
    Dim c1 As RGBA, c2 As RGBA, c3 As RGBA
    
    Line4.X2 = Picture4.ScaleWidth
    Line3.Y2 = Picture4.ScaleHeight
    
    Reflectance = 35
    iColorToChange = 0
    SpecularColor = RGB(235, 235, 235)
    DiffuseColor = RGB(128, 128, 128)
    AmbientColor = RGB(0, 0, 0)

    c1.r = 255
    c1.g = 255
    c1.b = 255
    c3.r = 0
    c3.g = 0
    c3.b = 0
    ' load the hues for the available light colors
    lH = pbColor.ScaleHeight
     For iLoop = 0 To 360
         SetHLS iLoop, 50, 100, r, g, b
         c2.r = r
         c2.g = g
         c2.b = b
         colrs = BlendColors(c1, c2, 31)
         For n = 0 To 30
             pbColor.PSet (iLoop, n), RGB(colrs(n).r, colrs(n).g, colrs(n).b)
         Next n
         colrs = BlendColors(c2, c3, 31)
         For n = 0 To 30
             pbColor.PSet (iLoop, 30 + n), RGB(colrs(n).r, colrs(n).g, colrs(n).b)
         Next n
     Next iLoop
     For n = 0 To 60
         colrs = BlendColors(c1, c3, 61)
         pbColor.Line (iLoop, n)-(pbColor.ScaleWidth, n), RGB(colrs(n).r, colrs(n).g, colrs(n).b)
     Next n
    pbColor.Refresh
    UpdateColors
    LightX = Picture1.ScaleWidth \ 2
    LightY = Picture1.ScaleHeight \ 2
    Picture4_MouseUp 0, 0, CSng(LightX), CSng(LightY)
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mSetLight = True
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mSetLight Then
        Line3.X1 = x
        Line3.X2 = x
        Line4.Y1 = y
        Line4.Y2 = y
    End If
End Sub

Private Sub Picture4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ww As Integer
    Dim hh As Integer
    Dim cx As Integer
    Dim cy As Integer
    
    mSetLight = False
    With Picture3
        ww = .ScaleWidth - 1
        hh = .ScaleHeight - 1
        cx = Line3.X1
        If cx > .ScaleWidth Then cx = .ScaleWidth - 1
        If cx <= 0 Then cx = 1
        cy = Line4.Y1
        If cy > .ScaleHeight Then cy = .ScaleHeight - 1
        If cy <= 0 Then cy = 1
        .PaintPicture Picture5.Picture, 0, 0, cx, cy, 0, 0, 256, 256
        .PaintPicture Picture5.Picture, 0, cy, cx, .ScaleHeight - cy, 0, 256, 256, 256
        .PaintPicture Picture5.Picture, cx, 0, .ScaleWidth - cx, cy, 256, 0, 256, 256
        .PaintPicture Picture5.Picture, cx, cy, .ScaleWidth - cx, .ScaleHeight - cy, 256, 256, 256, 256
        .Picture = .Image
        .Refresh
        light = LoadPic(.Picture)
    End With
''    If cx > Picture3.ScaleWidth \ 2 Then
''        LightX = x - ((Picture3.ScaleWidth - cx) \ 2)
''    Else
''        LightX = x + ((Picture3.ScaleWidth - cx) \ 2)
''    End If
''    If cy < Picture3.ScaleHeight \ 2 Then
''        LightY = y - ((Picture3.ScaleHeight - cy) \ 2) '- (Picture1.ScaleHeight \ 2)
''    Else
''        LightY = y + ((Picture3.ScaleHeight - cy) \ 2) '- (Picture1.ScaleHeight \ 2)
''    End If
    LightX = x
    LightY = y - (Picture4.ScaleHeight \ 2) '-y
    UpdateLight
End Sub

Private Sub scrollReflect_Change()
    Reflectance = scrollReflect.Value
    Label7 = Reflectance
    UpdateColors
    UpdateLight
End Sub

Private Sub imgSpecular_Click()
    Dim r As Long, g As Long, b As Long
    
    Line2(18).X1 = Line2(8).X1
    Line2(18).X2 = Line2(0).X1
    Line2(19).X2 = Line2(0).X1
    Line2(19).X1 = Line2(0).X1
    Line2(14).Visible = False
    Refresh
    iColorToChange = 0
    shpColour.BackColor = SpecularColor
    GetRGB SpecularColor, r, g, b
    labRed = r
    labGreen = g
    labBlue = b
    scrollRed.Value = r
    scrollGreen.Value = g
    scrollBlue.Value = b
End Sub

Private Sub imgDiffuse_Click()
    Dim r As Long, g As Long, b As Long
    
    Line2(19).X1 = Line2(14).X1
    Line2(19).X2 = Line2(6).X1
    Line2(18).X1 = Line2(3).X1
    Line2(18).X2 = Line2(8).X1
    Line2(14).Visible = True
    Refresh
    iColorToChange = 1
    shpColour.BackColor = DiffuseColor
    GetRGB DiffuseColor, r, g, b
    labRed = r
    labGreen = g
    labBlue = b
    scrollRed.Value = r
    scrollGreen.Value = g
    scrollBlue.Value = b
End Sub

Private Sub imgAmbient_Click()
    Dim r As Long, g As Long, b As Long
    
    Line2(18).X2 = Line2(8).X1
    Line2(18).X1 = Line2(11).X1
    Line2(19).X1 = Line2(3).X1
    Line2(19).X2 = Line2(14).X1
    Line2(14).Visible = True
    Refresh
    iColorToChange = 2
    shpColour.BackColor = AmbientColor
    GetRGB AmbientColor, r, g, b
    labRed = r
    labGreen = g
    labBlue = b
    scrollRed.Value = r
    scrollGreen.Value = g
    scrollBlue.Value = b
End Sub

Private Sub pbBrightness_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim c As Long
    
    If Button <> 0 Then
        c = pbBrightness.Point(x, y)
        If c <> -1 Then
            shpBrightness.BackColor = c
        End If
        lnBrightness.X1 = x
        lnBrightness.X2 = x
    End If
End Sub

Private Sub pbColor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim r As Long, g As Long, b As Long, c As Long
    
    c = pbColor.Point(x, y)
    shpColour.BackColor = c
    Select Case iColorToChange
        Case 0 ' Specular Colour
            SpecularColor = c
        Case 1 ' Diffuse Colour
            DiffuseColor = c
        Case 2 ' Ambient Colour
            AmbientColor = c
    End Select
    
    GetRGB c, r, g, b
    labRed = r
    labGreen = g
    labBlue = b
    scrollRed.Value = r
    scrollGreen.Value = g
    scrollBlue.Value = b
    UpdateColors
    UpdateLight
End Sub

Private Sub pbColor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim r As Long, g As Long, b As Long, c As Long
    
    If Button And 1 Then
        If x > 0 And y > 0 And x < pbColor.ScaleWidth And y < pbColor.ScaleHeight Then
            c = pbColor.Point(x, y)
            shpColour.BackColor = c
            Select Case iColorToChange
                Case 0 ' Specular Colour
                    SpecularColor = c
                Case 1 ' Diffuse Colour
                    DiffuseColor = c
                Case 2 ' Ambient Colour
                    AmbientColor = c
            End Select
            GetRGB c, r, g, b
            labRed = r
            labGreen = g
            labBlue = b
            scrollRed.Value = r
            scrollGreen.Value = g
            scrollBlue.Value = b
            UpdateColors
            UpdateLight
        End If
    End If
End Sub


Private Sub pbDirection_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim pt As ptDouble
    Dim rc As Rect
    
    If Button And 1 Then
        pt.x = x
        pt.y = y
        rc.left = Shape2.left
        rc.top = Shape2.top
        rc.right = Shape2.left + Shape2.Width
        rc.bottom = Shape2.top + Shape2.Height
        If PtInEllipse(pt, rc) Then
            Shape3.FillColor = RGB(64, 64, 64)
            bDraggingAngle = True
        End If
    End If
End Sub

Private Sub pbDirection_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim pt1 As ptDouble, pt2 As ptDouble, ptMouse As ptDouble
    
    If bDraggingAngle Then
        ptMouse.x = (x \ Screen.TwipsPerPixelX) + pbDirection.left
        ptMouse.y = (y \ Screen.TwipsPerPixelY) + pbDirection.top
        ' are we changing the angle of the gradient ? ...
        Dim ln As lineDbl
        Dim pt As ptDouble
        Dim degrees As Double
        
        ' make a line from center of circle to current x,y position ...
        ln.ptStart.x = Line1.X1
        ln.ptStart.y = Line1.Y1
        ln.ptEnd.x = x
        ln.ptEnd.y = y
        ' find out where this line would intersect the circles perimeter ...
        pt = PointOnLine(ln.ptStart, ln.ptEnd, Shape1.Height \ 2)
        ' get the angle of the line ...
        degrees = LineAngleDegrees(ln)
        ' shift the line and the red circle to the new position ...
        Line1.X2 = pt.x
        Line1.Y2 = pt.y
        Shape2.left = pt.x - (Shape2.Width \ 2)
        Shape2.top = pt.y - (Shape2.Height \ 2)
        Exit Sub
    End If
    
    pt1.x = Line1.X1
    pt1.y = Line1.Y1
    pt2.x = Line1.X2
    pt2.y = Line1.Y2
End Sub

Private Sub pbDirection_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Shape3.FillColor = BackColor
    bDraggingAngle = False
End Sub

Private Sub scrollBlue_Change()
    Dim r As Long, g As Long, b As Long

    GetRGB shpColour.BackColor, r, g, b
    r = Abs(r)
    g = Abs(g)
    labBlue = scrollBlue.Value
    shpColour.BackColor = RGB(r, g, CByte(labBlue))
    Select Case iColorToChange
        Case 0 ' Specular Colour
            SpecularColor = shpColour.BackColor
        Case 1 ' Diffuse Colour
            DiffuseColor = shpColour.BackColor
        Case 2 ' Ambient Colour
            AmbientColor = shpColour.BackColor
    End Select
    UpdateColors
    UpdateLight
End Sub

Private Sub scrollBlue_Scroll()
    scrollBlue_Change
End Sub

Private Sub scrollGreen_Change()
    Dim r As Long, g As Long, b As Long

    GetRGB shpColour.BackColor, r, g, b
    r = Abs(r)
    b = Abs(b)
    labGreen = scrollGreen.Value
    shpColour.BackColor = RGB(r, CByte(labGreen), b)
    Select Case iColorToChange
        Case 0 ' Specular Colour
            SpecularColor = shpColour.BackColor
        Case 1 ' Diffuse Colour
            DiffuseColor = shpColour.BackColor
        Case 2 ' Ambient Colour
            AmbientColor = shpColour.BackColor
    End Select
    UpdateColors
    UpdateLight
End Sub

Private Sub scrollGreen_Scroll()
    scrollGreen_Change
End Sub

Private Sub scrollRed_Change()
    Dim r As Long, g As Long, b As Long
    
    GetRGB shpColour.BackColor, r, g, b
    g = Abs(g)
    b = Abs(b)
    labRed = scrollRed.Value
    shpColour.BackColor = RGB(CByte(labRed), g, b) 'c
    Select Case iColorToChange
        Case 0 ' Specular Colour
            SpecularColor = shpColour.BackColor
        Case 1 ' Diffuse Colour
            DiffuseColor = shpColour.BackColor
        Case 2 ' Ambient Colour
            AmbientColor = shpColour.BackColor
    End Select
    UpdateColors
    UpdateLight
End Sub

Private Sub scrollRed_Scroll()
    scrollRed_Change
End Sub

Private Sub scrollReflect_Scroll()
    scrollReflect_Change
End Sub

