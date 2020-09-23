VERSION 5.00
Begin VB.Form fAbout 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About VBeam"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      Height          =   360
      Left            =   1635
      TabIndex        =   1
      Top             =   2940
      Width           =   1350
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "scythe@cablenet.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   435
      TabIndex        =   4
      Top             =   2355
      Width           =   3825
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HTTP://WWW25.BRINKSTER.COM/MFerris"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   255
      TabIndex        =   3
      Top             =   2685
      Width           =   4245
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by M Ferris and Scythe for the VB community. You can visit my personal site at "
      Height          =   390
      Left            =   240
      TabIndex        =   2
      Top             =   1935
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VBeam version 1.2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   1545
      Width           =   3675
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   1200
      Picture         =   "fAbout.frx":0000
      Top             =   -45
      Width           =   2250
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   -105
      Top             =   -15
      Width           =   4935
   End
End
Attribute VB_Name = "fAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

