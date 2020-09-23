VERSION 5.00
Begin VB.Form fMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "About..."
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3780
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lighting Options..."
      Height          =   360
      Left            =   195
      TabIndex        =   0
      Top             =   3810
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"vbeam.frx":0000
      Height          =   585
      Left            =   345
      TabIndex        =   2
      Top             =   2625
      Width           =   4485
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   0
      Picture         =   "vbeam.frx":0094
      Top             =   0
      Width           =   2250
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    fLightingProps.Show
End Sub

Private Sub Command2_Click()
    fAbout.Show 1
End Sub
