VERSION 5.00
Begin VB.Form frmsplash 
   BorderStyle     =   0  'None
   Caption         =   "SPLASH"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9570
   FillColor       =   &H00000080&
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmsplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   6015
      Left            =   0
      Picture         =   "frmsplash.frx":0CCA
      ScaleHeight     =   5955
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   0
      Width           =   9615
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   9120
      Top             =   0
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   0
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   2160
      Top             =   6000
      Width           =   4575
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
If Shape2.Width >= 3375 Then
FRMMENU.Show
Unload Me
Exit Sub
Else
Shape2.Width = Shape2.Width + 90
End If
End Sub
