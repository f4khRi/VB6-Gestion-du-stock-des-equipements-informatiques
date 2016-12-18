VERSION 5.00
Begin VB.Form Rimage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   0
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "Rimage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
' Zoom image
Unload Me
End Sub

