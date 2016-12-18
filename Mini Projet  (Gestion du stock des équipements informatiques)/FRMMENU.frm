VERSION 5.00
Begin VB.Form FRMMENU 
   BorderStyle     =   0  'None
   Caption         =   ".::  MENgU GENERAL  ::."
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11775
   Icon            =   "FRMMENU.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FRMMENU.frx":0442
   ScaleHeight     =   6345
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   1320
      Picture         =   "FRMMENU.frx":107CB
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "FOURNISSEURS"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   8640
      Picture         =   "FRMMENU.frx":11495
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "QUITTER L'APPLICATION"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   5040
      Picture         =   "FRMMENU.frx":1215F
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "LISTE DES EQUIPEMENTS  EN STOCK"
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0E0FF&
      Default         =   -1  'True
      Height          =   615
      Left            =   8640
      Picture         =   "FRMMENU.frx":12A29
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   1440
      Picture         =   "FRMMENU.frx":136F3
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "EQUIPEMENTS"
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton cmdReload 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   5040
      Picture         =   "FRMMENU.frx":143BD
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "SAUVEGARDER LA BASE DE DONNEE"
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "QUITTER"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   270
      Left            =   8400
      TabIndex        =   11
      Top             =   5640
      Width           =   1050
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "A PROPOS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   270
      Left            =   8400
      TabIndex        =   10
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "STOCK"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   270
      Left            =   4800
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "FOURNISSEURS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   270
      Left            =   840
      TabIndex        =   8
      Top             =   5640
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "SAUVEGARDE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   270
      Left            =   4560
      TabIndex        =   7
      Top             =   5640
      Width           =   1500
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "EQUIPEMENTS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   270
      Left            =   960
      TabIndex        =   6
      Top             =   1080
      Width           =   1650
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   855
      Index           =   0
      Left            =   8520
      Top             =   4560
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   855
      Index           =   1
      Left            =   1200
      Top             =   4560
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   855
      Index           =   2
      Left            =   4920
      Top             =   4560
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   855
      Index           =   3
      Left            =   1320
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   855
      Index           =   4
      Left            =   4920
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   855
      Index           =   5
      Left            =   8520
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "FRMMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub cmdAdd_Click()
' About
On Error Resume Next
Call ShellExecute(hwnd, "Open", "aide.docx", "", App.Path, 1)
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1(5).BackColor = &HFFFF&
End Sub

Private Sub cmdDelete_Click()
frmfournisseur.Show
Unload Me
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1(1).BackColor = &HFFFF&
End Sub

Private Sub cmdEdit_Click()
FRMLISTEEQUIPEMENTS.Show
Unload Me
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1(4).BackColor = &HFFFF&
End Sub
Private Sub cmdExit_Click()
' Quitter
End
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1(0).BackColor = &HFFFF&
End Sub

Private Sub cmdReload_Click()
' backup
frmbackup.Show
Unload Me
End Sub

Private Sub cmdReload_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1(2).BackColor = &HFFFF&
End Sub

Private Sub cmdSort_Click()
GESTEQ.Show
Unload Me
End Sub

Private Sub cmdSort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1(3).BackColor = &HFFFF&
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' mouse-move la couleur jaune
Shape1(5).BackColor = &H0&
Shape1(0).BackColor = &H0&
Shape1(1).BackColor = &H0&
Shape1(2).BackColor = &H0&
Shape1(3).BackColor = &H0&
Shape1(4).BackColor = &H0&
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Label2.Caption = KeyAscii
End Sub





