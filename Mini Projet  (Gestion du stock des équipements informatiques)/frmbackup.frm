VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmbackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAUVEGARDE"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8460
   Icon            =   "frmbackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmbackup.frx":08CA
   ScaleHeight     =   6105
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SAUVEGARDE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5880
      Picture         =   "frmbackup.frx":14D02
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog com1 
      Left            =   960
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim FileName As String
On Error GoTo error
com1.DialogTitle = "Sauvgarde de la base de donnée"
com1.DefaultExt = ".MDB"
com1.FileName = "Stock"
com1.Flags = &H1004
com1.Filter = "Access (.mdb)|*.mdb"
com1.ShowSave
FileName = com1.FileName
FileCopy App.Path & "\base\Stock.mdb", FileName
MsgBox "la base de données a été bien sauvegardée", vbInformation + vbOKOnly, "Gestion Stock"
Exit Sub
error:
MsgBox "Veuillez redémarrer l'application" & Chr$(13) & " et sauvegarder en premier temp", vbCritical + vbOKOnly, "Erreur"
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HFFFF&
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FRMMENU.Show
End Sub
