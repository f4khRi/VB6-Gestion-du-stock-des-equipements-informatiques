VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FRMLISTEEQUIPEMENTS 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTE DES EQUIPEMENTS"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11025
   Icon            =   "FRMLISTEEQUIPEMENTS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbtype 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      ItemData        =   "FRMLISTEEQUIPEMENTS.frx":0CCA
      Left            =   6120
      List            =   "FRMLISTEEQUIPEMENTS.frx":0CF2
      TabIndex        =   4
      Top             =   120
      Width           =   2895
   End
   Begin VB.ComboBox cmbfourni 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      ItemData        =   "FRMLISTEEQUIPEMENTS.frx":0DC8
      Left            =   6120
      List            =   "FRMLISTEEQUIPEMENTS.frx":0DCA
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton CMDIMPRIMER 
      BackColor       =   &H0000FFFF&
      Caption         =   "ACTUALISER"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "IMPRIMER"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid msarti 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   6376
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorFixed  =   -2147483628
      BackColorBkg    =   16761024
      Appearance      =   0
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   615
      Left            =   0
      Top             =   840
      Width           =   9135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LISTE DES EQUIPEMENTS PAR TYPE"
      BeginProperty Font 
         Name            =   "Eras Demi ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   9135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LISTE DES EQUIPEMENTS PAR FOURNISSEUR"
      BeginProperty Font 
         Name            =   "Eras Demi ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   5895
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   4200
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   6000
      Top             =   5520
      Width           =   1455
   End
End
Attribute VB_Name = "FRMLISTEEQUIPEMENTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbfourni_Click()
cmbtype.Text = ""
Set rs2 = New ADODB.Recordset
rs2.Open "select * from EQUIPEMENTS where [Nom_fournisseur] like '" & cmbfourni.Text & "'", cn, 1, 2
msarti.Clear
msarti.Rows = 1
msarti.FormatString = "Designation             |Réference          |Modèle         |Prix             |Type       |Marque     |Fournisseur      |Qte en Stock      |Stock Maxi  |Stock Alerte     |Etat du stock     |Date MAJ    |Emplacement       "
Do While Not rs2.EOF
msarti.AddItem rs2.Fields(0) & vbTab & rs2.Fields(1) & vbTab & rs2.Fields(2) & vbTab & rs2.Fields(4) & vbTab & rs2.Fields(5) & vbTab & rs2.Fields(3) & vbTab & rs2.Fields(12) & vbTab & rs2.Fields(6) & vbTab & rs2.Fields(7) & vbTab & rs2.Fields(8) & vbTab & rs2.Fields(13) & vbTab & rs2.Fields(11) & vbTab & rs2.Fields(9)
rs2.MoveNext
Loop
End Sub

Private Sub cmbtype_Click()
cmbfourni.Text = ""
Set rs2 = New ADODB.Recordset
rs2.Open "select * from EQUIPEMENTS where [Type] like '" & cmbtype.Text & "'", cn, 1, 2
msarti.Clear
msarti.Rows = 1
msarti.FormatString = "Designation             |Réference          |Modèle         |Prix             |Type       |Marque     |Fournisseur      |Qte en Stock      |Stock Maxi  |Stock Alerte     |Etat du stock     |Date MAJ    |Emplacement       "
Do While Not rs2.EOF
msarti.AddItem rs2.Fields(0) & vbTab & rs2.Fields(1) & vbTab & rs2.Fields(2) & vbTab & rs2.Fields(4) & vbTab & rs2.Fields(5) & vbTab & rs2.Fields(3) & vbTab & rs2.Fields(12) & vbTab & rs2.Fields(6) & vbTab & rs2.Fields(7) & vbTab & rs2.Fields(8) & vbTab & rs2.Fields(13) & vbTab & rs2.Fields(11) & vbTab & rs2.Fields(9)
rs2.MoveNext
Loop
End Sub

Private Sub CMDIMPRIMER_Click()
affichertout
End Sub

Private Sub CMDIMPRIMER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDIMPRIMER.BackColor = &HFFFFFF
End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("Voulez-vous vraiment Imprimer la liste", vbQuestion + vbYesNo) = vbYes Then
   Dim objXL As Excel.Application
    Dim objWB As Excel.Workbook
    Dim objWS As Excel.Worksheet
    Dim r As Long
    Dim C As Long
    Dim intRed As Integer
    Dim intGreen As Integer
    Dim intBlue As Integer
    
    Set objXL = New Excel.Application
    Set objWB = objXL.Workbooks.Add
    Set objWS = objWB.Worksheets(1)

    With objWS
        For r = 0 To msarti.Rows - 1
            For C = 0 To msarti.Cols - 1
                .Cells(r + 1, C + 1) = msarti.TextMatrix(r, C)
            Next
        Next

        .Cells.Columns.AutoFit
    End With

    objXL.Visible = True
    
    Set objWS = Nothing
    Set objWB = Nothing
    Set objXL = Nothing
Else
Exit Sub
End If
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = &HFFFFFF
End Sub



Private Sub Form_Load()
On Error Resume Next
connect
Set rs = New ADODB.Recordset
rs.Open "select * from fournisseur", cn, 1, 2
Do While Not rs.EOF
cmbfourni.AddItem rs.Fields(0)
rs.MoveNext
Loop
rs.Close
Set rs = Nothing
affichertout
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDIMPRIMER.BackColor = &HFFFF&  '  &HE0E0E0
Command2.BackColor = &HFFFF&    '  &HE0E0E0
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FRMMENU.Show
End Sub

Private Sub Label1_Click()
Unload Me
FRMMENU.Show
End Sub


Sub affichertout()
On Error Resume Next
msarti.Clear
msarti.Rows = 1
msarti.FormatString = "Designation             |Réference          |Modèle         |Prix             |Type       |Marque     |Fournisseur      |Qte en Stock      |Stock Maxi  |Stock Alerte     |Etat du stock     |Date MAJ    |Emplacement       "
Set rsEQUIPEMENT = New ADODB.Recordset
rsEQUIPEMENT.Open "select * from EQUIPEMENTS order by [Etat du stock]", cn, 1, 2
Do While Not rsEQUIPEMENT.EOF
msarti.AddItem rsEQUIPEMENT.Fields(0) & vbTab & rsEQUIPEMENT.Fields(1) & vbTab & rsEQUIPEMENT.Fields(2) & vbTab & rsEQUIPEMENT.Fields(4) & vbTab & rsEQUIPEMENT.Fields(5) & vbTab & rsEQUIPEMENT.Fields(3) & vbTab & rsEQUIPEMENT.Fields(12) & vbTab & rsEQUIPEMENT.Fields(6) & vbTab & rsEQUIPEMENT.Fields(7) & vbTab & rsEQUIPEMENT.Fields(8) & vbTab & rsEQUIPEMENT.Fields(13) & vbTab & rsEQUIPEMENT.Fields(11) & vbTab & rsEQUIPEMENT.Fields(9)
rsEQUIPEMENT.MoveNext
Loop
rsEQUIPEMENT.Close
Set rsEQUIPEMENT = Nothing
End Sub


