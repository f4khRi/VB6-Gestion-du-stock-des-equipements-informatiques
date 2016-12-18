VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form GESTEQ 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GESTION DES EQUIPEMENTS (MATERIELS)"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   11760
   Icon            =   "PRODUIT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtmini 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEECF7&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   4800
      Width           =   2775
   End
   Begin VB.TextBox txtsalerte 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEECF7&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   20
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox txtdes 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEECF7&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   19
      ToolTipText     =   "Référence de l'article"
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox txtmarque 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEECF7&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtcode 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEECF7&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   17
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtREF 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEECF7&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox txtetat 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CEECF7&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8400
      TabIndex        =   15
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox txtquantite 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEECF7&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1036
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox txtemplacement 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEECF7&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   8400
      MultiLine       =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "Ce champ peut comporter plusieurs lignes"
      Top             =   2640
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEECF7&
      BorderStyle     =   0  'None
      FillStyle       =   6  'Cross
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   8400
      ScaleHeight     =   2655
      ScaleWidth      =   2775
      TabIndex        =   12
      ToolTipText     =   "Cliquez Pour Agrandir"
      Top             =   4080
      Width           =   2775
   End
   Begin VB.CommandButton CMDNEW 
      BackColor       =   &H0000FFFF&
      Caption         =   "&NOUVEAU"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CMNDMODIFIER 
      BackColor       =   &H0000FFFF&
      Caption         =   "&MODIFIER"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "&NOUVEAU"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CMDMODIFIER 
      BackColor       =   &H0000FFFF&
      Caption         =   "&MODIFIER"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CMDENREGISTRER 
      BackColor       =   &H0000FFFF&
      Caption         =   "&AJOUTER"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CMDSUPPRIMER 
      BackColor       =   &H0000FFFF&
      Caption         =   "&SUPPRIMER"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CMDRECHERCHER 
      BackColor       =   &H0000FFFF&
      Caption         =   "&RECHERCHE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox txtprix 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEECF7&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   6240
      Width           =   2775
   End
   Begin VB.ComboBox cmbtype 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEECF7&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "PRODUIT.frx":0CCA
      Left            =   2760
      List            =   "PRODUIT.frx":0CF2
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.ComboBox cmbfournisseur 
      Appearance      =   0  'Flat
      BackColor       =   &H00CEECF7&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "PRODUIT.frx":0DC8
      Left            =   8400
      List            =   "PRODUIT.frx":0DCA
      TabIndex        =   1
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   6840
   End
   Begin VB.CommandButton CMNDIMPRIMER 
      BackColor       =   &H0000FFFF&
      Caption         =   "&IMPRIMER"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7440
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker datestock 
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12640511
      CalendarTitleBackColor=   16777215
      CalendarTrailingForeColor=   0
      Format          =   114032641
      CurrentDate     =   42581
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   600
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Alerte"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   36
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Etat du Stock"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   35
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      BackStyle       =   0  'Transparent
      Caption         =   "Référence "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   34
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Type de materiel"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Prix"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Marque"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Désignation "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Minimum"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantité en Stock"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   28
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Emplacement"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   27
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Image"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   26
      Top             =   5160
      Width           =   735
   End
   Begin VB.Image imginsérer 
      Height          =   480
      Left            =   6960
      Picture         =   "PRODUIT.frx":0DCC
      ToolTipText     =   "Cliquez pour choisir une image sur votre PC"
      Top             =   5040
      Width           =   480
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   120
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   4080
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   4080
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   4080
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   2040
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   6120
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   8160
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Image Dernier 
      Height          =   480
      Left            =   6360
      Picture         =   "PRODUIT.frx":10D6
      ToolTipText     =   "DERNIER"
      Top             =   6840
      Width           =   480
   End
   Begin VB.Image premier 
      Height          =   480
      Left            =   4560
      Picture         =   "PRODUIT.frx":12C5
      ToolTipText     =   "PREMIER"
      Top             =   6840
      Width           =   480
   End
   Begin VB.Image Precedent 
      Height          =   480
      Left            =   5160
      Picture         =   "PRODUIT.frx":14AE
      ToolTipText     =   "PRECEDENT"
      Top             =   6840
      Width           =   480
   End
   Begin VB.Image suivant 
      Height          =   480
      Left            =   5760
      Picture         =   "PRODUIT.frx":1676
      ToolTipText     =   "SUIVANT"
      Top             =   6840
      Width           =   480
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Modele"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   25
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Changement Stock"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fournisseur"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   23
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label TXTIMAGE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   10200
      Top             =   7560
      Width           =   1455
   End
End
Attribute VB_Name = "GESTEQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmbfournisseur_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub cmbtype_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub CMDENREGISTRER_Click()
On Error GoTo er
If Trim(txtREF) = "" Then
MsgBox "Entrez la référence SVP", vbExclamation
txtREF.SetFocus
Exit Sub
End If
If Trim(txtdes) = "" Then
MsgBox "Veuillez saisir la désignation SVP", vbExclamation
txtdes.SetFocus
Exit Sub
End If
'If Trim(txtcode) = rsEQUIPEMENT.Fields(2) And rsEQUIPEMENT.RecordCount < 0 Then
'MsgBox "Le Modèle est déjà utilisee", vbInformation
'txtcode.SetFocus
'Exit Sub
'End If
'If Trim(cmbfournisseur.Text) = "" Then
'MsgBox "Veuillez choisir le Fournisseur SVP", vbExclamation
'Exit Sub
'End If
'****************************************
'****************************************
rsEQUIPEMENT.AddNew
rsEQUIPEMENT.Fields(0) = txtdes
rsEQUIPEMENT.Fields(1) = txtREF
rsEQUIPEMENT.Fields(2) = txtcode
rsEQUIPEMENT.Fields(3) = txtmarque
rsEQUIPEMENT.Fields(4) = txtprix
rsEQUIPEMENT.Fields(5) = cmbtype.Text
rsEQUIPEMENT.Fields(6) = txtquantite
rsEQUIPEMENT.Fields(7) = txtmini
rsEQUIPEMENT.Fields(8) = txtsalerte
rsEQUIPEMENT.Fields(9) = txtemplacement
rsEQUIPEMENT.Fields(10) = TXTIMAGE.Caption
rsEQUIPEMENT.Fields(11) = datestock.Value
rsEQUIPEMENT.Fields(12) = cmbfournisseur.Text
rsEQUIPEMENT.Fields(13) = txtetat
rsEQUIPEMENT.Update
MsgBox "Equipement ajouté avec succès", vbInformation
'**************************
er:
Select Case Err.Number
Case -2147217887
MsgBox "La référence existe déja", vbCritical
rsEQUIPEMENT.CancelUpdate
End Select
'**************************

End Sub

Private Sub CMDENREGISTRER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDENREGISTRER.BackColor = &HFFFFFF
End Sub

Private Sub CMDRECHERCHER_Click()
On Error Resume Next
rsEQUIPEMENT.MoveFirst
Code = InputBox("Veuillez saisir le modèle à rechercher", RECHERCHE)
'If Code = "" Then
'MsgBox "Veuillez saisir un Modèle SVP !"
'End If
rsEQUIPEMENT.Find "[Modele]='" & Code & "'"
If rsEQUIPEMENT.EOF Then
MsgBox "Code introuvable!", vbCritical
Else
affichage
End If
End Sub

Private Sub CMDRECHERCHER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDRECHERCHER.BackColor = &HFFFFFF
End Sub



Private Sub CMDNEW_Click()
Dim C As Control
For Each C In Me
' vider les textbox et les combox
If TypeOf C Is TextBox Then
C.Text = ""
End If
If TypeOf C Is ComboBox Then
C.Text = ""
End If
Next
TXTIMAGE.Caption = ""
txtdes.SetFocus
txtetat.BackColor = &HCEECF7
txtetat = ""
''''''''Pour initialiser l'image
Picture1.Picture = Nothing
cmbtype.ListIndex = 4
datestock.Value = Date
End Sub

Private Sub CMDNEW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDNEW.BackColor = &HFFFFFF
End Sub



Private Sub CMDSUPPRIMER_Click()
On Error Resume Next
If Trim(txtREF) = "" Then
    MsgBox "Rien a supprimer la référence est vide", vbCritical
    Exit Sub
End If
'***************************
If txtREF = rsEQUIPEMENT.Fields(1) Then
rep = MsgBox("Voulez vous vraiment supprimer l'équipement  " & txtdes, vbYesNo + vbQuestion, " Supprimer")
Select Case rep
Case vbYes
rsEQUIPEMENT.Delete
MsgBox "Equipement supprimée", vbInformation
rsEQUIPEMENT.MoveLast
affichage
If rsEQUIPEMENT.EOF Then
CMDNEW_Click
End If
Case vbNo
rsEQUIPEMENT.Cancel
MsgBox "Suppression annulée", vbInformation
End Select
Else
MsgBox "L'équipement n'est pas enregistrée"
End If
End Sub

Private Sub CMDSUPPRIMER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDSUPPRIMER.BackColor = &HFFFFFF
End Sub




Private Sub CMNDMODIFIER_Click()

On Error Resume Next
If Trim(txtREF) = "" Then
    MsgBox "Le champ référence est vide", vbCritical
    Exit Sub
End If
If txtREF = rsEQUIPEMENT.Fields(1) Then
rsEQUIPEMENT.Fields(0) = txtdes
rsEQUIPEMENT.Fields(2) = txtcode
rsEQUIPEMENT.Fields(3) = txtmarque
rsEQUIPEMENT.Fields(4) = txtprix
rsEQUIPEMENT.Fields(5) = cmbtype.Text
rsEQUIPEMENT.Fields(6) = txtquantite
rsEQUIPEMENT.Fields(7) = txtmini
rsEQUIPEMENT.Fields(8) = txtsalerte
rsEQUIPEMENT.Fields(9) = txtemplacement
rsEQUIPEMENT.Fields(10) = TXTIMAGE.Caption
rsEQUIPEMENT.Fields(11) = datestock.Value
rsEQUIPEMENT.Fields(12) = cmbfournisseur.Text
rsEQUIPEMENT.Fields(13) = txtetat
rsEQUIPEMENT.UpdateBatch adAffectAllChapters
MsgBox "Modification effectuée", vbInformation
Else
MsgBox "L'equipement n'est pas enregistrée", vbCritical
End If


End Sub

Private Sub CMNDMODIFIERR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMNDMODIFIER.BackColor = &HFFFFFF
End Sub

Private Sub CMNDIMPRIMER_Click()
On Error Resume Next
If txtREF = "" Then
MsgBox "Rien à Imprimer", vbInformation
Exit Sub
End If
If MsgBox("Voulez-vous vraiment imprimer Le fiche EQUIPEMENT", vbQuestion + vbYesNo) = vbYes Then
    MsgBox "SVP ne pas enregistrer " & vbCrLf & "les modification sur le document Word", vbInformation
Dim wd As Word.Application
Set wd = CreateObject("word.application")
wd.Documents.Open App.Path & "\doc\EQUIPEMENTS.doc"
wd.Visible = True
wd.Selection.GoTo what:=wdGoToBookmark, Name:="numserie"
wd.Selection.TypeText Text:=txtdes.Text
wd.Selection.MoveDown unit:=wdLine, Count:=1
'******************************************
wd.Selection.TypeText Text:=txtREF
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=txtcode
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=txtmarque
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=cmbtype.Text
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=txtquantite
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=txtmini
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=txtsalerte
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=txtprix
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=datestock.Value
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=txtetat
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=cmbfournisseur.Text
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=txtemplacement.Text
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.GoTo what:=wdGoToBookmark, Name:="image"
wd.Selection.InlineShapes.AddPicture FileName:=TXTIMAGE.Caption, linktofile:=False, savewithdocument:=True
Else
Exit Sub
End If
End Sub

Private Sub CMNDIMPRIMER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMNDIMPRIMER.BackColor = &HFFFFFF
End Sub

Private Sub Dernier_Click()
On Error Resume Next
rsEQUIPEMENT.MoveLast
affichage
End Sub



Private Sub Form_Load()
On Error Resume Next
cmbtype.ListIndex = 4
'***********************
connect
Set rsEQUIPEMENT = New ADODB.Recordset
rsEQUIPEMENT.Open "select * from EQUIPEMENTS", cn, 1, 2
'**********************************
Set rsfourni = New ADODB.Recordset
rsfourni.Open "select * from fournisseur", cn, 1, 2
Do Until rsfourni.EOF
cmbfournisseur.AddItem rsfourni.Fields(0)
rsfourni.MoveNext
Loop
rsfourni.Close
Set rsfourni = Nothing
'**********************************
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FRMMENU.Show
End Sub

Private Sub txtmini_Change()
Timer1.Enabled = True
End Sub

Private Sub txtmini_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Not IsNumeric(Chr(KeyAscii)) Then
MsgBox "Veuillez entrez un nombre valide SVP", vbInformation
KeyAscii = 0
End If
End If
End Sub

Private Sub txtmini_LostFocus()
If txtmini.Text = "" Then
txtmini.Text = 0
End If
End Sub


Private Sub txtquantite_Change()
Timer1.Enabled = True
datestock.Value = Date
End Sub

Private Sub txtquantite_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Not IsNumeric(Chr(KeyAscii)) Then
MsgBox "Veuillez Entrez un nombre valide SVP", vbInformation
KeyAscii = 0
End If
End If
End Sub
Sub affichage()
On Error Resume Next
'moins.Text = Val(Val(txtquantite) - Val(txtmini))
'***************************
'If Val(moins.Text) < Val(txtmini) + 3 Then
'    LBLETAT.Caption = "ETAT DU STOCK EN ALERTE"
'    Else
'If Val(moins.Text) > Val(txtmini) + 3 Then
'    LBLETAT.Caption = "ETAT DU STOCK NORMAL"
'    LBLETAT.BackColor = vbWhite
'    Timer1.Enabled = False
'    End If
'End If
 Timer1.Enabled = True
'***************************
txtREF = rsEQUIPEMENT.Fields(1)
txtdes = rsEQUIPEMENT.Fields(0)
txtmarque = rsEQUIPEMENT.Fields(3)
txtprix = rsEQUIPEMENT.Fields(4)
cmbtype.Text = rsEQUIPEMENT.Fields(5)
txtcode = rsEQUIPEMENT.Fields(2)
txtquantite = rsEQUIPEMENT.Fields(6)
txtmini = rsEQUIPEMENT.Fields(7)
txtsalerte = rsEQUIPEMENT.Fields(8)
cmbfournisseur.Text = rsEQUIPEMENT.Fields(12)
datestock.Value = rsEQUIPEMENT.Fields(11)
TXTIMAGE.Caption = rsEQUIPEMENT.Fields(10)
txtemplacement = rsEQUIPEMENT.Fields(9)
txtetat = rsEQUIPEMENT.Fields(13)
If TXTIMAGE.Caption = "" Then
Picture1.Picture = Nothing
Else
Picture1.Picture = LoadPicture(TXTIMAGE.Caption)
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMNDIMPRIMER.BackColor = &HFFFF&
CMDNEW.BackColor = &HFFFF&
CMDENREGISTRER.BackColor = &HFFFF&
CMDRECHERCHER.BackColor = &HFFFF&
CMNDMODIFIER.BackColor = &HFFFF&
CMDSUPPRIMER.BackColor = &HFFFF&
End Sub
Private Sub imginsérer_Click()
On Error GoTo inc
Dim chemin As String
Dialog.ShowOpen
chemin = Dialog.FileName
Picture1.Picture = LoadPicture(chemin)
TXTIMAGE.Caption = chemin
inc:
Select Case Err.Number
Case 481
MsgBox "Image incorrect", vbCritical
End Select
End Sub

Private Sub Label12_Click()
Unload Me
FRMMENU.Show
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.Caption = "X"
End Sub

Private Sub Picture1_Click()
Rimage.Image1.Picture = Picture1.Picture
Rimage.Height = Rimage.Image1.Height
Rimage.Width = Rimage.Image1.Width
' zoomer le photo dans le form Rimage
Rimage.Show
End Sub

Private Sub Precedent_Click()
On Error Resume Next
rsEQUIPEMENT.MovePrevious
If rsEQUIPEMENT.BOF And rsEQUIPEMENT.RecordCount > 0 Then
    rsEQUIPEMENT.MoveFirst
    MsgBox "Premier equipement enregistrée", vbInformation
End If
affichage
End Sub

Private Sub premier_Click()
On Error Resume Next
rsEQUIPEMENT.MoveFirst
affichage
End Sub

Private Sub suivant_Click()
On Error Resume Next
rsEQUIPEMENT.MoveNext
If rsEQUIPEMENT.EOF And rsEQUIPEMENT.RecordCount > 0 Then
    rsEQUIPEMENT.MoveLast
    MsgBox "Dernier equipement enregistré", vbInformation
End If
affichage
End Sub

Private Sub Timer1_Timer()
If txtsalerte = "" And txtquantite = "" Then
txtetat.BackColor = &HCEECF7
txtetat = ""
Exit Sub
End If
If Val(txtmini) > Val(txtquantite) Or Val(txtquantite.Text) <= Val(txtsalerte) Then
txtetat.FontBold = True
txtetat.ForeColor = vbRed
txtetat = "STOCK EN ALERTE"
txtetat.BackColor = Rnd * 255
    Else
    txtetat.FontBold = True
txtetat.ForeColor = &H8000&
txtetat = "STOCK NORMAL"
txtetat.BackColor = &HCEECF7
End If
End Sub

Private Sub txtsalerte_Change()
Timer1.Enabled = True
End Sub

Private Sub txtsalerte_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If Not IsNumeric(Chr(KeyAscii)) Then
MsgBox "Veuillez Entrez un nombre valide SVP", vbInformation
KeyAscii = 0
End If
End If
End Sub

Private Sub txtsalerte_LostFocus()
If txtsalerte.Text = "" Then
txtsalerte.Text = 0
End If
End Sub

