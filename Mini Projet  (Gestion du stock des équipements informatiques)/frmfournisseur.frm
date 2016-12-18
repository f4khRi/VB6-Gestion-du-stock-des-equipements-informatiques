VERSION 5.00
Begin VB.Form frmfournisseur 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FOURNISSEURS"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11700
   Icon            =   "frmfournisseur.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Mailcontact 
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
      Height          =   375
      Left            =   8280
      TabIndex        =   22
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox Gsmcontact 
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
      Height          =   375
      Left            =   8280
      TabIndex        =   21
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox Faxcontact 
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
      Height          =   375
      Left            =   8280
      TabIndex        =   20
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox Télfixcontact 
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
      Height          =   375
      Left            =   8280
      TabIndex        =   19
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox Servicecontact 
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
      Height          =   375
      Left            =   8280
      TabIndex        =   18
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox Fonctioncontact 
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
      Height          =   375
      Left            =   8280
      TabIndex        =   17
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox Prénomcontact 
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
      Height          =   375
      Left            =   8280
      TabIndex        =   16
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox nomcontact 
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
      Height          =   375
      Left            =   8280
      TabIndex        =   15
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox delaireglement 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   5640
      Width           =   2895
   End
   Begin VB.TextBox Reglement 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   5160
      Width           =   2895
   End
   Begin VB.TextBox Paiement 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox WebFourni 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   4200
      Width           =   2895
   End
   Begin VB.TextBox pays 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox ville 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox cp 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox adressefourni 
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
      Height          =   855
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtnomfourni 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      ToolTipText     =   "Numéro de série de la machine"
      Top             =   1320
      Width           =   2895
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
      Top             =   7320
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
      TabIndex        =   4
      Top             =   7320
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
      TabIndex        =   3
      Top             =   7320
      Width           =   1455
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
      TabIndex        =   2
      Top             =   7320
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
      TabIndex        =   1
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton CMDIMPRIMER 
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
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   0
      Top             =   6480
      Width           =   5415
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
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
      Index           =   27
      Left            =   6120
      TabIndex        =   42
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Téléphone Mobile"
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
      Index           =   26
      Left            =   6120
      TabIndex        =   41
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
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
      Index           =   25
      Left            =   6120
      TabIndex        =   40
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Téléphone Fixe"
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
      Index           =   24
      Left            =   6120
      TabIndex        =   39
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Délai de paiement"
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
      Index           =   23
      Left            =   360
      TabIndex        =   38
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Service"
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
      Index           =   22
      Left            =   6120
      TabIndex        =   37
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fonction"
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
      Index           =   21
      Left            =   6120
      TabIndex        =   36
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Adresse"
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
      Index           =   20
      Left            =   360
      TabIndex        =   35
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Code Postal"
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
      Index           =   19
      Left            =   360
      TabIndex        =   34
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ville"
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
      Index           =   18
      Left            =   360
      TabIndex        =   33
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pays"
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
      Index           =   17
      Left            =   360
      TabIndex        =   32
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Site web"
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
      Index           =   12
      Left            =   360
      TabIndex        =   31
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Condition Paiement"
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
      Index           =   11
      Left            =   360
      TabIndex        =   30
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mode de paiment"
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
      Index           =   10
      Left            =   360
      TabIndex        =   29
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom"
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
      Index           =   6
      Left            =   6120
      TabIndex        =   28
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Prénom"
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
      Index           =   5
      Left            =   6120
      TabIndex        =   27
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom Fournisseur"
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
      Index           =   4
      Left            =   360
      TabIndex        =   26
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "INFORMATIONS FOURNISSEURS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   25
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "CONTACT/ FOURNISSEURS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   24
      Top             =   600
      Width           =   5415
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   6000
      Top             =   6480
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   6000
      Top             =   240
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      BorderStyle     =   0  'Transparent
      Height          =   6375
      Left            =   5640
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11400
      TabIndex        =   23
      Top             =   -480
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   0
      Top             =   240
      Width           =   5415
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   8160
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   6120
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   2040
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   120
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   4080
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   4080
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   4080
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Image suivant 
      Height          =   480
      Left            =   5880
      Picture         =   "frmfournisseur.frx":0CCA
      ToolTipText     =   "SUIVANT"
      Top             =   6720
      Width           =   480
   End
   Begin VB.Image Precedent 
      Height          =   480
      Left            =   5160
      Picture         =   "frmfournisseur.frx":0E96
      ToolTipText     =   "PRECEDENT"
      Top             =   6720
      Width           =   480
   End
   Begin VB.Image premier 
      Height          =   480
      Left            =   4440
      Picture         =   "frmfournisseur.frx":105E
      ToolTipText     =   "PREMIER"
      Top             =   6720
      Width           =   480
   End
   Begin VB.Image Dernier 
      Height          =   480
      Left            =   6600
      Picture         =   "frmfournisseur.frx":1247
      ToolTipText     =   "DERNIER"
      Top             =   6720
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      Height          =   495
      Left            =   10200
      Top             =   7440
      Width           =   1455
   End
End
Attribute VB_Name = "frmfournisseur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDENREGISTRER_Click()
On Error GoTo er
'****************************************
If Trim(txtnomfourni) = "" Then
MsgBox "Le nom du fournisseur est obligatoire", vbExclamation
txtnomfourni.SetFocus
Exit Sub
End If
'****************************************
rsfourni.AddNew
rsfourni.Fields(0) = txtnomfourni
rsfourni.Fields(1) = adressefourni
rsfourni.Fields(2) = cp
rsfourni.Fields(3) = ville
rsfourni.Fields(4) = pays
rsfourni.Fields(5) = WebFourni
rsfourni.Fields(6) = Paiement
rsfourni.Fields(7) = Reglement
rsfourni.Fields(8) = delaireglement
'***************************************
rsfourni.Fields(9) = nomcontact
rsfourni.Fields(10) = Prénomcontact
rsfourni.Fields(11) = Fonctioncontact
rsfourni.Fields(12) = Servicecontact
rsfourni.Fields(13) = Télfixcontact
rsfourni.Fields(14) = Faxcontact
rsfourni.Fields(15) = Gsmcontact
rsfourni.Fields(16) = Mailcontact
'***************************************
rsfourni.Update
MsgBox "Enregistrement effectué", vbInformation
'**************************
er:
Select Case Err.Number
Case -2147217887
MsgBox "Nom de fournisseur existe déja", vbCritical
rsfourni.CancelUpdate
End Select
'**************************

End Sub
Sub affichage()
On Error Resume Next
txtnomfourni = rsfourni.Fields(0)
adressefourni = rsfourni.Fields(1)
cp = rsfourni.Fields(2)
ville = rsfourni.Fields(3)
pays = rsfourni.Fields(4)
WebFourni = rsfourni.Fields(5)
Paiement = rsfourni.Fields(6)
Reglement = rsfourni.Fields(7)
delaireglement = rsfourni.Fields(8)
nomcontact = rsfourni.Fields(9)
Prénomcontact = rsfourni.Fields(10)
Fonctioncontact = rsfourni.Fields(11)
Servicecontact = rsfourni.Fields(12)
Télfixcontact = rsfourni.Fields(13)
Faxcontact = rsfourni.Fields(14)
Gsmcontact = rsfourni.Fields(15)
Mailcontact = rsfourni.Fields(16)
End Sub

Private Sub CMDRECHERCHER_Click()
On Error Resume Next
rsfourni.MoveFirst
Code = InputBox("Veuillez saisir le nom du fournisseur à rechercher", RECHERCHE)
'If Code = "" Then
'MsgBox "Veuillez saisir un Modèle SVP !"
'End If
rsfourni.Find "[Nom_fournisseur]='" & Code & "'"
If rsfourni.EOF Then
MsgBox "Fournisseur introuvable", vbCritical
Else
affichage
End If

End Sub

Private Sub CMDMODIFIER_Click()
On Error Resume Next
If Trim(txtnomfourni) = "" Then
    MsgBox "Nom du fournisseur est vide", vbCritical
    Exit Sub
End If
If txtnomfourni = rsfourni.Fields(0) Then
rsfourni.Fields(1) = adressefourni
rsfourni.Fields(2) = cp
rsfourni.Fields(3) = ville
rsfourni.Fields(4) = pays
rsfourni.Fields(5) = WebFourni
rsfourni.Fields(6) = Paiement
rsfourni.Fields(7) = Reglement
rsfourni.Fields(8) = delaireglement
'***************************************
rsfourni.Fields(9) = nomcontact
rsfourni.Fields(10) = Prénomcontact
rsfourni.Fields(11) = Fonctioncontact
rsfourni.Fields(12) = Servicecontact
rsfourni.Fields(13) = Télfixcontact
rsfourni.Fields(14) = Faxcontact
rsfourni.Fields(15) = Gsmcontact
rsfourni.Fields(16) = Mailcontact
'***************************************
rsfourni.Update
MsgBox "Modification effectuée", vbInformation
Else
MsgBox "Le fournisseur n'est pas enregistrée", vbCritical
End If
End Sub

Private Sub CMDNEW_Click()
Dim C As Control
' vider les textboxs
For Each C In Me
If TypeOf C Is TextBox Then
C.Text = ""
End If
Next
txtnomfourni.SetFocus 'curseur
End Sub

Private Sub CMDSUPPRIMER_Click()
On Error Resume Next
If Trim(txtnomfourni) = "" Then
    MsgBox "Le nom du fournisseur est  vide", vbCritical
    Exit Sub
End If
'***************************
If txtnomfourni = rsfourni.Fields(0) Then
rep = MsgBox("Voulez vous vraiment supprimer Le Fournisseur   " & txtnomfourni, vbYesNo + vbQuestion, " Supprimer")
Select Case rep
Case vbYes
rsfourni.Delete
MsgBox "Fournisseur supprimé", vbInformation
rsfourni.MoveLast
affichage
If rsfourni.EOF Then
CMDNEW_Click
End If
Case vbNo
rsfourni.Cancel
MsgBox "Suppression annulée", vbInformation
End Select
Else
MsgBox "Le Fournisseur n'est pas Enregistrée"
End If

End Sub

Private Sub CMDIMPRIMER_Click()
On Error Resume Next
If txtnomfourni = "" Then
MsgBox "Rien à imprimer", vbInformation
Exit Sub
End If
If MsgBox("Voulez-vous vraiment imprimer Le fiche Fournisseur", vbQuestion + vbYesNo) = vbYes Then
    MsgBox "Veuillez SVP ne pas enregistrer " & vbCrLf & "les modification sur le document Word", vbInformation
Dim wd As Word.Application
Set wd = CreateObject("word.application")
wd.Documents.Open App.Path & "\doc\Fournisseurs.doc"
wd.Visible = True
wd.Selection.GoTo what:=wdGoToBookmark, Name:="NomF"
wd.Selection.TypeText Text:=txtnomfourni.Text
wd.Selection.MoveDown unit:=wdLine, Count:=1
'******************************************
wd.Selection.TypeText Text:=adressefourni
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=cp
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=ville
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=pays
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=WebFourni
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=Paiement
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=Reglement
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=delaireglement
'***************************************************
wd.Selection.GoTo what:=wdGoToBookmark, Name:="NomC"
wd.Selection.TypeText Text:=nomcontact
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=Prénomcontact
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=Fonctioncontact
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=Servicecontact
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=Télfixcontact
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=Faxcontact
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=Gsmcontact
wd.Selection.MoveDown unit:=wdLine, Count:=1
wd.Selection.TypeText Text:=Mailcontact
wd.Selection.MoveDown unit:=wdLine, Count:=1
'*************************************
Else
Exit Sub
End If
End Sub

Private Sub CMDIMPRIMER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDIMPRIMER.BackColor = &HFFFFFF
End Sub

Private Sub Dernier_Click()
On Error Resume Next
rsfourni.MoveLast
affichage
End Sub



Private Sub Form_Load()
On Error Resume Next
'**********************************
connect
Set rsfourni = New ADODB.Recordset
rsfourni.Open "select * from Fournisseur", cn, 1, 2
'**********************************
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FRMMENU.Show
End Sub

Private Sub Label1_Click()
Unload Me
FRMMENU.Show
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "X"
End Sub
Private Sub CMDENREGISTRER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDENREGISTRER.BackColor = &HFFFFFF
End Sub
Private Sub CMDRECHERCHER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDRECHERCHER.BackColor = &HFFFFFF
End Sub
Private Sub CMDMODIFIER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDMODIFIER.BackColor = &HFFFFFF
End Sub

Private Sub CMDNEW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDNEW.BackColor = &HFFFFFF
End Sub
Private Sub CMDSUPPRIMER_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDSUPPRIMER.BackColor = &HFFFFFF
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDIMPRIMER.BackColor = &HFFFF&
CMDNEW.BackColor = &HFFFF&
CMDENREGISTRER.BackColor = &HFFFF&
CMDRECHERCHER.BackColor = &HFFFF&
CMDMODIFIER.BackColor = &HFFFF&
CMDNEW.BackColor = &HFFFF&
CMDSUPPRIMER.BackColor = &HFFFF&
End Sub

Private Sub Precedent_Click()
On Error Resume Next
rsfourni.MovePrevious
If rsfourni.BOF Then
    rsfourni.MoveFirst
    MsgBox "Premier fournisseur", vbInformation
End If
affichage
End Sub

Private Sub premier_Click()
On Error Resume Next
rsfourni.MoveFirst
affichage
End Sub

Private Sub suivant_Click()
On Error Resume Next
rsfourni.MoveNext
If rsfourni.EOF Then
    rsfourni.MoveLast
    MsgBox "Dérnier fournisseur", vbInformation
End If
affichage
End Sub

