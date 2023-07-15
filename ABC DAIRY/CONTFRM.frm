VERSION 5.00
Begin VB.Form contfrm 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DAIRY DAY - CONTACT"
   ClientHeight    =   8325
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   13485
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CONTFRM.frx":0000
   ScaleHeight     =   8325
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "ASK YOUR QUESTION"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      MaskColor       =   &H0080FF80&
      Picture         =   "CONTFRM.frx":804C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "YOU WILL BE REDIRECTED TO CONTACT FORM"
      Top             =   7680
      Width           =   2895
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   480
      Picture         =   "CONTFRM.frx":1B32F
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   7560
      Picture         =   "CONTFRM.frx":1EB05
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   3975
      Index           =   1
      Left            =   -720
      Picture         =   "CONTFRM.frx":2108A
      Stretch         =   -1  'True
      Top             =   -1080
      Width           =   5280
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights are Reserved @ ABC DAIRY "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   2
      Left            =   10080
      TabIndex        =   16
      Top             =   7920
      Width           =   3300
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING :"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   1
      Left            =   8880
      TabIndex        =   15
      Top             =   7920
      Width           =   1080
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "E-MAIL ID                  :  XXXXX@GMAIL.COM"
      Height          =   285
      Left            =   8760
      TabIndex        =   14
      Top             =   7080
      Width           =   3870
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MOBILE NO              :  XXXXXXXXXXXXXX"
      Height          =   285
      Left            =   8760
      TabIndex        =   13
      Top             =   6720
      Width           =   3795
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ASSISTANT NAME :  XXXXXXXXXXXXXX"
      Height          =   285
      Left            =   8760
      TabIndex        =   12
      Top             =   6360
      Width           =   3795
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "E-MAIL ID                 : XXXXXXXXX@GMAIL.COM"
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   6720
      Width           =   4320
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MOBILE NO             : XXXXXXXXXXXXX"
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   7080
      Width           =   3570
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MANAGER NAME   : XXXXXXXXXXXXX"
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   6360
      Width           =   3585
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VISION OF THE CONPANY :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      TabIndex        =   8
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONTACT DETAILS :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   360
      TabIndex        =   7
      Top             =   5880
      Width           =   3195
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"CONTFRM.frx":23B2C
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   6
      Top             =   4800
      Width           =   12855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"CONTFRM.frx":23C56
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   3720
      Width           =   12975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"CONTFRM.frx":23DA5
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   12975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVIDING HEALTH TO INDIA"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4800
      TabIndex        =   3
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   7920
      TabIndex        =   2
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ABC DAIRY"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   585
      Left            =   4815
      TabIndex        =   0
      Top             =   1200
      Width           =   2865
   End
   Begin VB.Menu login 
      Caption         =   "LOGIN"
      Begin VB.Menu cal 
         Caption         =   "COMPANY ADMIN LOGIN"
         Shortcut        =   ^C
      End
      Begin VB.Menu dl 
         Caption         =   "DEALER LOGIN"
         Shortcut        =   ^D
      End
      Begin VB.Menu fl 
         Caption         =   "FORMER LOGIN"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu au 
      Caption         =   "ABOUT US"
   End
   Begin VB.Menu raq 
      Caption         =   "ASK YOUR QUESTION"
   End
End
Attribute VB_Name = "contfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub au_Click()
aboutus.Show
dlogin.Hide
clogin.Hide
flogin.Hide
connfrm.Hide
End Sub
Private Sub cal_Click()
clogin.Show
dlogin.Hide
aboutus.Hide
connfrm.Hide
flogin.Hide
End Sub
Private Sub Command1_Click()
connfrm.Show
clogin.Hide
dlogin.Hide
flogin.Hide
aboutus.Hide
End Sub
Private Sub dl_Click()
dlogin.Show
connfrm.Hide
clogin.Hide
aboutus.Hide
flogin.Hide
End Sub
Private Sub fl_Click()
flogin.Show
connfrm.Hide
dlogin.Hide
clogin.Hide
aboutus.Hide
End Sub


Private Sub raq_Click()
flogin.Hide
connfrm.Show
dlogin.Hide
clogin.Hide
aboutus.Hide
End Sub
