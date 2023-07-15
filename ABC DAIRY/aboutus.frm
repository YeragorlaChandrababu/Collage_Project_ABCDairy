VERSION 5.00
Begin VB.Form aboutus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ABOUT US"
   ClientHeight    =   7950
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "aboutus.frx":0000
   ScaleHeight     =   7950
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image3 
      Height          =   975
      Left            =   9120
      Picture         =   "aboutus.frx":5883
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   4455
      Left            =   -720
      Picture         =   "aboutus.frx":9059
      Stretch         =   -1  'True
      Top             =   -1440
      Width           =   5280
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights are Reserved @ ABC DAIRY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   1
      Left            =   6720
      TabIndex        =   16
      Top             =   7680
      Width           =   3315
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
      Left            =   5520
      TabIndex        =   15
      Top             =   7680
      Width           =   1080
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MOBILE NO: XXXXXXXXXXXXX"
      Height          =   255
      Index           =   2
      Left            =   5880
      TabIndex        =   14
      Top             =   840
      Width           =   2520
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "E-MAIL ID : XXXXXXXXX@GMAIL.COM"
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   480
      Width           =   3165
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MANAGER NAME   : XXXXXXXXXXXXX"
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   12
      Top             =   120
      Width           =   3045
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Research about milk is conflicting, however, with different studies claiming milk is either good or bad for the body. "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Index           =   1
      Left            =   1080
      TabIndex        =   11
      Top             =   7080
      Width           =   8535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"aboutus.frx":BAFB
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   6240
      Width           =   8175
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   9
      Top             =   7080
      Width           =   90
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   6240
      Width           =   90
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MILK  FACTS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Index           =   0
      Left            =   4680
      TabIndex        =   7
      Top             =   6000
      Width           =   1590
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"aboutus.frx":BBF5
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   10095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"aboutus.frx":BD1F
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   10095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"aboutus.frx":BE6E
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   10095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VISION OF THE CONPANY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   3720
      TabIndex        =   3
      Top             =   2160
      Width           =   4035
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
      ForeColor       =   &H00FF00FF&
      Height          =   270
      Left            =   3960
      TabIndex        =   2
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
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   7080
      TabIndex        =   1
      Top             =   1320
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
      ForeColor       =   &H0080FFFF&
      Height          =   585
      Left            =   3975
      TabIndex        =   0
      Top             =   1320
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
   Begin VB.Menu raq 
      Caption         =   "ASK YOUR QUESTION"
   End
End
Attribute VB_Name = "aboutus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cal_Click()
clogin.Show
connfrm.Hide
dlogin.Hide
flogin.Hide
End Sub
Private Sub dl_Click()
dlogin.Show
flogin.Hide
clogin.Hide
connfrm.Hide
End Sub
Private Sub fl_Click()
connfrm.Hide
clogin.Hide
dlogin.Hide
flogin.Show
End Sub
Private Sub raq_Click()
connfrm.Show
clogin.Hide
dlogin.Hide
flogin.Hide
End Sub
