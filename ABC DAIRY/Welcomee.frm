VERSION 5.00
Begin VB.Form Welcomee 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WELCOME TO ABC DAIRY"
   ClientHeight    =   8595
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Welcomee.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ABOUT US"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF80FF&
      Caption         =   "CONTACT US"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Image Image6 
      Height          =   1305
      Left            =   10440
      Picture         =   "Welcomee.frx":12929
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1365
   End
   Begin VB.Image Image4 
      Height          =   1275
      Left            =   5640
      Picture         =   "Welcomee.frx":15282
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   240
      Picture         =   "Welcomee.frx":1878E
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Image Image5 
      Height          =   1335
      Left            =   8880
      Picture         =   "Welcomee.frx":1D236
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   1335
      Left            =   2760
      Picture         =   "Welcomee.frx":20098
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   4695
      Index           =   1
      Left            =   -600
      Picture         =   "Welcomee.frx":22681
      Stretch         =   -1  'True
      Top             =   -1320
      Width           =   5280
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Here You Can Contact the Company on Your Questions."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   4200
      TabIndex        =   20
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Here You Can Find All the Information About the Company."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   10320
      TabIndex        =   19
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ABOUT US"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   9720
      TabIndex        =   18
      Top             =   5520
      Width           =   1530
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT US"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   315
      Left            =   3600
      TabIndex        =   17
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Farmers Login Page Here."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   11880
      TabIndex        =   16
      Top             =   3240
      Width           =   2655
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
      Left            =   10080
      TabIndex        =   15
      Top             =   8160
      Width           =   1080
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
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   11280
      TabIndex        =   14
      Top             =   8160
      Width           =   3315
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Dealers(Authorised) Login Page. "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   765
      Left            =   7080
      TabIndex        =   13
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Company Authorised Persons, Technical Assistants of the Company and Application Admins LOGIN Page."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1125
      Left            =   1680
      TabIndex        =   12
      Top             =   3240
      Width           =   3195
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "FARMER LOGIN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   11280
      TabIndex        =   11
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEALER LOGIN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   315
      Left            =   6360
      TabIndex        =   10
      Top             =   2880
      Width           =   2265
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY LOGIN "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   9
      Top             =   2880
      Width           =   2625
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome To ABC Dairy Pvt.Ltd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   555
      Index           =   0
      Left            =   4560
      TabIndex        =   8
      Top             =   1920
      Width           =   7125
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
      ForeColor       =   &H0000FFFF&
      Height          =   585
      Left            =   6495
      TabIndex        =   7
      Top             =   1080
      Width           =   2865
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
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   6360
      TabIndex        =   6
      Top             =   1560
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
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   9600
      TabIndex        =   0
      Top             =   1080
      Width           =   465
   End
   Begin VB.Menu login 
      Caption         =   "LOGIN"
      Begin VB.Menu cad 
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
   Begin VB.Menu cs 
      Caption         =   "CONTACT US"
   End
   Begin VB.Menu as 
      Caption         =   "ABOUT US"
   End
End
Attribute VB_Name = "Welcomee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub as_Click()
aboutus.Show
flogin.Hide
dlogin.Hide
contfrm.Hide
clogin.Hide
End Sub
Private Sub cad_Click()
clogin.Show
dlogin.Hide
flogin.Hide
aboutus.Hide
contfrm.Hide
End Sub
Private Sub Command1_Click()
clogin.Show
aboutus.Hide
contfrm.Hide
dlogin.Hide
flogin.Hide
End Sub
Private Sub Command2_Click()
dlogin.Show
clogin.Hide
flogin.Hide
aboutus.Hide
contfrm.Hide
End Sub

Private Sub Command3_Click()
flogin.Show
clogin.Hide
dlogin.Hide
aboutus.Hide
contfrm.Hide
End Sub
Private Sub Command4_Click()
contfrm.Show
clogin.Hide
dlogin.Hide
flogin.Hide
aboutus.Hide
End Sub
Private Sub Command5_Click()
aboutus.Show
dlogin.Hide
clogin.Hide
flogin.Hide
contfrm.Hide
End Sub
Private Sub cs_Click()
contfrm.Show
clogin.Hide
dlogin.Hide
flogin.Hide
aboutus.Hide
End Sub
Private Sub dl_Click()
dlogin.Show
clogin.Hide
flogin.Hide
aboutus.Hide
contfrm.Hide
End Sub
Private Sub fl_Click()
flogin.Show
clogin.Hide
dlogin.Hide
aboutus.Hide
contfrm.Hide
End Sub
Private Sub Image2_Click()
clogin.Show
aboutus.Hide
contfrm.Hide
dlogin.Hide
flogin.Hide
End Sub
Private Sub Image3_Click()
contfrm.Show
clogin.Hide
dlogin.Hide
flogin.Hide
aboutus.Hide
End Sub

Private Sub Image4_Click()
dlogin.Show
clogin.Hide
flogin.Hide
aboutus.Hide
contfrm.Hide
End Sub

Private Sub Image5_Click()
aboutus.Show
dlogin.Hide
clogin.Hide
flogin.Hide
contfrm.Hide
End Sub

Private Sub Image6_Click()
flogin.Show
clogin.Hide
dlogin.Hide
aboutus.Hide
contfrm.Hide
End Sub
