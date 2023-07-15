VERSION 5.00
Begin VB.Form dloginw 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DEALER DASH BOARD"
   ClientHeight    =   5550
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "dloginw.frx":0000
   ScaleHeight     =   5550
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "LOGOUT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "                                       YOUR DETAILS                                            "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2895
      Left            =   3120
      TabIndex        =   5
      Top             =   1800
      Width           =   5535
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "You are not Allowed to Edit Your Details.If There are any Corrections Please Bring The Information to Company Notice !"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   600
         TabIndex        =   15
         Top             =   2280
         Width           =   4275
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   480
      End
      Begin VB.Label lbla 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   13
         Top             =   1560
         Width           =   105
      End
      Begin VB.Label lblc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Top             =   1200
         Width           =   105
      End
      Begin VB.Label lbln 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Top             =   840
         Width           =   105
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS  :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT NO  :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NAME :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   8
         Top             =   840
         Width           =   585
      End
      Begin VB.Label lblu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "YOUR USER ID :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   6
         Top             =   480
         Width           =   1320
      End
   End
   Begin VB.Image Image1 
      Height          =   3975
      Index           =   1
      Left            =   0
      Picture         =   "dloginw.frx":63F0
      Stretch         =   -1  'True
      Top             =   -1320
      Width           =   3600
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
      Left            =   6120
      TabIndex        =   17
      Top             =   5280
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
      Left            =   7320
      TabIndex        =   16
      Top             =   5280
      Width           =   3315
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DATE :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   7800
      TabIndex        =   4
      Top             =   240
      Width           =   660
   End
   Begin VB.Label date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   8520
      TabIndex        =   3
      Top             =   240
      Width           =   105
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
      Left            =   4080
      TabIndex        =   2
      Top             =   1440
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
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   7200
      TabIndex        =   1
      Top             =   960
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
      Left            =   4080
      TabIndex        =   0
      Top             =   960
      Width           =   2865
   End
   Begin VB.Menu f 
      Caption         =   "FARMER"
      Begin VB.Menu af 
         Caption         =   "ADD FARMER"
         Shortcut        =   ^F
      End
      Begin VB.Menu uf 
         Caption         =   "UPDATE (or) DELETE FARMERS"
         Shortcut        =   ^U
      End
      Begin VB.Menu fr 
         Caption         =   "FARMER REPORTS"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu md 
      Caption         =   "MILK DETAILS"
      Index           =   2
      Begin VB.Menu umd 
         Caption         =   "UPDATE MILK DETAIL"
         Shortcut        =   ^M
      End
      Begin VB.Menu dr 
         Caption         =   "MILK REPORTES"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu cu 
      Caption         =   "CONTACT US"
   End
   Begin VB.Menu au 
      Caption         =   "ABOUT US"
   End
   Begin VB.Menu cyp 
      Caption         =   "CHANGE YOUR PASSWORD"
   End
   Begin VB.Menu l 
      Caption         =   "LOGOUT"
   End
End
Attribute VB_Name = "dloginw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub af_Click()
affrm.Show
End Sub
Private Sub au_Click()
aboutus.Show
End Sub
Private Sub Command1_Click()
answer = MsgBox("Do you want to LOGOUT?", vbExclamation + vbYesNo, "Warning")
If answer = vbYes Then
Me.Hide
Else
End If
End Sub

Private Sub cu_Click()
contfrm.Show
End Sub
Private Sub cyp_Click()
dpw.Show
End Sub
Private Sub dr_Click()
dmrw.Show
End Sub
Private Sub Form_Load()
lblu = dlogin.txtu.Text
date = DateTime.Now
connectdb
Set rs = con.Execute("select * from Dealerdb")
Set rs = con.Execute("select * from Dealerdb where uname='" + lblu.Caption + "'")
If (Not rs.EOF) Then
lbln = rs("Name")
lblc = rs("CONTACT")
lbla = rs("ADDRESS")
Else
End If
End Sub
Private Sub fr_Click()
dffr.Show
End Sub
Private Sub l_Click()
dloginw.Hide
End Sub
Private Sub uf_Click()
fduw.Show
End Sub
Private Sub umd_Click()
mduw.Show
End Sub
