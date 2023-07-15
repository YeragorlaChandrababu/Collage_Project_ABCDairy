VERSION 5.00
Begin VB.Form clogw 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "COMPANY LOGIN WINDOW"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   11070
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "clogw.frx":0000
   ScaleHeight     =   6135
   ScaleWidth      =   11070
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "                                        YOUR  DETAILS                                          "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   2640
      TabIndex        =   4
      Top             =   2040
      Width           =   6015
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"clogw.frx":63F0
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   17
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label Label13 
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
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   495
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
         Left            =   2640
         TabIndex        =   15
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
         Left            =   2640
         TabIndex        =   14
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
         Left            =   2640
         TabIndex        =   13
         Top             =   840
         Width           =   105
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
         Left            =   2640
         TabIndex        =   12
         Top             =   480
         Width           =   105
      End
      Begin VB.Label Label10 
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
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   " CONTACT NO :"
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
         Left            =   1200
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label7 
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
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights are Reserved @ ABC DAIRY "
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
      Left            =   6960
      TabIndex        =   18
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   3975
      Index           =   1
      Left            =   -600
      Picture         =   "clogw.frx":6482
      Stretch         =   -1  'True
      Top             =   -1320
      Width           =   4440
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
      Left            =   5760
      TabIndex        =   7
      Top             =   5880
      Width           =   1080
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
      Left            =   7920
      TabIndex        =   6
      Top             =   120
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
      Left            =   8640
      TabIndex        =   5
      Top             =   120
      Width           =   105
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
      Index           =   1
      Left            =   4095
      TabIndex        =   3
      Top             =   1200
      Width           =   2865
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
      TabIndex        =   2
      Top             =   1200
      Width           =   465
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
      ForeColor       =   &H00FF80FF&
      Height          =   270
      Left            =   4080
      TabIndex        =   1
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Menu d 
      Caption         =   "DEALER"
      Index           =   1
      Begin VB.Menu ad 
         Caption         =   "ADD DEALER"
         Shortcut        =   ^D
      End
      Begin VB.Menu ud 
         Caption         =   "UPDATE (or) DELETE DEALERS"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu f 
      Caption         =   "FARMER"
      Index           =   2
      Begin VB.Menu af 
         Caption         =   "ADD FARMER"
         Index           =   2
         Shortcut        =   ^F
      End
      Begin VB.Menu uf 
         Caption         =   "UPDATE (or) DELETE FARMERS"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu r 
      Caption         =   "REPORTS"
      Begin VB.Menu mr 
         Caption         =   "MILK REPORTS"
         Shortcut        =   ^R
      End
      Begin VB.Menu dr 
         Caption         =   "DEALERS REPORTS"
         Shortcut        =   ^S
      End
      Begin VB.Menu fr 
         Caption         =   "FARMERS REPORTS"
         Shortcut        =   ^T
      End
      Begin VB.Menu qr 
         Caption         =   "QUERY REPORT"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu cu 
      Caption         =   "CONTACT US"
   End
   Begin VB.Menu au 
      Caption         =   "ABOUT US"
   End
   Begin VB.Menu cp 
      Caption         =   "CHANGE PASSWORD"
   End
   Begin VB.Menu l 
      Caption         =   "LOGOUT"
   End
End
Attribute VB_Name = "clogw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ad_Click()
adw.Show
End Sub
Private Sub af_Click(Index As Integer)
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
Private Sub cp_Click()
cpw.Show
End Sub
Private Sub cu_Click()
contfrm.Show
End Sub
Private Sub dr_Click()
cdr.Show
End Sub
Private Sub Form_Load()
lblu = clogin.txtu.Text
date = DateTime.Now
connectdb
Set rs = con.Execute("select * from Companydb where uname='" + lblu.Caption + "'")
If (Not rs.EOF) Then
lbln = rs("NAME")
lblc = rs("CONTACT")
lbla = rs("ADDRESS")
Else
End If
End Sub
Private Sub fr_Click()
cfr.Show
End Sub
Private Sub l_Click()
clogw.Hide
End Sub
Private Sub mr_Click()
cmd.Show
End Sub
Private Sub qr_Click()
qrt.Show
End Sub
Private Sub ud_Click()
udud.Show
End Sub
Private Sub uf_Click()
fduw.Show
End Sub
