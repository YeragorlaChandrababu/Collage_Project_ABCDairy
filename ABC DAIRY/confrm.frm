VERSION 5.00
Begin VB.Form connfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CONTACT & FEEDBACK FORM DAIRY DAY"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "confrm.frx":0000
   ScaleHeight     =   7785
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "YOU CAN CONTACT HERE ON YOUR QUERIES"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4935
      Left            =   1200
      TabIndex        =   9
      Top             =   2280
      Width           =   8175
      Begin VB.CommandButton EXIT 
         BackColor       =   &H000000FF&
         Caption         =   "EXIT"
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
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "YOU WILL EXIT FROM THE CONTACT FORM"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.CommandButton SUBMIT 
         BackColor       =   &H00FF8080&
         Caption         =   "SUBMIT"
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "SUBMIT'S YOUR QUERY"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton NEW 
         BackColor       =   &H00FF8080&
         Caption         =   "CLEAR"
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "RAIS A NEW QUERY"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox txtq 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   3360
         TabIndex        =   3
         ToolTipText     =   "ENTER YOUR QUESTION"
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox txtm 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         ToolTipText     =   "ENTER ONLY NUMBERS"
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox txtn 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         ToolTipText     =   "ENTER YOUR NAME"
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"confrm.frx":1E6CF
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1035
         Left            =   360
         TabIndex        =   16
         Top             =   2520
         Width           =   2925
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUESTION                        :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   13
         Top             =   2040
         Width           =   2025
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "MOBILE NUMBER           :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   12
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "YOUR FULL NAME          :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   11
         Top             =   1080
         Width           =   2010
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "* FILL OUT THE FORM HERE.! WE WILL COME IN TOUCH WITH YOU AS SOON AS POSSIBLE.. THROUGH MOBILE CALL......."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   435
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   6765
      End
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
      Left            =   7320
      TabIndex        =   18
      Top             =   360
      Width           =   105
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DATE:"
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
      Left            =   6600
      TabIndex        =   17
      Top             =   360
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   3975
      Index           =   1
      Left            =   -120
      Picture         =   "confrm.frx":1E75F
      Stretch         =   -1  'True
      Top             =   -1200
      Width           =   3840
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
      Index           =   2
      Left            =   6600
      TabIndex        =   15
      Top             =   7440
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
      Left            =   5400
      TabIndex        =   14
      Top             =   7440
      Width           =   1080
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
      Left            =   6720
      TabIndex        =   8
      Top             =   1320
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
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   3600
      TabIndex        =   7
      Top             =   1800
      Width           =   3615
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
      ForeColor       =   &H00FF80FF&
      Height          =   585
      Left            =   3735
      TabIndex        =   0
      Top             =   1320
      Width           =   2865
   End
End
Attribute VB_Name = "connfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EXIT_Click()
Me.Hide
End Sub
Private Sub Form_Load()
connectdb
date = DateTime.Now
End Sub
Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub
Private Sub NEW_Click()
txtn = " "
txtm = " "
txtq = " "
End Sub
Private Sub SUBMIT_Click()
If txtn.Text = "" Then
MsgBox "Your Name Con't be Blank .", vbInformation
txtn.SetFocus
Exit Sub
End If

If txtm.Text = "" Then
MsgBox "Mobile No Con't be Blank .", vbInformation
txtm.SetFocus
Exit Sub
End If

If txtq = "" Then
MsgBox "Query Con't be Blank .", vbInformation
txtq.SetFocus
Exit Sub
End If

Set rs = con.Execute("select * from query where MOBILENO='" + txtm.Text + " '")
If (Not rs.EOF) Then
MsgBox "Your Last Question Is Pending 'You Are not Allowed to Submit Another Question'", vbCritical, "Warning"
txtn = ""
txtm = ""
txtq = ""
Welcomee.Show
date = DateTime.Now
Else

con.Execute ("insert into query values('" + txtn.Text + "', " + txtm.Text + ",'" + txtq.Text + "')")
MsgBox "Your Query Recorded Successfully.....", vbInformation, "Warning"
txtn = ""
txtm = ""
txtq = ""
End If
End Sub

