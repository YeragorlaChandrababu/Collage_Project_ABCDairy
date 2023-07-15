VERSION 5.00
Begin VB.Form affrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ADD FARMER"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "affrm.frx":0000
   ScaleHeight     =   6600
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "                                                     ADD NEW FARMER DETAILS                                               "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   4215
      Left            =   1680
      TabIndex        =   12
      Top             =   1800
      Width           =   7335
      Begin VB.ComboBox TXTDU 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         TabIndex        =   21
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   " *Check Farmer USER NAME Availability"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtu 
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
         Left            =   3240
         TabIndex        =   0
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox txtp 
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
         Left            =   3240
         TabIndex        =   2
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtn 
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
         Left            =   3240
         TabIndex        =   3
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtc 
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
         Left            =   3240
         TabIndex        =   4
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txta 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3240
         TabIndex        =   5
         Top             =   2400
         Width           =   2655
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
         Left            =   1080
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton RESET 
         BackColor       =   &H00FF8080&
         Caption         =   "RESET"
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
         Left            =   3000
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton BACK 
         BackColor       =   &H000000FF&
         Caption         =   "BACK"
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
         Left            =   4920
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ASSIGNED DEALER ID :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   720
         TabIndex        =   20
         Top             =   960
         Width           =   2070
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FARMER USER NAME :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         TabIndex        =   17
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NAME  :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2040
         TabIndex        =   16
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT NO  :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1320
         TabIndex        =   15
         Top             =   2040
         Width           =   1410
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS  :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   14
         Top             =   2400
         Width           =   1035
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1560
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
      End
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
      Left            =   4440
      TabIndex        =   23
      Top             =   6240
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
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   0
      Left            =   6720
      TabIndex        =   22
      Top             =   840
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   3975
      Index           =   1
      Left            =   -600
      Picture         =   "affrm.frx":D542
      Stretch         =   -1  'True
      Top             =   -1320
      Width           =   4440
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
      Index           =   0
      Left            =   3720
      TabIndex        =   19
      Top             =   840
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
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   0
      Left            =   3480
      TabIndex        =   18
      Top             =   1440
      Width           =   3615
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
      Left            =   7560
      TabIndex        =   11
      Top             =   240
      Width           =   105
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
      Left            =   6840
      TabIndex        =   10
      Top             =   240
      Width           =   660
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
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   5640
      TabIndex        =   9
      Top             =   6240
      Width           =   3375
   End
End
Attribute VB_Name = "affrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set rs = con.Execute("select * from Formerdb where uname='" + txtu.Text + "' ")
If (Not rs.EOF) Then
MsgBox "Try Another USER NAME", vbInformation, "USER NAME"
txtu.Text = ""
Else
MsgBox "You Can Proceed With Your USER NAME", vbInformation, "Warning"
End If
End Sub

Private Sub Form_Load()
date = DateTime.Now
connectdb
Set rs = con.Execute("select * from DEALERDB")
While (Not rs.EOF)
TXTDU.AddItem rs(0)
rs.MoveNext
Wend
rs.Close
End Sub
Private Sub BACK_Click()
affrm.Hide
End Sub
Private Sub RESET_Click()
txtu = ""
txtp = ""
TXTDU = ""
txtn = ""
txta = ""
txtc = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub
Private Sub SUBMIT_Click()
If txtu = "" Then
MsgBox "Enter Former User Name.", vbInformation
txtu.SetFocus
Exit Sub
End If

If txtn.Text = "" Then
MsgBox "Enter Name.", vbInformation
txtn.SetFocus
Exit Sub
End If

If txtp.Text = "" Then
MsgBox "Enter Your Password.", vbInformation
txtp.SetFocus
Exit Sub
End If

If txtc.Text = "" Then
MsgBox "Enter Contact Number.", vbInformation
txtc.SetFocus
Exit Sub
End If

If txta.Text = "" Then
MsgBox "Enter Address.", vbInformation
txta.SetFocus
Exit Sub
End If

If TXTDU.Text = "" Then
MsgBox "Enter Dealer User Name.", vbInformation
TXTDU.SetFocus
Exit Sub
End If

con.Execute ("insert into Formerdb values('" + txtu.Text + "','" + txtn.Text + "','" + txtc.Text + "', '" + txta.Text + "','" + txtp.Text + "','" + TXTDU.Text + "')")
MsgBox ("Former Details Saved Successfully..!")
txtu = ""
txtp = ""
TXTDU = ""
txtn = ""
txta = ""
txtc = ""
End Sub
Private Sub txtu_Change()
If KeyAscii = 13 Then
txtu.Text = UCase(txtnumber.Text)
txtu.SetFocus
End If
End Sub

